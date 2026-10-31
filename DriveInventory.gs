/**
 * Google Drive Folder Inventory — 1.0.0
 * Install in a NEW, spreadsheet-bound Apps Script project with appsscript.json.
 * Run setupInventory once, then use the Drive Inventory menu in the spreadsheet.
 *
 * Reads Drive METADATA ONLY. Includes every file type; does not download files,
 * parse PDF metadata or MP3 tags, change source files, or follow shortcuts.
 * All output tabs are generated. Keep personal annotations on a separate tab,
 * keyed by File ID. One spreadsheet per collection is recommended.
 *
 * A Drive page, its new queue entries, and the checkpoint are committed in one
 * atomic Sheets batchUpdate. After an uncertain write, the next invocation reads
 * the committed checkpoint instead of replaying an append. An expired page token
 * rolls back only the current folder, whose children have not yet been scanned.
 */

const DI = Object.freeze({
  VERSION: '1.0.0',
  APP: 'drive-folder-inventory',
  FOLDER_MIME: 'application/vnd.google-apps.folder',
  SHORTCUT_MIME: 'application/vnd.google-apps.shortcut',
  PAGE_SIZE: 200,
  WORK_MS: 180000,
  MIN_STEP_MS: 1400,
  TRIGGER_MINUTES: 5,
  // Reserve a conservative budget for background workers, leaving headroom for
  // startup, status checks, and other scripts on a personal Google account.
  AUTO_BUDGET_MS: 45 * 60 * 1000,
  AUTO_RESERVATION_MS: 4 * 60 * 1000,
  // Current documented Sheets capacity is 20 million cells; leave some room.
  MAX_CELLS: 19000000,
  MAX_TEXT: 48000,
  EXCLUDED_FOLDER_IDS: [],
  STATE_TITLE: '_DI_State',
  STATUS_TITLE: 'Status',
  WORK_FILES_TITLE: '_DI_Work_Inventory',
  WORK_FOLDERS_TITLE: '_DI_Work_Folders',
  FILES_TITLE: 'Inventory',
  FOLDERS_TITLE: 'Folders'
});

const DI_FILE_HEADERS = [
  'Filename', 'Type', 'Extension', 'Category', 'Folder path',
  'Drive description', 'Size (MiB)', 'Drive modified (UTC)',
  'Drive created (UTC)', 'Starred', 'Drive owners', 'File link',
  'Folder link', 'Relative file path', 'MIME type', 'Size (bytes)',
  'Original filename', 'Last modified by', 'Shared',
  'Viewed by me (UTC)', 'MD5 checksum', 'SHA-256 checksum', 'File ID',
  'Parent folder ID', 'Shared drive ID', 'Resource key',
  'Shortcut target ID', 'Shortcut target type', 'Shortcut target link',
  'Storage used (bytes)', 'Version', 'Additional metadata (JSON)'
];

const DI_FOLDER_HEADERS = [
  'Folder', 'Folder path', 'Category', 'Depth', 'Direct files',
  'Direct subfolders', 'Direct file size (MiB)', 'Drive created (UTC)',
  'Drive modified (UTC)', 'Drive description', 'Folder link', 'Folder ID',
  'Parent folder ID', 'Resource key', 'Shared drive ID', 'Scan status'
];

const DI_FILE_FIELDS = [
  'id', 'name', 'mimeType', 'parents', 'driveId', 'webViewLink',
  'description', 'size', 'quotaBytesUsed', 'createdTime', 'modifiedTime',
  'viewedByMeTime', 'sharedWithMeTime', 'originalFilename',
  'fileExtension', 'fullFileExtension', 'md5Checksum', 'sha1Checksum',
  'sha256Checksum', 'starred', 'shared', 'ownedByMe',
  'owners(displayName,emailAddress)',
  'lastModifyingUser(displayName,emailAddress)', 'version', 'headRevisionId',
  'resourceKey', 'shortcutDetails(targetId,targetMimeType,targetResourceKey)',
  'imageMediaMetadata', 'videoMediaMetadata', 'properties', 'appProperties',
  'capabilities(canListChildren)', 'trashed'
].join(',');

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Drive Inventory')
    .addItem('Choose root folder and scan…', 'chooseRootFolder')
    .addItem('Refresh — start a new scan', 'refreshInventory')
    .addItem('Resume / run next batch', 'resumeInventory')
    .addItem('Pause scanning', 'pauseInventory')
    .addSeparator()
    .addItem('Show status', 'showInventoryStatus')
    .addItem('Show all metadata columns', 'showAllMetadata')
    .addItem('Show compact columns', 'compactView')
    .addToUi();
}

function setupInventory() {
  ensureInfrastructure_();
  onOpen();
  SpreadsheetApp.getUi().alert('Drive Inventory is ready',
    'Use Drive Inventory → Choose root folder and scan. Paste a folder URL or ID. ' +
    'The script reads file metadata and writes generated tabs in this spreadsheet.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function chooseRootFolder() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.prompt('Choose a Google Drive root folder',
    'Paste a folder URL or folder ID. This starts a fresh scan; any previous ' +
    'completed inventory stays available until the new scan finishes.', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const target = parseFolderInput_(answer.getResponseText());
  const root = getDriveFile_(target.id, target.resourceKey);
  if (root.mimeType !== DI.FOLDER_MIME || root.trashed) {
    throw new Error('Choose a non-trashed folder itself, not a file or a shortcut.');
  }
  root.resourceKey = root.resourceKey || target.resourceKey || '';
  if (root.capabilities && root.capabilities.canListChildren === false) {
    throw new Error('This account cannot list the selected folder.');
  }
  const infra = ensureInfrastructure_();
  startInventory_(infra, root);
  inventoryWorker();
}

function refreshInventory() {
  const infra = ensureInfrastructure_();
  const target = readState_(infra).root;
  if (!target) return chooseRootFolder();
  const root = getDriveFile_(target.id, target.resourceKey);
  if (root.mimeType !== DI.FOLDER_MIME || root.trashed) {
    throw new Error('The configured root is no longer an accessible, non-trashed folder.');
  }
  root.resourceKey = root.resourceKey || target.resourceKey || '';
  startInventory_(infra, root);
  inventoryWorker();
}

function resumeInventory() {
  const infra = ensureInfrastructure_();
  withLock_(function () {
    const s = readState_(infra);
    if (s.phase === 'IDLE' || s.phase === 'COMPLETE') {
      throw new Error('Use Choose root folder or Refresh to start a scan.');
    }
    s.phase = s.resumePhase || s.phase;
    if (s.phase === 'PAUSED') s.phase = 'SCANNING';
    delete s.resumePhase;
    s.waitUntil = 0;
    s.error = '';
    s.updated = new Date().toISOString();
    commit_(infra, s, []);
    ensureWorkerTrigger_();
  });
  inventoryWorker();
}

function pauseInventory() {
  const infra = ensureInfrastructure_();
  withLock_(function () {
    const s = readState_(infra);
    if (s.phase !== 'COMPLETE' && s.phase !== 'IDLE' && s.phase !== 'PAUSED') {
      s.resumePhase = s.phase;
      s.phase = 'PAUSED';
      s.error = 'Paused by you. Use Resume / run next batch to continue.';
      s.updated = new Date().toISOString();
      commit_(infra, s, []);
    }
    removeWorkerTriggers_();
  });
}

function showInventoryStatus() {
  const infra = ensureInfrastructure_();
  const ss = SpreadsheetApp.openById(infra.spreadsheetId);
  ss.setActiveSheet(ss.getSheetById(infra.statusId));
}

function showAllMetadata() { setMetadataVisibility_(false); }
function compactView() { setMetadataVisibility_(true); }

/** Background handler. Runs without an active browser tab or active sheet. */
function inventoryWorker(event) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  let infra;
  const started = Date.now();
  try {
    infra = existingInfrastructure_();
    let s = readState_(infra);
    if (s.phase === 'IDLE' || s.phase === 'COMPLETE' || s.phase === 'PAUSED') {
      removeWorkerTriggers_();
      return;
    }
    const automatic = !!(event && event.triggerUid);

    // Do not impose an additional artificial runtime budget.
    // Google Apps Script enforces its own execution and daily quotas.
    if (s.waitUntil) {
      s.waitUntil = 0;
      s.updated = new Date().toISOString();
      commit_(infra, s, []);
    }

    let layout = sheetLayout_(infra.spreadsheetId);
    let seenFiles = readIds_(infra.spreadsheetId, s.filesId, 22, s.files, layout);
    let seenFolders = readIds_(infra.spreadsheetId, s.foldersId, 11, s.folders, layout);

    let previousStep = 0;
    while (Date.now() - started < DI.WORK_MS) {
      // Each step normally makes one Sheets read and one write. Pace small
      // folders as well as large pages to leave headroom under per-minute quotas.
      if (previousStep) {
        const remaining = DI.MIN_STEP_MS - (Date.now() - previousStep);
        if (remaining > 0) Utilities.sleep(remaining);
      }
      previousStep = Date.now();
      if (s.phase === 'SCANNING' && s.cursor >= s.folders) {
        s.phase = 'FINALIZING';
        s.updated = new Date().toISOString();
        commit_(infra, s, []);
      }
      if (s.phase === 'FINALIZING') {
        publishInventory_(infra, s, layout);
        removeWorkerTriggers_();
        return;
      }
      if (s.phase !== 'SCANNING') return;

      const folder = readFolder_(infra.spreadsheetId, s, layout);
      if (folder.status === 'Excluded') {
        s.cursor++;
        s.folderStart = null;
        s.pageToken = '';
        s.current = { files: 0, folders: 0, bytes: 0 };
        s.updated = new Date().toISOString();
        commit_(infra, s, []);
        continue;
      }
      let page;
      try {
        page = listDriveChildren_(folder, s.pageToken);
      } catch (error) {
        if (s.pageToken && isPageTokenError_(error)) {
          rollbackFolder_(infra, s, layout);
          return;
        }
        throw error;
      }
      if (page.incompleteSearch) {
        throw new Error('Drive reported an incomplete search for ' + folder.path +
          '. No partial result has replaced the completed inventory.');
      }
      const planned = planPage_(s, folder, page, seenFiles, seenFolders);
      const requests = [];
      growSheet_(requests, layout, s.filesId, planned.state.files + 1);
      growSheet_(requests, layout, s.foldersId, planned.state.folders + 1);
      if (planned.fileRows.length) {
        requests.push(writeRows_(s.filesId, s.files + 1, 0, planned.fileRows));
      }
      if (planned.folderRows.length) {
        requests.push(writeRows_(s.foldersId, s.folders + 1, 0, planned.folderRows));
      }
      requests.push(writeRows_(s.foldersId, s.cursor + 1, 4, [[
        planned.state.current.files, planned.state.current.folders,
        planned.state.current.bytes / 1048576
      ]]));
      requests.push(writeRows_(s.foldersId, s.cursor + 1, 15,
        [[page.nextPageToken ? 'Scanning' : 'Complete']]));
      if (!page.nextPageToken) {
        planned.state.cursor++;
        planned.state.current = { files: 0, folders: 0, bytes: 0 };
        planned.state.folderStart = null;
        planned.state.tokenRestarts = 0;
      }
      checkCapacity_(layout, requests);
      commit_(infra, planned.state, requests);
      s = planned.state;
      planned.fileIds.forEach(function (id) { seenFiles.add(id); });
      planned.folderIds.forEach(function (id) { seenFolders.add(id); });
      applyLayoutChanges_(layout, requests);
    }
  } catch (error) {
    // A request may have committed before its response was lost. Reload state;
    // never write a stale checkpoint or automatically replay an uncertain batch.
    const retryable = isTransient_(error);
    if (infra) {
      try {
        const latest = readState_(infra);
        if (latest.phase !== 'COMPLETE' && latest.phase !== 'IDLE') {
          if (!retryable) {
            latest.resumePhase = latest.phase === 'PAUSED'
              ? latest.resumePhase : latest.phase;
            latest.phase = 'PAUSED';
          }
          latest.error = (retryable ? 'Temporary service error; the next background run will retry. ' : '') +
            String(error.message || error).slice(0, 3500);
          latest.updated = new Date().toISOString();
          commit_(infra, latest, []);
          if (!retryable) removeWorkerTriggers_();
        } else {
          removeWorkerTriggers_();
        }
      } catch (statusError) {
        console.error('Could not record pause: ' + statusError);
        // Leave the trigger available to recover from a temporary Sheets outage.
      }
    }
    console.error(error && error.stack ? error.stack : String(error));
    if (!(event && event.triggerUid) && !retryable) throw error;
  } finally {
    lock.releaseLock();
  }
}

function startInventory_(infra, root) {
  withLock_(function () {
    const old = readState_(infra);
    const layout = sheetLayout_(infra.spreadsheetId);
    const requests = [];
    [old.filesId, old.foldersId].filter(function (id) {
      return id && (!old.published ||
        (id !== old.published.filesId && id !== old.published.foldersId));
    }).forEach(function (id) {
      if (layout[id]) requests.push({ deleteSheet: { sheetId: id } });
    });
    const used = new Set(Object.keys(layout).map(Number));
    const filesId = newSheetId_(used);
    const foldersId = newSheetId_(used);
    assertOutputNames_(layout, old);
    requests.push(addSheet_(filesId, DI.WORK_FILES_TITLE, DI_FILE_HEADERS.length, true));
    requests.push(addSheet_(foldersId, DI.WORK_FOLDERS_TITLE, DI_FOLDER_HEADERS.length, true));
    requests.push(writeRows_(filesId, 0, 0, [DI_FILE_HEADERS]));
    requests.push(writeRows_(foldersId, 0, 0, [DI_FOLDER_HEADERS]));
    const rootFolder = {
      id: root.id, name: root.name, path: '.', category: '(root)', depth: 0,
      parentId: (root.parents || [])[0] || '', resourceKey: root.resourceKey || '',
      driveId: root.driveId || '', url: driveUrl_(root), status: 'Queued'
    };
    requests.push(writeRows_(foldersId, 1, 0, [folderRow_(root, rootFolder)]));
    const now = new Date().toISOString();
    const s = {
      app: DI.APP, version: DI.VERSION, phase: 'SCANNING',
      spreadsheetId: infra.spreadsheetId, stateId: infra.stateId, statusId: infra.statusId,
      filesId: filesId, foldersId: foldersId,
      root: { id: root.id, name: root.name, url: driveUrl_(root),
        resourceKey: root.resourceKey || '', driveId: root.driveId || '' },
      files: 0, folders: 1, cursor: 0, bytes: 0, types: {},
      shortcuts: 0, excluded: 0, repeatedIds: 0, pageToken: '',
      current: { files: 0, folders: 0, bytes: 0 }, folderStart: null,
      tokenRestarts: 0, started: now, updated: now, completed: '', error: '',
      published: old.published || null,
      autoWindow: old.autoWindow || Date.now(), autoReservedMs: old.autoReservedMs || 0,
      waitUntil: 0
    };
    checkCapacity_(layout, requests);
    commit_(infra, s, requests);
    ensureWorkerTrigger_();
  });
}

/** Pure planning: no service calls and no mutation of the committed state. */
function planPage_(state, folder, page, seenFiles, seenFolders) {
  const s = JSON.parse(JSON.stringify(state));
  if (!s.folderStart) {
    s.folderStart = {
      files: s.files, folders: s.folders, bytes: s.bytes,
      types: Object.assign({}, s.types), shortcuts: s.shortcuts,
      excluded: s.excluded, repeatedIds: s.repeatedIds
    };
  }
  const fileRows = [], folderRows = [], fileIds = [], folderIds = [];
  const localFiles = new Set(), localFolders = new Set();
  (page.files || []).forEach(function (file) {
    if (!file.id || file.trashed) return;
    if (file.mimeType === DI.FOLDER_MIME) {
      if (seenFolders.has(file.id) || localFolders.has(file.id)) {
        s.repeatedIds++;
        return;
      }
      localFolders.add(file.id);
      folderIds.push(file.id);
      const child = {
        id: file.id, name: file.name,
        path: joinPath_(folder.path, file.name),
        category: folder.depth === 0 ? file.name : folder.category,
        depth: folder.depth + 1, parentId: folder.id,
        resourceKey: file.resourceKey || '', driveId: file.driveId || folder.driveId || '',
        url: driveUrl_(file),
        status: DI.EXCLUDED_FOLDER_IDS.indexOf(file.id) >= 0 ? 'Excluded' : 'Queued'
      };
      folderRows.push(folderRow_(file, child));
      s.folders++;
      s.current.folders++;
      if (child.status === 'Excluded') s.excluded++;
    } else {
      if (seenFiles.has(file.id) || localFiles.has(file.id)) {
        s.repeatedIds++;
        return;
      }
      localFiles.add(file.id);
      fileIds.push(file.id);
      fileRows.push(fileRow_(file, folder));
      s.files++;
      s.current.files++;
      const bytes = numeric_(file.size);
      if (bytes !== '') { s.bytes += bytes; s.current.bytes += bytes; }
      const type = fileType_(file);
      s.types[type] = (s.types[type] || 0) + 1;
      if (file.mimeType === DI.SHORTCUT_MIME) s.shortcuts++;
    }
  });
  s.pageToken = page.nextPageToken || '';
  s.error = '';
  s.updated = new Date().toISOString();
  return { state: s, fileRows: fileRows, folderRows: folderRows,
    fileIds: fileIds, folderIds: folderIds };
}

function rollbackFolder_(infra, s, layout) {
  if (!s.folderStart || (s.tokenRestarts || 0) >= 3) {
    throw new Error('Drive repeatedly rejected the pagination token. Use Refresh to ' +
      'start a new scan. The previous completed inventory is unchanged.');
  }
  const base = s.folderStart;
  const requests = [];
  // The active folder has not finished, so its newly appended descendants have
  // never been visited; these tail rows can be removed without losing older work.
  if (s.files > base.files) requests.push(clearRows_(s.filesId, base.files + 1,
    s.files + 1, DI_FILE_HEADERS.length));
  if (s.folders > base.folders) requests.push(clearRows_(s.foldersId, base.folders + 1,
    s.folders + 1, DI_FOLDER_HEADERS.length));
  Object.keys(base).forEach(function (key) { s[key] = base[key]; });
  s.pageToken = '';
  s.current = { files: 0, folders: 0, bytes: 0 };
  s.tokenRestarts = (s.tokenRestarts || 0) + 1;
  s.updated = new Date().toISOString();
  requests.push(writeRows_(s.foldersId, s.cursor + 1, 4, [[0, 0, 0]]));
  requests.push(writeRows_(s.foldersId, s.cursor + 1, 15, [['Restarting expired page token']]));
  commit_(infra, s, requests);
}

function publishInventory_(infra, s, layout) {
  const requests = [];
  // Delete only previously published sheets whose IDs were recorded by us.
  if (s.published) {
    [s.published.filesId, s.published.foldersId].forEach(function (id) {
      if (id !== s.filesId && id !== s.foldersId && layout[id]) {
        requests.push({ deleteSheet: { sheetId: id } });
      }
    });
  }
  requests.push({ updateSheetProperties: { properties: {
    sheetId: s.filesId, title: DI.FILES_TITLE, hidden: false, index: 0,
    gridProperties: { rowCount: Math.max(2, s.files + 1), frozenRowCount: 1, frozenColumnCount: 1 }
  }, fields: 'title,hidden,index,gridProperties.rowCount,gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });
  requests.push({ updateSheetProperties: { properties: {
    sheetId: s.foldersId, title: DI.FOLDERS_TITLE, hidden: false, index: 1,
    gridProperties: { rowCount: Math.max(2, s.folders + 1), frozenRowCount: 1, frozenColumnCount: 1 }
  }, fields: 'title,hidden,index,gridProperties.rowCount,gridProperties.frozenRowCount,gridProperties.frozenColumnCount' } });
  requests.push({ updateSheetProperties: { properties: { sheetId: infra.statusId, index: 2 }, fields: 'index' } });
  requests.push.apply(requests, formatTableRequests_(s.filesId, s.files + 1, DI_FILE_HEADERS.length, false));
  requests.push.apply(requests, formatTableRequests_(s.foldersId, s.folders + 1, DI_FOLDER_HEADERS.length, true));
  if (s.files > 1) requests.push({ sortRange: {
    range: range_(s.filesId, 1, s.files + 1, 0, DI_FILE_HEADERS.length),
    sortSpecs: [{ dimensionIndex: 4, sortOrder: 'ASCENDING' },
      { dimensionIndex: 0, sortOrder: 'ASCENDING' }]
  } });
  if (s.folders > 1) requests.push({ sortRange: {
    range: range_(s.foldersId, 1, s.folders + 1, 0, DI_FOLDER_HEADERS.length),
    sortSpecs: [{ dimensionIndex: 1, sortOrder: 'ASCENDING' }]
  } });
  s.phase = 'COMPLETE';
  s.completed = new Date().toISOString();
  s.updated = s.completed;
  s.error = '';
  s.published = { filesId: s.filesId, foldersId: s.foldersId,
    rootName: s.root.name, rootId: s.root.id, completed: s.completed,
    files: s.files, folders: s.folders, excluded: s.excluded };
  commit_(infra, s, requests);
}

function fileRow_(file, folder) {
  const bytes = numeric_(file.size);
  const shortcut = file.shortcutDetails || {};
  const extra = {};
  ['sha1Checksum', 'headRevisionId', 'ownedByMe', 'sharedWithMeTime',
    'fullFileExtension', 'imageMediaMetadata', 'videoMediaMetadata',
    'properties', 'appProperties'].forEach(function (key) {
    if (file[key] !== undefined) extra[key] = file[key];
  });
  const targetUrl = shortcut.targetId ? driveUrl_({ id: shortcut.targetId,
    mimeType: shortcut.targetMimeType, resourceKey: shortcut.targetResourceKey }) : '';
  const url = driveUrl_(file);
  return [
    linked_(file.name || '', url), fileType_(file), extension_(file),
    folder.category, folder.path, file.description || '',
    bytes === '' ? '' : bytes / 1048576,
    dateCell_(file.modifiedTime), dateCell_(file.createdTime),
    boolOrBlank_(file.starred), people_(file.owners), linked_(url, url),
    linked_(folder.url, folder.url), joinPath_(folder.path, file.name || ''),
    file.mimeType || '', bytes, file.originalFilename || '',
    people_(file.lastModifyingUser ? [file.lastModifyingUser] : []),
    boolOrBlank_(file.shared), dateCell_(file.viewedByMeTime),
    file.md5Checksum || '', file.sha256Checksum || '', file.id,
    folder.id, file.driveId || '', file.resourceKey || '',
    shortcut.targetId || '', shortcut.targetMimeType || '', linked_(targetUrl, targetUrl),
    numeric_(file.quotaBytesUsed), file.version || '',
    Object.keys(extra).length ? JSON.stringify(extra) : ''
  ];
}

function folderRow_(file, folder) {
  const count = folder.status === 'Excluded' ? '' : 0;
  return [linked_(file.name || '', folder.url), folder.path, folder.category,
    folder.depth, count, count, count, dateCell_(file.createdTime), dateCell_(file.modifiedTime),
    file.description || '', linked_(folder.url, folder.url), file.id,
    folder.parentId || '', folder.resourceKey || '', folder.driveId || '', folder.status];
}

function readFolder_(spreadsheetId, s, layout) {
  const name = quoteSheet_(layout[s.foldersId].title);
  const row = (Sheets.Spreadsheets.Values.get(spreadsheetId,
    name + '!A' + (s.cursor + 2) + ':P' + (s.cursor + 2),
    { valueRenderOption: 'UNFORMATTED_VALUE' }).values || [])[0];
  if (!row || !row[11]) throw new Error('The scan queue was changed or removed. Use Refresh.');
  return { name: row[0], path: row[1], category: row[2], depth: Number(row[3]),
    url: row[10], id: row[11], parentId: row[12] || '', resourceKey: row[13] || '',
    driveId: row[14] || '', status: row[15] || '' };
}

function readIds_(spreadsheetId, sheetId, column, count, layout) {
  if (!count) return new Set();
  if (!layout[sheetId]) throw new Error('A generated scan sheet was removed. Use Refresh.');
  const letter = columnName_(column);
  const values = Sheets.Spreadsheets.Values.get(spreadsheetId,
    quoteSheet_(layout[sheetId].title) + '!' + letter + '2:' + letter + (count + 1),
    { valueRenderOption: 'UNFORMATTED_VALUE' }).values || [];
  return new Set(values.map(function (row) { return String(row[0] || ''); }));
}

function getDriveFile_(id, resourceKey) {
  const parameters = { supportsAllDrives: true, fields: DI_FILE_FIELDS };
  return driveRead_('files/' + encodeURIComponent(id), parameters, id, resourceKey,
    function () { return Drive.Files.get(id, parameters); });
}

function listDriveChildren_(folder, token) {
  const parameters = {
    q: "'" + folder.id.replace(/'/g, "\\'") + "' in parents and trashed = false",
    spaces: 'drive', pageSize: DI.PAGE_SIZE,
    fields: 'nextPageToken,incompleteSearch,files(' + DI_FILE_FIELDS + ')',
    supportsAllDrives: true, includeItemsFromAllDrives: true,
    corpora: folder.driveId ? 'drive' : 'user'
  };
  if (folder.driveId) parameters.driveId = folder.driveId;
  if (token) parameters.pageToken = token;
  return driveRead_('files', parameters, folder.id, folder.resourceKey,
    function () { return Drive.Files.list(parameters); });
}

function driveRead_(path, parameters, id, resourceKey, advancedCall) {
  let lastError;
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      if (!resourceKey) return advancedCall();
      // The Advanced Drive wrapper does not document this custom header. The
      // REST fallback preserves link-only folder access using the same scope.
      const query = Object.keys(parameters).map(function (key) {
        return encodeURIComponent(key) + '=' + encodeURIComponent(parameters[key]);
      }).join('&');
      const response = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/' + path + '?' + query, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
          'X-Goog-Drive-Resource-Keys': id + '/' + resourceKey },
        muteHttpExceptions: true
      });
      const status = response.getResponseCode();
      const body = JSON.parse(response.getContentText());
      if (status >= 200 && status < 300) return body;
      const error = new Error('Drive HTTP ' + status + ': ' +
        (body.error ? body.error.message : 'request failed'));
      error.status = status;
      error.reasons = body.error && body.error.errors
        ? body.error.errors.map(function (item) { return item.reason; }).join(' ') : '';
      throw error;
    } catch (error) {
      lastError = error;
      if (!isTransient_(error) || attempt === 2) throw error;
      Utilities.sleep((Math.pow(2, attempt) * 1000) + Math.floor(Math.random() * 500));
    }
  }
  throw lastError;
}

function isTransient_(error) {
  const text = String(error.message || error) + ' ' + (error.reasons || '');
  return [429, 500, 502, 503, 504].indexOf(error.status) >= 0 ||
    /rateLimitExceeded|userRateLimitExceeded|Rate Limit Exceeded|Too Many Requests|quota.*per minute|too many times in a short time|backendError|Internal Error|Service Unavailable|timed out|timeout/i.test(text);
}

function isPageTokenError_(error) {
  return /page.?token|invalid.*token|token.*(invalid|expired)|invalid value/i.test(String(error.message || error));
}

function ensureInfrastructure_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Install and run this script from Extensions → Apps Script in a Google Sheet.');
  const spreadsheetId = ss.getId();
  const properties = PropertiesService.getScriptProperties();
  const stored = properties.getProperty('DI_INFRA');
  if (stored) {
    const infra = JSON.parse(stored);
    if (infra.spreadsheetId !== spreadsheetId) {
      throw new Error('This project is bound to a different spreadsheet. Install the files in a fresh bound project.');
    }
    readState_(infra);
    return infra;
  }
  const layout = sheetLayout_(spreadsheetId);
  const stateSheet = Object.values(layout).find(function (p) { return p.title === DI.STATE_TITLE; });
  if (stateSheet) {
    const value = Sheets.Spreadsheets.Values.get(spreadsheetId,
      quoteSheet_(DI.STATE_TITLE) + '!A1').values;
    let recovered;
    try { recovered = JSON.parse(value[0][0]); } catch (error) { /* handled below */ }
    if (recovered && recovered.app === DI.APP && recovered.spreadsheetId === spreadsheetId && layout[recovered.statusId]) {
      const infra = { spreadsheetId: spreadsheetId, stateId: stateSheet.id, statusId: recovered.statusId };
      properties.setProperty('DI_INFRA', JSON.stringify(infra));
      return infra;
    }
    throw new Error('A sheet named ' + DI.STATE_TITLE + ' already exists and is not this inventory. Use a new blank spreadsheet.');
  }
  const reserved = [DI.STATUS_TITLE, DI.FILES_TITLE, DI.FOLDERS_TITLE,
    DI.WORK_FILES_TITLE, DI.WORK_FOLDERS_TITLE];
  if (Object.values(layout).some(function (p) { return reserved.indexOf(p.title) >= 0; })) {
    throw new Error('This spreadsheet has tabs using inventory names. Use a new blank spreadsheet to avoid changing existing work.');
  }
  const used = new Set(Object.keys(layout).map(Number));
  const infra = { spreadsheetId: spreadsheetId, stateId: newSheetId_(used), statusId: newSheetId_(used) };
  const s = { app: DI.APP, version: DI.VERSION, phase: 'IDLE',
    spreadsheetId: spreadsheetId, stateId: infra.stateId, statusId: infra.statusId,
    updated: new Date().toISOString(), published: null };
  const requests = [
    { addSheet: { properties: { sheetId: infra.statusId, title: DI.STATUS_TITLE,
      gridProperties: { rowCount: 60, columnCount: 2, frozenRowCount: 1 } } } },
    { addSheet: { properties: { sheetId: infra.stateId, title: DI.STATE_TITLE,
      hidden: true, gridProperties: { rowCount: 1, columnCount: 1 } } } },
    { updateDimensionProperties: { range: { sheetId: infra.statusId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 230 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId: infra.statusId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 },
      properties: { pixelSize: 820 }, fields: 'pixelSize' } },
    { repeatCell: { range: range_(infra.statusId, 0, 1, 0, 2), cell: { userEnteredFormat: {
      backgroundColor: { red: 0.12, green: 0.20, blue: 0.27 },
      textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } }
    } }, fields: 'userEnteredFormat' } }
  ];
  commit_(infra, s, requests);
  properties.setProperty('DI_INFRA', JSON.stringify(infra));
  return infra;
}

function existingInfrastructure_() {
  const value = PropertiesService.getScriptProperties().getProperty('DI_INFRA');
  if (!value) throw new Error('Run setupInventory from your spreadsheet first.');
  return JSON.parse(value);
}

function readState_(infra) {
  const values = Sheets.Spreadsheets.Values.get(infra.spreadsheetId,
    quoteSheet_(DI.STATE_TITLE) + '!A1', { valueRenderOption: 'UNFORMATTED_VALUE' }).values;
  if (!values || !values[0] || !values[0][0]) throw new Error('The inventory checkpoint is missing.');
  const s = JSON.parse(values[0][0]);
  if (s.app !== DI.APP || s.spreadsheetId !== infra.spreadsheetId || s.stateId !== infra.stateId) {
    throw new Error('The inventory checkpoint does not belong to this spreadsheet.');
  }
  return s;
}

function commit_(infra, s, requests) {
  const all = requests.slice();
  all.push(writeRows_(infra.stateId, 0, 0, [[JSON.stringify(s)]]));
  all.push({ updateCells: { range: range_(infra.statusId, 0, 60, 0, 2),
    rows: statusRows_(s).map(rowData_),
    fields: 'userEnteredValue,userEnteredFormat.textFormat.link,userEnteredFormat.numberFormat,note' } });
  Sheets.Spreadsheets.batchUpdate({ requests: all }, infra.spreadsheetId);
}

function statusRows_(s) {
  let phase = s.phase || 'IDLE';
  if (s.waitUntil && Date.now() < s.waitUntil && phase === 'SCANNING') phase = 'WAITING FOR AUTOMATIC RUNTIME BUDGET';
  const rows = [
    ['GOOGLE DRIVE FOLDER INVENTORY', 'Version ' + DI.VERSION],
    ['Status', phase], ['Root folder', s.root ? linked_(s.root.name, s.root.url) : 'Choose a folder from the Drive Inventory menu'],
    ['Files collected in this scan', s.files || 0], ['Folders discovered (includes root)', s.folders || 0],
    ['Folders processed', s.cursor || 0], ['File size total (MiB, where available)', (s.bytes || 0) / 1048576],
    ['Shortcuts listed', s.shortcuts || 0], ['Excluded folders', s.excluded || 0],
    ['Repeated file/folder IDs suppressed', s.repeatedIds || 0],
    ['Scan started (UTC)', s.started || ''], ['Last progress (UTC)', s.updated || ''],
    ['Scan completed (UTC)', s.completed || ''],
    ['Previous / current completed root', s.published ? s.published.rootName : 'None yet'],
    ['Last successful completion (UTC)', s.published ? s.published.completed : ''],
    ['Automatic work eligible after (UTC)', s.waitUntil ? new Date(s.waitUntil).toISOString() : ''],
    ['Message', s.error || (s.phase === 'COMPLETE' ? 'Inventory and Folders are ready.' : '')],
    ['Search', 'Use Find in sheet, or Data → Create a filter / filter views.'],
    ['Extra fields', 'Drive Inventory → Show all metadata columns. Blank means unavailable, not zero.'],
    ['Scope', 'Accessible, non-trashed descendants of the chosen root; shortcut targets are not traversed.'],
    ['Metadata', 'Drive metadata only. PDF author/page count and MP3 ID3 tags are not extracted.'],
    ['Dates and people', 'Drive dates and users are not publication/broadcast dates or authors/artists.'],
    ['Generated output', 'Refresh rebuilds Inventory and Folders. Keep personal notes elsewhere, keyed by File ID.'],
    ['Folder counts', 'Direct children only; subfolder totals are not rolled up into their parents.'],
    ['Large collections', 'A refresh needs space for both the old and new inventory. Split roots if capacity is reached.'],
    ['', ''], ['FILE TYPES', 'FILES']
  ];
  Object.keys(s.types || {}).sort().forEach(function (type) { rows.push([type, s.types[type]]); });
  return rows;
}

function sheetLayout_(spreadsheetId) {
  const metadata = Sheets.Spreadsheets.get(spreadsheetId,
    { fields: 'sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))' });
  const layout = {};
  (metadata.sheets || []).forEach(function (sheet) {
    const p = sheet.properties;
    layout[p.sheetId] = { id: p.sheetId, title: p.title,
      rows: p.gridProperties ? p.gridProperties.rowCount || 0 : 0,
      columns: p.gridProperties ? p.gridProperties.columnCount || 0 : 0 };
  });
  return layout;
}

function assertOutputNames_(layout, old) {
  const ours = new Set([old.filesId, old.foldersId,
    old.published && old.published.filesId, old.published && old.published.foldersId]);
  const names = [DI.FILES_TITLE, DI.FOLDERS_TITLE, DI.WORK_FILES_TITLE, DI.WORK_FOLDERS_TITLE];
  Object.values(layout).forEach(function (p) {
    if (names.indexOf(p.title) >= 0 && !ours.has(p.id)) {
      throw new Error('The tab "' + p.title + '" is not owned by this script. Rename it before scanning.');
    }
  });
}

function checkCapacity_(layout, requests) {
  const projected = JSON.parse(JSON.stringify(layout));
  applyLayoutChanges_(projected, requests);
  const total = Object.values(projected).reduce(function (sum, p) { return sum + p.rows * p.columns; }, 0);
  if (total > DI.MAX_CELLS) throw new Error('Spreadsheet capacity reached (' + total.toLocaleString() +
    ' allocated cells including old and new tables). The scan is paused and the previous catalog is intact. ' +
    'Use separate inventories for smaller roots. No files were silently omitted.');
}

function applyLayoutChanges_(layout, requests) {
  requests.forEach(function (request) {
    if (request.deleteSheet) delete layout[request.deleteSheet.sheetId];
    const add = request.addSheet && request.addSheet.properties;
    if (add) layout[add.sheetId] = { id: add.sheetId, title: add.title,
      rows: add.gridProperties.rowCount, columns: add.gridProperties.columnCount };
    const update = request.updateSheetProperties && request.updateSheetProperties.properties;
    if (update && layout[update.sheetId]) {
      const p = layout[update.sheetId];
      if (update.title) p.title = update.title;
      if (update.gridProperties && update.gridProperties.rowCount !== undefined) p.rows = update.gridProperties.rowCount;
      if (update.gridProperties && update.gridProperties.columnCount !== undefined) p.columns = update.gridProperties.columnCount;
    }
  });
}

function growSheet_(requests, layout, id, neededRows) {
  if (!layout[id]) throw new Error('A scan sheet was removed. Use Refresh to start again.');
  if (neededRows > layout[id].rows) requests.push({ updateSheetProperties: {
    properties: { sheetId: id, gridProperties: { rowCount: neededRows } }, fields: 'gridProperties.rowCount'
  } });
}

function addSheet_(id, title, columns, hidden) {
  return { addSheet: { properties: { sheetId: id, title: title, hidden: hidden,
    gridProperties: { rowCount: 2, columnCount: columns } } } };
}

function formatTableRequests_(id, rows, columns, folders) {
  const end = Math.max(2, rows);
  const requests = [
    { repeatCell: { range: range_(id, 0, end, 0, columns), cell: { userEnteredFormat: {
      wrapStrategy: 'CLIP', verticalAlignment: 'MIDDLE', textFormat: { fontFamily: 'Arial', fontSize: 10 }
    } }, fields: 'userEnteredFormat.wrapStrategy,userEnteredFormat.verticalAlignment,userEnteredFormat.textFormat.fontFamily,userEnteredFormat.textFormat.fontSize' } },
    { repeatCell: { range: range_(id, 0, 1, 0, columns), cell: { userEnteredFormat: {
      backgroundColor: { red: 0.12, green: 0.20, blue: 0.27 },
      textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } }
    } }, fields: 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold,userEnteredFormat.textFormat.foregroundColor' } },
    { setBasicFilter: { filter: { range: range_(id, 0, end, 0, columns) } } },
    { updateDimensionProperties: { range: { sheetId: id, dimension: 'COLUMNS', startIndex: 0, endIndex: columns },
      properties: { pixelSize: 150 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId: id, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 360 }, fields: 'pixelSize' } },
    { updateDimensionProperties: { range: { sheetId: id, dimension: 'ROWS', startIndex: 0, endIndex: end },
      properties: { pixelSize: 26 }, fields: 'pixelSize' } }
  ];
  const wide = folders ? [1, 9] : [4, 5];
  wide.forEach(function (column) {
    requests.push({ updateDimensionProperties: { range: { sheetId: id, dimension: 'COLUMNS', startIndex: column, endIndex: column + 1 },
      properties: { pixelSize: column === 5 || column === 9 ? 400 : 300 }, fields: 'pixelSize' } });
  });
  [6].forEach(function (column) {
    requests.push({ repeatCell: { range: range_(id, 1, end, column, column + 1), cell: {
      userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: '#,##0.00' } }
    }, fields: 'userEnteredFormat.numberFormat' } });
  });
  requests.push({ updateDimensionProperties: { range: { sheetId: id, dimension: 'COLUMNS', startIndex: folders ? 11 : 12, endIndex: columns },
    properties: { hiddenByUser: true }, fields: 'hiddenByUser' } });
  return requests;
}

function setMetadataVisibility_(hidden) {
  const infra = ensureInfrastructure_();
  const s = readState_(infra);
  if (!s.published) throw new Error('There is no completed inventory yet.');
  const requests = [
    { updateDimensionProperties: { range: { sheetId: s.published.filesId, dimension: 'COLUMNS', startIndex: 12, endIndex: DI_FILE_HEADERS.length },
      properties: { hiddenByUser: hidden }, fields: 'hiddenByUser' } },
    { updateDimensionProperties: { range: { sheetId: s.published.foldersId, dimension: 'COLUMNS', startIndex: 11, endIndex: DI_FOLDER_HEADERS.length },
      properties: { hiddenByUser: hidden }, fields: 'hiddenByUser' } }
  ];
  Sheets.Spreadsheets.batchUpdate({ requests: requests }, infra.spreadsheetId);
}

function ensureWorkerTrigger_() {
  const matches = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'inventoryWorker'; });
  if (!matches.length) ScriptApp.newTrigger('inventoryWorker').timeBased().everyMinutes(DI.TRIGGER_MINUTES).create();
  matches.slice(1).forEach(function (t) { ScriptApp.deleteTrigger(t); });
}

function removeWorkerTriggers_() {
  ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'inventoryWorker'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
}

function withLock_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) throw new Error('A scan batch is running. Try this command again shortly.');
  try { return callback(); } finally { lock.releaseLock(); }
}

function writeRows_(id, row, column, rows) {
  return { updateCells: { start: { sheetId: id, rowIndex: row, columnIndex: column },
    rows: rows.map(rowData_),
    fields: 'userEnteredValue,userEnteredFormat.textFormat.link,userEnteredFormat.numberFormat,note' } };
}

function clearRows_(id, start, end, columns) {
  return { updateCells: { range: range_(id, start, end, 0, columns),
    fields: 'userEnteredValue,userEnteredFormat.textFormat.link,userEnteredFormat.numberFormat,note' } };
}

function rowData_(row) { return { values: row.map(cell_) }; }

function cell_(value) {
  if (value && typeof value === 'object' && value.userEnteredValue) return value;
  if (typeof value === 'number' && Number.isFinite(value)) return { userEnteredValue: { numberValue: value } };
  if (typeof value === 'boolean') return { userEnteredValue: { boolValue: value } };
  const original = value === null || value === undefined ? '' : String(value);
  const text = truncate_(original);
  const cell = { userEnteredValue: { stringValue: text } };
  if (original.length > DI.MAX_TEXT) cell.note = 'Metadata exceeds the per-cell text budget; this value was truncated.';
  return cell;
}

function linked_(text, url) {
  const cell = cell_(text);
  if (url && /^https:\/\//i.test(url)) cell.userEnteredFormat = { textFormat: { link: { uri: url } } };
  return cell;
}

function dateCell_(value) {
  if (!value) return '';
  const ms = Date.parse(value);
  if (!Number.isFinite(ms)) return String(value);
  return { userEnteredValue: { numberValue: ms / 86400000 + 25569 },
    userEnteredFormat: { numberFormat: { type: 'DATE_TIME', pattern: 'yyyy-mm-dd hh:mm:ss' } } };
}

function numeric_(value) {
  if (value === undefined || value === null || value === '') return '';
  const result = Number(value);
  return Number.isFinite(result) ? result : '';
}

function boolOrBlank_(value) { return typeof value === 'boolean' ? value : ''; }

function extension_(file) {
  if (file.fileExtension) return String(file.fileExtension).toLowerCase();
  if ((file.mimeType || '').indexOf('application/vnd.google-apps.') === 0) return '';
  const name = file.name || '';
  const match = name.match(/\.([^./]+)$/);
  return match ? match[1].toLowerCase() : '';
}

function fileType_(file) {
  const mime = (file.mimeType || '').toLowerCase();
  const ext = extension_(file);
  if (mime === DI.SHORTCUT_MIME) return 'Shortcut';
  if (mime === 'application/pdf' || ext === 'pdf') return 'PDF';
  if (/^audio\//.test(mime) || /^(mp3|m4a|aac|flac|wav|ogg|opus|aiff|wma)$/.test(ext)) return 'Audio';
  if (/^image\//.test(mime)) return 'Image';
  if (/^video\//.test(mime)) return 'Video';
  if (/^(zip|7z|rar|tar|gz|bz2|xz)$/.test(ext)) return 'Archive';
  if (/^(epub|mobi|azw|azw3|fb2)$/.test(ext)) return 'E-book';
  if (mime === 'application/vnd.google-apps.document') return 'Google Doc';
  if (mime === 'application/vnd.google-apps.spreadsheet') return 'Google Sheet';
  if (mime === 'application/vnd.google-apps.presentation') return 'Google Slides';
  if (mime.indexOf('application/vnd.google-apps.') === 0) return 'Google file';
  if (/^text\//.test(mime)) return 'Text';
  return 'Other';
}

function people_(people) {
  return (people || []).map(function (person) {
    if (person.displayName && person.emailAddress) return person.displayName + ' <' + person.emailAddress + '>';
    return person.displayName || person.emailAddress || '';
  }).filter(Boolean).join('; ');
}

function driveUrl_(file) {
  if (file.webViewLink && /^https:\/\//i.test(file.webViewLink)) return file.webViewLink;
  const id = encodeURIComponent(file.id || '');
  const base = file.mimeType === DI.FOLDER_MIME
    ? 'https://drive.google.com/drive/folders/' + id
    : 'https://drive.google.com/file/d/' + id + '/view';
  return base + (file.resourceKey ? '?resourcekey=' + encodeURIComponent(file.resourceKey) : '');
}

function parseFolderInput_(input) {
  const text = String(input || '').trim();
  if (!text) throw new Error('Paste a folder URL or ID.');
  let id = '', resourceKey = '';
  if (/^https?:\/\//i.test(text)) {
    if (!/^https:\/\/drive\.google\.com\//i.test(text)) throw new Error('Use a Google Drive folder URL or its ID.');
    const pathMatch = text.match(/\/folders\/([A-Za-z0-9_-]+)/);
    const idMatch = text.match(/[?&]id=([A-Za-z0-9_-]+)/);
    const keyMatch = text.match(/[?&]resourcekey=([^&#]+)/i);
    id = pathMatch ? pathMatch[1] : (idMatch ? idMatch[1] : '');
    if (keyMatch) resourceKey = decodeURIComponent(keyMatch[1]);
  } else id = text;
  if (!/^[A-Za-z0-9_-]{10,}$/.test(id) && id !== 'root') throw new Error('Could not find a valid Drive folder ID.');
  if (resourceKey && !/^[A-Za-z0-9_-]+$/.test(resourceKey)) throw new Error('The resource key in that URL is not valid.');
  return { id: id, resourceKey: resourceKey };
}

function joinPath_(parent, child) { return !parent || parent === '.' ? child : parent + ' / ' + child; }
function truncate_(text) { return text.length > DI.MAX_TEXT ? text.slice(0, DI.MAX_TEXT - 30) + ' … [truncated for cell limit]' : text; }
function quoteSheet_(name) { return "'" + name.replace(/'/g, "''") + "'"; }
function range_(id, startRow, endRow, startColumn, endColumn) {
  return { sheetId: id, startRowIndex: startRow, endRowIndex: endRow,
    startColumnIndex: startColumn, endColumnIndex: endColumn };
}
function columnName_(index) {
  let result = '', n = index + 1;
  while (n) { n--; result = String.fromCharCode(65 + n % 26) + result; n = Math.floor(n / 26); }
  return result;
}
function newSheetId_(used) {
  let id;
  do { id = Math.floor(Math.random() * 2000000000) + 1; } while (used.has(id));
  used.add(id);
  return id;
}
