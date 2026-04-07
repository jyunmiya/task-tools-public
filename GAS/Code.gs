const APP_CONFIG = {
  spreadsheetIdProperty: 'QUALITY_TOOL_SPREADSHEET_ID',
  timezone: Session.getScriptTimeZone() || 'Asia/Tokyo',
  sheets: {
    taskMaster: 'TaskMaster',
    projects: 'Projects',
    projectChecks: 'ProjectChecks',
  },
  statuses: ['ok', 'ng', 'pending', 'na'],
  priorities: ['高', '中', '低'],
  phases: ['制作前', '制作', '公開前', '公開後'],
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('WEBサイト品質管理ツール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const payload = e && e.postData && e.postData.contents
    ? JSON.parse(e.postData.contents)
    : {};
  const action = payload.action;

  const handlers = {
    initializeSpreadsheet,
    getAppState,
    createProject,
    saveProjectTaskUpdates,
    getMasterItems,
    saveMasterItem,
    toggleMasterItem,
  };

  if (!handlers[action]) {
    return jsonOutput_({ ok: false, error: 'Unsupported action.' });
  }

  try {
    const result = handlers[action](payload.data || payload);
    return jsonOutput_({ ok: true, data: result });
  } catch (error) {
    return jsonOutput_({ ok: false, error: error.message });
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function jsonOutput_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
