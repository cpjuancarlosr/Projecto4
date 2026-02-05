function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Html');
  template.baseUrl = ScriptApp.getService().getUrl();
  return template.evaluate().setTitle('Capture Hub').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function handleApiRequest(payload) {
  const action = payload.action;
  if (!authorizeApiRequest(payload)) {
    return ContentService.createTextOutput('Unauthorized');
  }

  if (action === 'capture') {
    const record = handleCapture(payload.data || {});
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', record }));
  }

  if (action === 'query') {
    const result = runQuery(payload.query || {});
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', result }));
  }

  if (action === 'entity') {
    const record = upsertEntity(payload.data || {});
    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', record }));
  }

  return ContentService.createTextOutput('Unknown action');
}

function authorizeApiRequest(payload) {
  if (!CONFIG.API_SHARED_SECRET) {
    return true;
  }
  return payload && payload.secret === CONFIG.API_SHARED_SECRET;
}
