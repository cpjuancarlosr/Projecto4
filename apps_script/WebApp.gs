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
    return buildJsonResponse({ status: 'error', message: 'Unauthorized' });
  }

  if (action === 'capture') {
    const record = handleCapture(payload.data || {});
    return buildJsonResponse({ status: 'ok', record });
  }

  if (action === 'query') {
    const result = runQuery(payload.query || {});
    return buildJsonResponse({ status: 'ok', result });
  }

  if (action === 'entity') {
    const record = upsertEntity(payload.data || {});
    return buildJsonResponse({ status: 'ok', record });
  }

  return buildJsonResponse({ status: 'error', message: 'Unknown action' });
}

function authorizeApiRequest(payload) {
  if (!CONFIG.API_SHARED_SECRET) {
    return true;
  }
  return payload && payload.secret === CONFIG.API_SHARED_SECRET;
}
