function doPost(e) {
  if (!e || !e.postData) {
    return ContentService.createTextOutput('No payload');
  }

  const payload = JSON.parse(e.postData.contents);
  if (payload && payload.message) {
    return handleTelegramUpdate(payload);
  }

  if (payload && payload.action) {
    return handleApiRequest(payload);
  }

  return ContentService.createTextOutput('Unsupported payload');
}

function handleTelegramUpdate(update) {
  const message = update.message || {};
  const chatId = String(message.chat && message.chat.id ? message.chat.id : '');
  if (!isAuthorizedTelegramChat(chatId)) {
    return ContentService.createTextOutput('Unauthorized');
  }

  const text = message.text || '';
  const commandPayload = parseTelegramCommand(text);
  const capturePayload = {
    text: commandPayload.text,
    intent: commandPayload.intent,
    source: 'telegram',
    device: 'telegram',
    actor: message.from ? message.from.username || message.from.id : '',
    entity_id: commandPayload.entity_id,
    status: 'open',
    priority: commandPayload.priority,
    tags: commandPayload.tags
  };

  if (message.voice) {
    const transcription = transcribeTelegramVoice(message.voice.file_id);
    if (transcription) {
      capturePayload.text = `${capturePayload.text} ${transcription}`.trim();
    }
    capturePayload.metadata = {
      voice_file_id: message.voice.file_id,
      duration: message.voice.duration
    };
  }

  if (message.photo) {
    capturePayload.metadata = {
      photo: message.photo,
      caption: message.caption || ''
    };
  }

  const record = handleCapture(capturePayload);
  sendTelegramMessage(chatId, `✅ Capturado: ${record.id}`);
  return ContentService.createTextOutput('ok');
}

function parseTelegramCommand(text) {
  if (!text) {
    return { text: '', intent: '' };
  }
  const parts = text.trim().split(' ');
  const command = parts[0].toLowerCase();
  const body = text.replace(parts[0], '').trim();
  const base = {
    text: body || text,
    intent: '',
    priority: '',
    tags: '',
    entity_id: ''
  };

  if (command === '/note') {
    base.intent = 'note';
    base.text = body;
  } else if (command === '/task') {
    base.intent = 'task';
    base.text = body;
  } else if (command === '/money') {
    base.intent = 'money';
    base.text = body;
  } else if (command === '/treasure') {
    base.intent = 'treasure';
    base.text = body;
  } else if (command === '/email') {
    base.intent = 'email';
    base.text = body;
  } else if (command === '/event') {
    base.intent = 'calendar';
    base.text = body;
  }

  const priorityMatch = base.text.match(/\|(alta|media|baja|high|medium|low)/i);
  if (priorityMatch) {
    base.priority = priorityMatch[1];
  }

  const tagMatch = base.text.match(/#(\w+)/g);
  if (tagMatch) {
    base.tags = tagMatch.map((tag) => tag.replace('#', '')).join(',');
  }

  const entityMatch = base.text.match(/entidad:([\w-]+)/i);
  if (entityMatch) {
    base.entity_id = entityMatch[1];
  }

  return base;
}

function isAuthorizedTelegramChat(chatId) {
  if (!CONFIG.TELEGRAM_ALLOWED_CHAT_IDS.length) {
    return true;
  }
  return CONFIG.TELEGRAM_ALLOWED_CHAT_IDS.includes(chatId);
}

function sendTelegramMessage(chatId, text) {
  if (!CONFIG.TELEGRAM_BOT_TOKEN) {
    return;
  }
  const url = `https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`;
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: chatId,
      text
    })
  });
}

function transcribeTelegramVoice(fileId) {
  if (!CONFIG.TELEGRAM_BOT_TOKEN) {
    return '';
  }
  const fileUrl = `https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/getFile?file_id=${fileId}`;
  const fileResponse = UrlFetchApp.fetch(fileUrl);
  const fileData = JSON.parse(fileResponse.getContentText());
  if (!fileData || !fileData.result) {
    return '';
  }
  const filePath = fileData.result.file_path;
  const audioUrl = `https://api.telegram.org/file/bot${CONFIG.TELEGRAM_BOT_TOKEN}/${filePath}`;
  const audioBlob = UrlFetchApp.fetch(audioUrl).getBlob();
  return transcribeAudio(audioBlob);
}

function transcribeAudio(audioBlob) {
  if (!CONFIG.SPEECH_API_KEY) {
    return '';
  }
  const audioBytes = Utilities.base64Encode(audioBlob.getBytes());
  const requestBody = {
    config: {
      encoding: 'OGG_OPUS',
      languageCode: 'es-MX',
      enableAutomaticPunctuation: true
    },
    audio: {
      content: audioBytes
    }
  };
  const response = UrlFetchApp.fetch(`https://speech.googleapis.com/v1/speech:recognize?key=${CONFIG.SPEECH_API_KEY}`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody)
  });
  const result = JSON.parse(response.getContentText());
  if (!result.results || !result.results.length) {
    return '';
  }
  return result.results.map((entry) => entry.alternatives[0].transcript).join(' ');
}
