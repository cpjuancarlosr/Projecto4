function handleTelegramUpdate_(update) {
  const message = update.message || update.edited_message || {};
  const chat = message.chat || {};
  const text = message.text || '';
  const actor = chat.username || chat.title || String(chat.id || 'telegram');
  const payload = {
    source: 'telegram',
    actor,
    device: 'telegram',
    text,
    telegram_update_id: update.update_id,
    telegram_chat_id: chat.id,
    telegram_message_id: message.message_id,
    intent: deriveTelegramIntent_(text)
  };
  if (message.voice) {
    payload.voice = message.voice;
    payload.text = transcribeTelegramVoice_(message.voice);
  }
  if (message.photo) {
    payload.photo = message.photo;
  }
  return handleCapture_(payload, 'telegram');
}

function deriveTelegramIntent_(text) {
  if (!text) return '';
  if (text.startsWith('/')) {
    return text.replace('/', '').split(' ')[0];
  }
  return '';
}

function transcribeTelegramVoice_(voice) {
  if (!voice || !voice.file_id) return '';
  const fileUrl = getTelegramFileUrl_(voice.file_id);
  if (!fileUrl) return '';
  return transcribeAudio_(fileUrl);
}

function getTelegramFileUrl_(fileId) {
  if (!CONFIG.telegram.botToken) return '';
  const url = `https://api.telegram.org/bot${CONFIG.telegram.botToken}/getFile?file_id=${fileId}`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());
  if (!data.ok) return '';
  return `https://api.telegram.org/file/bot${CONFIG.telegram.botToken}/${data.result.file_path}`;
}

function transcribeAudio_(fileUrl) {
  try {
    const response = UrlFetchApp.fetch(fileUrl);
    const blob = response.getBlob();
    const audioBytes = Utilities.base64Encode(blob.getBytes());
    const request = {
      config: {
        encoding: 'OGG_OPUS',
        languageCode: CONFIG.speech.languageCode
      },
      audio: {
        content: audioBytes
      }
    };
    const speechUrl = 'https://speech.googleapis.com/v1/speech:recognize';
    const speechResponse = UrlFetchApp.fetch(speechUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(request),
      headers: {
        Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
      }
    });
    const result = JSON.parse(speechResponse.getContentText());
    if (!result.results || !result.results.length) return '';
    return result.results[0].alternatives[0].transcript || '';
  } catch (error) {
    logEvent_('error', 'Speech-to-text failed', { error: error.message, fileUrl }, { source: 'telegram' });
    return '';
  }
}
