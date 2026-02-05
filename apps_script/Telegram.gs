function isTelegramWebhook_(e, params, payload) {
  return payload && (payload.message || payload.edited_message || payload.callback_query);
}

function parseTelegramUpdate_(update) {
  if (!update) return null;
  const message = update.message || update.edited_message || null;
  if (!message) return null;
  const text = message.text || '';
  const voice = message.voice || null;
  const photo = message.photo ? message.photo[message.photo.length - 1] : null;

  return {
    chat_id: message.chat && message.chat.id,
    from: message.from && message.from.username,
    text,
    voice,
    photo,
    date: message.date
  };
}

function telegramEventToPayload_(event) {
  let text = event.text;
  let transcription = '';
  if (event.voice) {
    transcription = transcribeTelegramVoice_(event.voice);
    text = text || transcription;
  }

  return {
    source: 'telegram',
    device: 'telegram',
    actor: event.from || 'telegram_user',
    timestamp: new Date(event.date * 1000).toISOString(),
    text: text,
    data: {
      transcription: transcription,
      voice_file_id: event.voice ? event.voice.file_id : '',
      photo_file_id: event.photo ? event.photo.file_id : ''
    }
  };
}

function transcribeTelegramVoice_(voice) {
  const token = getScriptProperty_(CONFIG.telegramTokenKey);
  const apiKey = getScriptProperty_(CONFIG.speechApiKey);
  if (!token || !apiKey || !voice) return '';

  const fileInfo = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getFile?file_id=' + voice.file_id).getContentText());
  if (!fileInfo.ok) return '';

  const filePath = fileInfo.result.file_path;
  const audioResponse = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + token + '/' + filePath);
  const audioBlob = audioResponse.getBlob();

  const speechPayload = {
    config: {
      encoding: 'OGG_OPUS',
      languageCode: 'es-MX',
      enableAutomaticPunctuation: true
    },
    audio: {
      content: Utilities.base64Encode(audioBlob.getBytes())
    }
  };

  const speechResponse = UrlFetchApp.fetch('https://speech.googleapis.com/v1/speech:recognize?key=' + apiKey, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(speechPayload)
  });

  const speechResult = JSON.parse(speechResponse.getContentText());
  if (!speechResult.results || !speechResult.results.length) return '';
  return speechResult.results.map((result) => result.alternatives[0].transcript).join(' ');
}
