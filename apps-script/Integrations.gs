function syncTaskToGoogleTasks_(task) {
  if (!task || !task.title) return null;
  if (!Tasks) {
    logEvent_('warn', 'Google Tasks advanced service not enabled', task, { source: 'system' });
    return null;
  }
  const taskListId = '@default';
  const newTask = {
    title: task.title,
    notes: task.details,
    due: task.due_date ? new Date(task.due_date).toISOString() : undefined
  };
  return Tasks.Tasks.insert(newTask, taskListId);
}

function createCalendarEventHook_(event) {
  if (!event || !event.title) return null;
  const calendar = CalendarApp.getDefaultCalendar();
  const start = event.start_time ? new Date(event.start_time) : new Date();
  const end = event.end_time ? new Date(event.end_time) : new Date(start.getTime() + 60 * 60 * 1000);
  return calendar.createEvent(event.title, start, end, { location: event.location || '' });
}

function logOutgoingEmail_(emailLog) {
  if (!emailLog || !emailLog.recipient) return null;
  GmailApp.sendEmail(emailLog.recipient, emailLog.subject, emailLog.body || '');
  return emailLog;
}
