function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('weeklySummary').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
}

function dailySummary() {
  const today = Utilities.formatDate(new Date(), CONFIG.timezone, 'yyyy-MM-dd');
  const tasks = queryTable_('tasks', { status: 'open' }, { limit: 20 });
  logEvent_('info', `Daily summary ${today}`, { tasks_count: tasks.length }, { source: 'system' });
}

function weeklySummary() {
  const tasks = queryTable_('tasks', { status: 'open' }, { limit: 50 });
  logEvent_('info', 'Weekly summary', { tasks_count: tasks.length }, { source: 'system' });
}
