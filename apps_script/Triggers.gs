function createTimeTriggers() {
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('weeklySummary').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
}

function dailySummary() {
  const tasks = queryRecords(CONFIG.TABLES.TASKS, { status: 'open' }, HEADERS.tasks);
  logEvent('info', 'triggers', `Daily summary: ${tasks.length} open tasks`, { count: tasks.length });
}

function weeklySummary() {
  const movements = queryRecords(CONFIG.TABLES.FINANCIAL_MOVEMENTS, {}, HEADERS.financial_movements);
  logEvent('info', 'triggers', `Weekly summary: ${movements.length} movements`, { count: movements.length });
}
