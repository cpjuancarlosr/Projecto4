function runQuery(query) {
  const table = query.table;
  if (!table || !HEADERS[table]) {
    throw new Error('Invalid table.');
  }

  let rows = queryRecords(table, query.filters || {}, HEADERS[table]);

  if (query.sort && query.sort.field) {
    const direction = query.sort.direction === 'desc' ? -1 : 1;
    rows = rows.sort((a, b) => {
      const left = a[query.sort.field] || '';
      const right = b[query.sort.field] || '';
      if (left < right) {
        return -1 * direction;
      }
      if (left > right) {
        return 1 * direction;
      }
      return 0;
    });
  }

  if (query.limit) {
    rows = rows.slice(0, query.limit);
  }

  return rows;
}
