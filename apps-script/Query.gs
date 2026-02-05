function queryTable_(table, filters, options) {
  const sheet = getSheet_(table);
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  let data = values.map((row) => mapRow_(headers, row));
  const filterKeys = Object.keys(filters || {}).filter((key) => filters[key] !== '' && filters[key] !== undefined);
  if (filterKeys.length) {
    data = data.filter((row) => {
      return filterKeys.every((key) => {
        const target = String(row[key] || '').toLowerCase();
        const expected = String(filters[key] || '').toLowerCase();
        return target.includes(expected);
      });
    });
  }
  if (options?.sortBy) {
    data.sort((a, b) => {
      const left = a[options.sortBy] || '';
      const right = b[options.sortBy] || '';
      return String(left).localeCompare(String(right));
    });
  }
  if (options?.limit) {
    data = data.slice(0, options.limit);
  }
  return data;
}

function mapRow_(headers, row) {
  return headers.reduce((acc, header, index) => {
    acc[header] = row[index];
    return acc;
  }, {});
}
