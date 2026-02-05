function appendRecord(tableName, record, headers) {
  const sheet = ensureSheet(tableName, headers);
  const row = headers.map((header) => (record[header] !== undefined ? record[header] : ''));
  sheet.appendRow(row);
  return record.id;
}

function updateRecord(tableName, recordId, updates, headers) {
  const sheet = ensureSheet(tableName, headers);
  const data = sheet.getDataRange().getValues();
  const headerIndex = headers.reduce((acc, header, idx) => {
    acc[header] = idx;
    return acc;
  }, {});
  for (let i = 1; i < data.length; i += 1) {
    if (data[i][headerIndex.id] === recordId) {
      Object.keys(updates).forEach((key) => {
        if (headerIndex[key] !== undefined) {
          data[i][headerIndex[key]] = updates[key];
        }
      });
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
      return true;
    }
  }
  return false;
}

function queryRecords(tableName, filters, headers) {
  const sheet = ensureSheet(tableName, headers);
  const data = sheet.getDataRange().getValues();
  const headerIndex = headers.reduce((acc, header, idx) => {
    acc[header] = idx;
    return acc;
  }, {});
  let rows = data.slice(1).map((row) => {
    const record = {};
    headers.forEach((header) => {
      record[header] = row[headerIndex[header]];
    });
    return record;
  });

  if (filters) {
    Object.keys(filters).forEach((key) => {
      rows = rows.filter((row) => {
        if (filters[key] === undefined || filters[key] === '') {
          return true;
        }
        return String(row[key]).toLowerCase().includes(String(filters[key]).toLowerCase());
      });
    });
  }

  return rows;
}

function deleteRecord(tableName, recordId, headers) {
  const sheet = ensureSheet(tableName, headers);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i += 1) {
    if (data[i][0] === recordId) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}
