function upsertEntity(entity) {
  const record = {
    id: entity.id || generateUuid(),
    timestamp: nowIso(),
    name: entity.name || '',
    type: entity.type || 'general',
    status: entity.status || 'active',
    priority: parsePriority(entity.priority),
    notes: entity.notes || '',
    metadata: entity.metadata ? JSON.stringify(entity.metadata) : ''
  };

  const updated = updateRecord(CONFIG.TABLES.ENTITIES, record.id, record, HEADERS.entities);
  if (!updated) {
    appendRecord(CONFIG.TABLES.ENTITIES, record, HEADERS.entities);
  }
  return record;
}
