// ---------- CONFIG ----------
const MASTER_SHEET = 'Master_Ocorrencias';
const ASSESSORES_SHEET = 'Assessores';
const TIPOS_SHEET = 'Tipos_Ocorrencia';
const ORIGENS_SHEET = 'Origens';
const AUDIT_SHEET = 'Audit_Log';
const ATTACH_FOLDER_PREFIX = 'NAC_Attachments_';
const ALLOWED_STATUSES = ['Aberta','Em andamento','Concluída','Cancelada'];

/** Serve o SPA */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('NAC — Ocorrências')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Utilitário para incluir HTML (se quiser modularizar futuramente) */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Retorna email do usuário */
function getCurrentUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email || 'unknown';
  } catch(e) {
    return 'unknown';
  }
}

/** Retorna dados de configuração: assessores, tipos e origens */
function getConfigData() {
  const ss = SpreadsheetApp.getActive();
  const readTable = (name) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return [];
    const vals = sh.getDataRange().getValues();
    const headers = vals[0] || [];
    const out = [];
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      if (row.join('').trim() === '') continue;
      const obj = {};
      for (let c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
      out.push(obj);
    }
    return out;
  };
  return {
    user: getCurrentUserEmail(),
    assessores: readTable(ASSESSORES_SHEET),
    tipos: readTable(TIPOS_SHEET),
    origens: readTable(ORIGENS_SHEET)
  };
}

/** Lê master pela planilha e converte em array de objetos */
function readMasterAsObjects() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(MASTER_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];
  const data = sh.getRange(1,1,lastRow,lastCol).getValues();
  const headers = data[0].map(h => String(h||'').trim());
  const out = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const obj = {};
    for (let c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
    out.push(obj);
  }
  return out;
}

/** Lista ocorrências com filtros, ordenação e paginação
 * params: {search, assessor_name, type_name, status, startDate, endDate, page, pageSize}
 */
function listOccurrences(params) {
  params = params || {};
  const page = Math.max(1, parseInt(params.page || 1));
  const pageSize = Math.max(10, parseInt(params.pageSize || 20));
  const all = readMasterAsObjects();

  // filtrar
  let rows = all.filter(r => r['occurrence_id']); // somente válidos
  if (params.search) {
    const q = String(params.search).toLowerCase();
    rows = rows.filter(r => (
      (r['occurrence_id'] && String(r['occurrence_id']).toLowerCase().indexOf(q) >= 0) ||
      (r['assessor_name'] && String(r['assessor_name']).toLowerCase().indexOf(q) >= 0) ||
      (r['description'] && String(r['description']).toLowerCase().indexOf(q) >= 0)
    ));
  }
  if (params.assessor_name) rows = rows.filter(r => String(r['assessor_name']) === String(params.assessor_name));
  if (params.type_name) rows = rows.filter(r => String(r['type_name']) === String(params.type_name));
  if (params.status) rows = rows.filter(r => String(r['status']) === String(params.status));
  if (params.turno) {
    rows = rows.filter(r => {
      const turno = detectShift(r);
      return turno && String(turno).toUpperCase() === String(params.turno).toUpperCase();
    });
  }

  // date filter (date_time stored as ISO string 'YYYY-MM-DDTHH:MM:SS' or Date object)
  const parseDate = v => {
    if (!v) return null;
    if (v instanceof Date) return v;
    try { return new Date(String(v)); } catch(e){ return null; }
  };
  if (params.startDate) {
    const sd = new Date(params.startDate + 'T00:00:00');
    rows = rows.filter(r => {
      const dt = parseDate(r['date_time']);
      return dt && dt >= sd;
    });
  }
  if (params.endDate) {
    const ed = new Date(params.endDate + 'T23:59:59');
    rows = rows.filter(r => {
      const dt = parseDate(r['date_time']);
      return dt && dt <= ed;
    });
  }

  // ordenar por date_time desc
  rows.sort((a,b) => {
    const da = parseDate(a['date_time']) || new Date(0);
    const db = parseDate(b['date_time']) || new Date(0);
    return db - da;
  });

  const total = rows.length;
  const startIndex = (page - 1) * pageSize;
  const pageRows = rows.slice(startIndex, startIndex + pageSize);

  return { total: total, page: page, pageSize: pageSize, rows: pageRows };
}

/** Valida campos essenciais e sanitiza tipos básicos */
function validateOccurrenceData(data, isUpdate) {
  if (!data) return {valid:false, message:'Dados vazios'};
  const str = (v) => v === undefined || v === null ? '' : String(v).trim();
  const date = str(data.date);
  const time = str(data.time);
  const assessor = str(data.assessor_name);
  const type = str(data.type_name);
  const status = str(data.status || 'Aberta');
  const description = str(data.description);
  if (!date) return {valid:false, message:'Data é obrigatória'};
  if (!time) return {valid:false, message:'Hora é obrigatória'};
  if (!assessor) return {valid:false, message:'Assessor é obrigatório'};
  if (!type) return {valid:false, message:'Tipo é obrigatório'};
  if (description.length < 8) return {valid:false, message:'Descrição deve ter ao menos 8 caracteres'};
  if (ALLOWED_STATUSES.indexOf(status) === -1) return {valid:false, message:'Status inválido'};
  if (!isUpdate) {
    data.status = status; // normaliza
  }
  data.assessor_name = assessor;
  data.type_name = type;
  data.description = description;
  return {valid:true};
}

/** Deduz turno MT/SN a partir do campo shift/turno ou horário */
function detectShift(row) {
  if (!row) return '';
  const explicit = row['shift'] || row['turno'] || row['turn'] || row['periodo'];
  if (explicit) return String(explicit).toUpperCase();
  if (row['date_time']) {
    const dt = row['date_time'] instanceof Date ? row['date_time'] : new Date(String(row['date_time']));
    if (dt && !isNaN(dt)) return dt.getHours() < 18 ? 'MT' : 'SN';
  }
  return '';
}

/** Recupera uma ocorrência por id */
function getOccurrenceById(id) {
  if (!id) return null;
  const all = readMasterAsObjects();
  for (let r of all) if (String(r['occurrence_id']) === String(id)) return r;
  return null;
}

/** Gera novo occurrence_id para uma data (yyyymmdd-###) */
function generateOccurrenceId(dateObj) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const ymd = Utilities.formatDate(dateObj, tz, 'yyyyMMdd');
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  const colIds = master.getRange(2,1,Math.max(0, master.getLastRow()-1),1).getValues().flat();
  let seq = 1;
  for (let id of colIds) {
    if (String(id).indexOf('OC' + ymd) === 0) seq++;
  }
  return 'OC' + ymd + '-' + String(seq).padStart(3,'0');
}

/** Salva arquivo base64 no Drive e retorna URL */
function saveBase64File(filename, mimeType, base64Data) {
  if (!base64Data) return null;
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType, filename);
  const folder = getOrCreateAttachmentFolder();
  const file = folder.createFile(blob);
  // definir permissão de leitura por link
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}
  return file.getUrl();
}

/** Cria pasta de anexos para essa planilha (ou retorna se já existe) */
function getOrCreateAttachmentFolder() {
  const ss = SpreadsheetApp.getActive();
  const id = ss.getId();
  const name = ATTACH_FOLDER_PREFIX + id;
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();

  // cria no Drive raiz
  const folder = DriveApp.createFolder(name);
  folder.setDescription('Anexos do webapp NAC — planilha ' + id);
  return folder;
}

/** Cria uma ocorrência (form = objeto com campos) */
function createOccurrence(form) {
  const validation = validateOccurrenceData(form);
  if (!validation.valid) throw new Error(validation.message);
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  if (!master) throw new Error('Master_Ocorrencias não encontrada.');

  const tz = ss.getSpreadsheetTimeZone();
  const dt = form.date && form.time ? new Date(form.date + 'T' + form.time) : (form.date ? new Date(form.date+'T09:00:00') : new Date());
  const occurrenceId = generateOccurrenceId(dt);
  const createdAt = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");

  const headers = master.getRange(1,1,1,master.getLastColumn()).getValues()[0];
  const newRow = Array(headers.length).fill('');
  const setField = (key, value) => {
    const idx = headers.indexOf(key);
    if (idx >= 0) newRow[idx] = value;
  };

  setField('occurrence_id', occurrenceId);
  setField('date_time', Utilities.formatDate(dt, tz, "yyyy-MM-dd'T'HH:mm:ss"));
  setField('assessor_id', form.assessor_id || '');
  setField('assessor_name', form.assessor_name || '');
  setField('type_name', form.type_name || '');
  setField('origin', form.origin || '');
  setField('priority', form.priority || '');
  setField('status', form.status || 'Aberta');
  setField('description', form.description || '');
  setField('action_taken', form.action_taken || '');
  setField('follow_up_date', form.follow_up_date || '');
  setField('attachments', form.attachments || '');
  setField('related_case_id', form.related_case_id || '');
  setField('created_by', getCurrentUserEmail());
  setField('created_at', createdAt);
  setField('shift', form.shift || '');
  setField('turno', form.shift || '');
  setField('turn', form.shift || '');

  master.appendRow(newRow);

  // audit
  const audit = ss.getSheetByName(AUDIT_SHEET);
  if (audit) audit.appendRow([Utilities.getUuid(), occurrenceId, createdAt, getCurrentUserEmail(), 'CREATE', '', JSON.stringify(newRow), '']);

  return { success: true, id: occurrenceId };
}

/** Atualiza ocorrência por id (data = objeto com campos para atualizar) */
function updateOccurrence(id, data) {
  if (typeof id === 'object' && data === undefined) { // suporte a chamada via {id, data}
    data = id.data;
    id = id.id;
  }
  if (!id) throw new Error('id vazio');
  const validation = validateOccurrenceData(data, true);
  if (!validation.valid) throw new Error(validation.message);
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  const hdr = master.getRange(1,1,1,master.getLastColumn()).getValues()[0];
  const lastRow = master.getLastRow();
  const vals = master.getRange(2,1,lastRow-1,master.getLastColumn()).getValues();

  let found = false;
  let foundRowIndex = -1; // index in sheet (1-based)
  for (let r = 0; r < vals.length; r++) {
    if (String(vals[r][0]) === String(id)) {
      found = true;
      foundRowIndex = r + 2; // sheet row
      break;
    }
  }
  if (!found) return { success: false, message: 'Registro não encontrado' };

  // criar mapa header->col
  const headerMap = {};
  for (let c = 0; c < hdr.length; c++) headerMap[hdr[c]] = c+1;

  // coletar alterações para audit
  const oldRow = master.getRange(foundRowIndex,1,1,master.getLastColumn()).getValues()[0];
  const updates = [];
  for (let key in data) {
    if (!headerMap[key]) continue;
    const col = headerMap[key];
    const oldVal = oldRow[col-1];
    const newVal = data[key];
    if (String(oldVal) !== String(newVal)) {
      master.getRange(foundRowIndex, col).setValue(newVal);
      updates.push({field:key, old:oldVal, new:newVal});
    }
  }
  const updatedAt = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  if (headerMap['updated_by']) master.getRange(foundRowIndex, headerMap['updated_by']).setValue(getCurrentUserEmail());
  if (headerMap['updated_at']) master.getRange(foundRowIndex, headerMap['updated_at']).setValue(updatedAt);

  // gravar audit log
  const audit = ss.getSheetByName(AUDIT_SHEET);
  if (audit && updates.length) {
    updates.forEach(u => {
      audit.appendRow([Utilities.getUuid(), id, updatedAt, getCurrentUserEmail(), u.field, u.old, u.new, 'update via webapp']);
    });
  }
  return { success:true, updated: updates.length };
}

/** Marca como cancelado (ou deleta) */
function deleteOccurrence(id) {
  // neste app, fazemos soft delete: status = Cancelada
  const res = updateOccurrence(id, { status: 'Cancelada' });
  return res;
}

/** Retorna dados do dashboard */
function getDashboardData() {
  const all = readMasterAsObjects();
  const total = all.length;
  const byStatus = {};
  const byType = {};
  const byDay = {};
  const byShift = { MT:0, SN:0 };
  const openVsClosed = { abertas:0, concluidas:0 };
  const lastList = all.slice(0,10); // ordered already by date desc in listOccurrences, but here gets original order; we'll sort
  // garantir ordenação por date desc
  lastList.sort((a,b) => new Date(b['date_time'] || 0) - new Date(a['date_time'] || 0));

  // calcular avg resolution (dias) para registros com status Concluída e updated_at
  let sumDays = 0, closedCount = 0;
  for (let r of all) {
    const st = r['status'] || 'Aberta';
    byStatus[st] = (byStatus[st] || 0) + 1;
    const t = r['type_name'] || 'Outros';
    byType[t] = (byType[t] || 0) + 1;
    const dayKey = (r['date_time'] || '').toString().slice(0,10);
    if (dayKey) byDay[dayKey] = (byDay[dayKey] || 0) + 1;
    const turno = detectShift(r);
    if (turno && byShift[turno] !== undefined) byShift[turno] += 1;
    if ((String(st).toLowerCase().indexOf('conclu') >= 0)) openVsClosed.concluidas += 1; else openVsClosed.abertas += 1;
    if ((String(st).toLowerCase().indexOf('conclu') >= 0) && r['created_at'] && r['updated_at']) {
      const d1 = new Date(r['created_at']);
      const d2 = new Date(r['updated_at']);
      if (d1 && d2 && !isNaN(d1) && !isNaN(d2)) {
        const diffDays = (d2 - d1) / (1000*60*60*24);
        sumDays += diffDays;
        closedCount++;
      }
    }
  }
  const avgResolution = closedCount ? (sumDays / closedCount) : null;

  // últimas 10 ocorrências
  const sorted = all.slice().sort((a,b) => new Date(b['date_time']||0) - new Date(a['date_time']||0));
  const latest = sorted.slice(0,10);

  const totalsByDay = Object.keys(byDay).sort((a,b)=> new Date(b) - new Date(a)).slice(0,7).map(day => ({day:day, total: byDay[day]}));

  return { total, byStatus, byType, avgResolutionDays: avgResolution, latest, totalsByDay, byShift, openVsClosed };
}

/** Recupera audit log (opcional filtro por occurrence_id) */
function getAuditLog(occurrence_id) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(AUDIT_SHEET);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  const headers = vals[0];
  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const obj = {};
    for (let c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
    if (!occurrence_id || String(obj['occurrence_id']) === String(occurrence_id)) out.push(obj);
  }
  return out;
}

/** Função utilitária: migrar planilhas numa pasta -> Master (já existente no template anterior) */
function migrarArquivosNaPastaParaMaster(folderId) {
  // reusa a função criada antes se preferir. Aqui mantemos simples: abre todas as planilhas na pasta e tenta migrar cada aba.
  const results = [];
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(MASTER_SHEET);
  if (!master) return {success:false, message:'Master não encontrada'};
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    const file = files.next();
    const id = file.getId();
    const sourcedoc = SpreadsheetApp.openById(id);
    const sheets = sourcedoc.getSheets();
    for (let s of sheets) {
      try {
        // tentar migrar aba usando lógica adaptada ao layout tipo x 01D/01N
        const r = migrarAbaSheetToMasterObject(s, master);
        results.push({file:file.getName(), sheet:s.getName(), result:r});
      } catch(e) {
        results.push({file:file.getName(), sheet:s.getName(), result:{error: e.message}});
      }
    }
  }
  return results;
}

/** Função de migração por aba (versão simplificada; reusa lógica do template original) */
function migrarAbaSheetToMasterObject(sheetObj, master) {
  // Aqui podemos reutilizar a lógica do script que você já executou — para simplicidade deixo uma implementação robusta e curta
  const data = sheetObj.getDataRange().getValues();
  let tipoRow = -1, tipoCol = -1;
  for (let r = 0; r < Math.min(30, data.length); r++) {
    for (let c = 0; c < data[r].length; c++) {
      if (String(data[r][c]).trim().toUpperCase() === 'TIPO') { tipoRow=r; tipoCol=c; break; }
    }
    if (tipoRow>=0) break;
  }
  if (tipoRow < 0) return {migrated:0, message:'TIPO não encontrado em ' + sheetObj.getName()};
  // localizar linha de cabeçalhos 01D/01N
  let headerRowCandidate = -1;
  for (let r = tipoRow; r < Math.min(tipoRow+10, data.length); r++) {
    for (let c = tipoCol+1; c < data[r].length; c++) {
      if (/^\d{2}[DN]$/i.test(String(data[r][c]).trim())) { headerRowCandidate = r; break; }
    }
    if (headerRowCandidate >=0) break;
  }
  if (headerRowCandidate < 0) return {migrated:0, message:'colunas 01D/01N não encontradas'};
  const dayCols = [];
  for (let c = tipoCol+1; c < data[headerRowCandidate].length; c++) {
    const h = String(data[headerRowCandidate][c]).trim();
    if (/^\d{2}[DN]$/i.test(h)) dayCols.push({colIndex:c, headerText:h.toUpperCase()});
  }
  const tipos = [];
  for (let r = tipoRow+1; r < data.length; r++) {
    const v = String(data[r][tipoCol]).trim();
    if (v === '') break;
    tipos.push({rowIndex:r, typeName:v});
  }
  // inferir mês da planilha
  const monthYear = inferMonthYearFromText(sheetObj.getParent().getName() + ' ' + sheetObj.getName());
  const detectedMonth = monthYear.month;
  const detectedYear = monthYear.year;
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const out = [];
  const seqByDay = {};
  tipos.forEach(t => {
    dayCols.forEach(dc => {
      const cell = data[t.rowIndex][dc.colIndex];
      const cellStr = (cell === null||cell===undefined) ? '' : String(cell).trim();
      if (!cellStr || cellStr === '#REF!') return;
      const dayNum = parseInt(dc.headerText.slice(0,2),10);
      const turno = dc.headerText.slice(2).toUpperCase();
      const hour = (turno === 'D') ? 9 : 21;
      const dt = new Date(detectedYear, detectedMonth-1, dayNum, hour,0,0);
      const ymd = Utilities.formatDate(dt, tz, 'yyyyMMdd');
      seqByDay[ymd] = (seqByDay[ymd]||0)+1;
      const occurrenceId = 'OC' + ymd + '-' + String(seqByDay[ymd]).padStart(3,'0');
      let status = 'Aberta';
      const up = cellStr.toUpperCase();
      if (up.indexOf('CONCLU') >= 0) status = 'Concluída';
      if (up.indexOf('PENDENTE') >= 0) status = 'Pendente';
      const rowOut = [
        occurrenceId,
        Utilities.formatDate(dt, tz, "yyyy-MM-dd'T'HH:mm:ss"),
        '', '', // assessor
        t.typeName, '', '', status, cellStr, '', '', '', '', 'migration', Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss"), '', ''
      ];
      out.push(rowOut);
    });
  });
  if (out.length) master.getRange(master.getLastRow()+1,1,out.length,out[0].length).setValues(out);
  return {migrated: out.length, message: 'OK'};
}

/** tenta inferir mês/ano do nome */
function inferMonthYearFromText(text) {
  const meses = {'JANEIRO':1,'FEVEREIRO':2,'MARÇO':3,'MARCO':3,'ABRIL':4,'MAIO':5,'JUNHO':6,'JULHO':7,'AGOSTO':8,'SETEMBRO':9,'OUTUBRO':10,'NOVEMBRO':11,'DEZEMBRO':12};
  const upper = (text||'').toUpperCase();
  for (let m in meses) if (upper.indexOf(m) >= 0) return {month: meses[m], year: (new Date()).getFullYear()};
  return {month: (new Date()).getMonth()+1, year: (new Date()).getFullYear()};
}
