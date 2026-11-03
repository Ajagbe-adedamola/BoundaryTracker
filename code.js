// HQ TERRITORY MAPPER v6 (Sun King) - bound to the territory sheet.
// v7: reads ONLY the first tab (the first one that is not a system sheet). Everything else is generated.
// (1) "Active Areas" sheet lists every registered HQ with all coordinates, rebuilt after every change;
// (2) edits typed in the first tab are captured as timestamped "latest" entries, but only AFTER you stop typing
//     for a minute and only if the row is complete (debounced, so half-typed rows are never saved);
// (3) a post-sync check compares the Active Areas sheet with what the map reads.
// ONE-TIME SETUP: run Territory > Install auto-capture (needed because timers require authorisation).
// Function names/return shapes are unchanged, so WebApp.html needs no edits.
const SPREADSHEET_ID = '1jhLQN7XWHTIdiZ1_3mNIi0zxpKaMCSn0FihWPaOTOeg'; // fallback if the script is not bound to the sheet
const E = 'HQ_Entries', CFG = 'Config', ERR = 'Error_Log', AA = 'Active Areas', MAXP = 10, FIRSTP = 7; // FIRSTP = column of Point 1 Lat
const HQ_COLORS = ['#E53935','#8E24AA','#1E88E5','#00897B','#F4511E','#3949AB','#039BE5','#43A047','#FB8C00','#6D4C41'];
const NON_SOURCE = [E, CFG, ERR, AA, 'Boundaries', 'Distances', 'HQs', 'HQ_Areas'];

function ss() { return SpreadsheetApp.getActive() || SpreadsheetApp.openById(SPREADSHEET_ID); }
function doGet() {
  return HtmlService.createHtmlOutputFromFile('WebApp').setTitle('Sun King Territory Mapper')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Territory')
      .addItem('Check what the map reads', 'checkSheet')
      .addItem('Refresh Active Areas', 'refreshActiveAreas')
      .addItem('Capture first-tab edits now', 'commitPendingEdits')
      .addItem('Install auto-capture (run once)', 'installTriggers')
      .addToUi();
  } catch (e) {}
}
function who() { try { return Session.getActiveUser().getEmail() || 'unknown'; } catch (e) { return 'unknown'; } }
function fail(fn, e, d) { logError(fn, e.message, d); return { error: e.message }; }
function logError(fn, m, d) {
  try { const s = ss().getSheetByName(ERR); if (s) s.appendRow([new Date().toISOString(), fn, String(m).slice(0, 200), String(d || '').slice(0, 500)]); } catch (_) {}
}
function initSheets() {
  const b = ss();
  if (!b.getSheetByName(E)) {
    const s = b.insertSheet(E), h = ['HQ_Name', 'Region', 'Status', 'Entry_On', 'Entered_By', 'Source'];
    for (let i = 1; i <= MAXP; i++) h.push('Point ' + i + ' Lat', 'Point ' + i + ' Lng');
    s.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#FFC700');
    s.setFrozenRows(1);
    s.setColumnWidth(1, 170); s.setColumnWidth(4, 170); s.setColumnWidth(5, 190); s.setColumnWidth(6, 200);
    s.getRange(2, FIRSTP, 2000, MAXP * 2).setNumberFormat('0.000000');
  }
  if (!b.getSheetByName(ERR)) b.insertSheet(ERR).appendRow(['Timestamp', 'Function', 'Message', 'Detail']);
}

// ---------- core: latest entry per HQ wins ----------
// Order, oldest to newest: (1) the table(s) in the FIRST tab, top to bottom;
// (2) HQ_Entries rows (app edits AND first-tab edits captured by the debounced trigger), by row order = time order.
let _cur = null;
function readCurrent() {
  if (_cur) return _cur;
  initSheets();
  const by = {}, order = [];
  const add = (name, pts, m) => {
    const k = name.toLowerCase();
    if (!by[k]) { by[k] = { entries: 0, region: '' }; order.push(k); }
    const o = by[k]; o.entries++; o.name = name; o.points = pts; o.active = m.active !== false;
    o.on = m.on || ''; o.by = m.by || ''; o.source = m.source || '';
    if (m.region) o.region = m.region;
  };
  sourceSheets().forEach(sh => {
    try { parseTab(sh).entries.forEach(e => add(e.name, e.points, { source: sh.getName() + ' row ' + e.row })); }
    catch (x) { logError('readCurrent', x.message, sh.getName()); }
  });
  const s = ss().getSheetByName(E), n = s.getLastRow();
  if (n > 1) s.getRange(2, 1, n - 1, FIRSTP - 1 + MAXP * 2).getValues().forEach(r => {
    const name = String(r[0]).trim(); if (!name) return;
    const pts = [];
    for (let j = 0; j < MAXP; j++) {
      const la = parseFloat(r[FIRSTP - 1 + j * 2]), lo = parseFloat(r[FIRSTP + j * 2]);
      if (isNaN(la) || isNaN(lo)) break;
      pts.push({ lat: la, lng: lo });
    }
    add(name, pts, { active: String(r[2] || 'Active') !== 'Inactive', region: String(r[1]),
      on: r[3] ? new Date(r[3]).toISOString() : '', by: String(r[4]), source: String(r[5]) });
  });
  _cur = order.map((k, i) => Object.assign(by[k], { id: 'HQ' + String(i + 1).padStart(3, '0') })).filter(h => h.active);
  return _cur;
}
function put(name, region, pts, status, src) {
  const a = [name, region || '', status || 'Active', new Date().toISOString(), who(), src || 'webapp'];
  for (let i = 0; i < MAXP; i++) a.push(pts[i] ? pts[i].lat : '', pts[i] ? pts[i].lng : '');
  ss().getSheetByName(E).appendRow(a); _cur = null;
}
function find(id) { return readCurrent().find(h => h.id === id); }
function parseManualCoords(lat, lng) {
  const a = parseFloat(lat), b = parseFloat(lng);
  if (isNaN(a) || isNaN(b) || Math.abs(a) > 90 || Math.abs(b) > 180) return null;
  return { lat: a, lng: b };
}
function centroid(p) {
  return p.length ? { lat: p.reduce((s, x) => s + x.lat, 0) / p.length, lng: p.reduce((s, x) => s + x.lng, 0) / p.length } : { lat: null, lng: null };
}
function hav(a, b) {
  const R = Math.PI / 180, x = Math.sin((b.lat - a.lat) * R / 2) ** 2 + Math.cos(a.lat * R) * Math.cos(b.lat * R) * Math.sin((b.lng - a.lng) * R / 2) ** 2;
  return 12742 * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x));
}
const osm = (lat, lng, z) => 'https://www.openstreetmap.org/?lat=' + lat + '&lon=' + lng + '&zoom=' + (z || 12) + '&layers=M';

// ---------- Active Areas sheet: every registered HQ + all its coordinates ----------
// Rebuilt from readCurrent() after every change, then read back and checked against what the map reads.
function syncActiveAreas() {
  try {
    _cur = null;
    const hs = readCurrent().slice().sort((a, b) => (a.region || '').localeCompare(b.region || '') || a.name.localeCompare(b.name));
    const b = ss();
    const s = b.getSheetByName(AA) || b.insertSheet(AA);
    const head = ['HQ ID', 'HQ Name', 'Region', 'Points', 'Status', 'Latest Entry (UTC)', 'Entered By', 'Source', 'Centre Lat', 'Centre Lng'];
    for (let i = 1; i <= MAXP; i++) head.push('Point ' + i + ' Lat', 'Point ' + i + ' Lng');
    const rows = hs.map(h => {
      const c = centroid(h.points);
      const r = [h.id, h.name, h.region, h.points.length, h.points.length >= 3 ? 'On map' : 'Needs 3+ points',
        h.on ? h.on.slice(0, 16).replace('T', ' ') : '', h.by, h.source, c.lat === null ? '' : c.lat, c.lng === null ? '' : c.lng];
      for (let i = 0; i < MAXP; i++) r.push(h.points[i] ? h.points[i].lat : '', h.points[i] ? h.points[i].lng : '');
      return r;
    });

    s.getBandings().forEach(x => x.remove());
    s.clear();
    s.getRange(1, 1, 1, head.length).setValues([head]);
    if (rows.length) s.getRange(2, 1, rows.length, head.length).setValues(rows);

    const last = rows.length + 1, cols = head.length;
    s.getRange(1, 1, last, cols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    s.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#FFC700').setFontColor('#000000')
      .setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
    s.setRowHeight(1, 34);
    s.setFrozenRows(1); s.setFrozenColumns(2);
    if (rows.length) {
      s.getRange(2, 9, rows.length, cols - 8).setNumberFormat('0.000000').setHorizontalAlignment('right');
      s.getRange(2, 4, rows.length, 1).setHorizontalAlignment('center');
      s.getRange(2, 5, rows.length, 1).setFontColors(rows.map(r => [r[4] === 'On map' ? '#1B7F3B' : '#C62828'])).setFontWeight('bold');
    }
    s.autoResizeColumns(1, cols);
    for (let c = 1; c <= cols; c++) if (s.getColumnWidth(c) < 80) s.setColumnWidth(c, 80);
    s.getRange(1, 1).setNote('Auto-generated. Do not type here: it is rebuilt after every change.\nLast synced: ' + new Date().toISOString());

    // ---- post-sync check: read the sheet back and compare with what the map reads ----
    const back = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 4).getValues() : [];
    const problems = [];
    if (back.length !== hs.length) problems.push('row count ' + back.length + ' vs ' + hs.length + ' HQs');
    hs.forEach((h, i) => { if (!back[i] || back[i][1] !== h.name || Number(back[i][3]) !== h.points.length) problems.push('mismatch at ' + h.name); });
    if (problems.length) logError('syncActiveAreas', 'check failed', problems.join('; '));
    return { ok: !problems.length, rows: hs.length, problems };
  } catch (e) { logError('syncActiveAreas', e.message, ''); return { ok: false, rows: 0, problems: [e.message] }; }
}
function refreshActiveAreas() {
  const r = syncActiveAreas();
  try { SpreadsheetApp.getUi().alert(r.ok ? 'Active Areas updated: ' + r.rows + ' HQ(s), check passed.' : 'Active Areas check FAILED:\n' + r.problems.join('\n')); } catch (_) {}
}

// ---------- capture edits typed in the first tab (debounced) ----------
// onSheetEdit only records WHICH rows changed and (re)starts a 1-minute timer, so typing stays fast.
// commitPendingEdits runs once you stop typing: it saves a row as a new latest entry only if the row parses
// completely (3+ points, no missing lat/lng). Incomplete rows stay queued until you finish them.
const PEND = 'pendingRows', DEBOUNCE_MIN = 1;
function installTriggers() {
  ScriptApp.getProjectTriggers().filter(t => ['onSheetEdit', 'commitPendingEdits'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onSheetEdit').forSpreadsheet(ss()).onEdit().create();
  try { SpreadsheetApp.getUi().alert('Auto-capture installed. Edits in "' + (firstTab() ? firstTab().getName() : '?') + '" are saved about ' + DEBOUNCE_MIN + ' minute after you stop typing.'); } catch (_) {}
}
function onSheetEdit(e) {
  const lock = LockService.getScriptLock();
  try {
    const ft = firstTab();
    if (!ft || e.range.getSheet().getSheetId() !== ft.getSheetId()) return;
    lock.waitLock(5000);
    const props = PropertiesService.getScriptProperties(), rows = JSON.parse(props.getProperty(PEND) || '[]');
    const r1 = e.range.getRow(), r2 = Math.min(e.range.getLastRow(), r1 + 500);
    for (let r = r1; r <= r2; r++) if (rows.indexOf(r) < 0) rows.push(r);
    props.setProperty(PEND, JSON.stringify(rows.slice(-300)));
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'commitPendingEdits').forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('commitPendingEdits').timeBased().after(DEBOUNCE_MIN * 60 * 1000).create();
  } catch (x) { logError('onSheetEdit', x.message, ''); }
  finally { try { lock.releaseLock(); } catch (_) {} }
}
function commitPendingEdits() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'commitPendingEdits').forEach(t => ScriptApp.deleteTrigger(t));
    const props = PropertiesService.getScriptProperties(), rows = JSON.parse(props.getProperty(PEND) || '[]'), ft = firstTab();
    props.deleteProperty(PEND);
    if (ft && rows.length) {
      _cur = null;
      const nm = ft.getName(), p = parseTab(ft);
      const incomplete = r => p.warn.some(w => w.indexOf(nm + ' row ' + r + ':') === 0 || w.indexOf(nm + ' row ' + r + ' (') === 0);
      const keep = [];
      rows.forEach(r => { if (incomplete(r)) keep.push(r); });
      p.entries.filter(x => rows.indexOf(x.row) >= 0 && !incomplete(x.row)).forEach(x => {
        const cur = readCurrent().find(h => h.name.toLowerCase() === x.name.toLowerCase());
        if (!cur || JSON.stringify(cur.points) !== JSON.stringify(x.points))
          put(x.name, cur ? cur.region : '', x.points, 'Active', 'sheet edit: ' + nm + ' row ' + x.row);
      });
      if (keep.length) props.setProperty(PEND, JSON.stringify(keep)); // still unfinished: wait for your next edit
    }
    syncActiveAreas();
  } catch (x) { logError('commitPendingEdits', x.message, ''); }
  finally { try { lock.releaseLock(); } catch (_) {} }
}

// ---------- HQs ----------
function getAllHQs() {
  return readCurrent().map(h => { const c = centroid(h.points); return { hqId: h.id, hqName: h.name, hqAddress: '', lat: c.lat, lng: c.lng, region: h.region }; });
}
function addHQ(name, region) {
  try {
    name = String(name || '').trim(); if (!name) return { error: 'HQ name required' };
    if (readCurrent().some(h => h.name.toLowerCase() === name.toLowerCase())) return { error: 'HQ already exists' };
    put(name, region, []);
    const h = readCurrent().find(x => x.name.toLowerCase() === name.toLowerCase());
    syncActiveAreas();
    return { hqId: h.id, hqName: name };
  } catch (e) { return fail('addHQ', e, name); }
}
function editHQ(hqId, name, region) {
  try {
    const h = find(hqId); if (!h) return { error: 'HQ not found' };
    name = String(name || '').trim();
    if (name.toLowerCase() !== h.name.toLowerCase()) put(h.name, h.region, h.points, 'Inactive'); // rename = retire old name
    put(name, region, h.points);
    syncActiveAreas();
    return { success: true, hqId, hqName: name };
  } catch (e) { return fail('editHQ', e, hqId); }
}
function removeHQ(hqId) {
  try { const h = find(hqId); if (!h) return { error: 'HQ not found' }; put(h.name, h.region, h.points, 'Inactive'); syncActiveAreas(); return { success: true }; }
  catch (e) { return fail('removeHQ', e, hqId); }
}
function getHQHistory(hqId) { // every entry for one HQ, newest first
  const h = find(hqId); if (!h) return [];
  const s = ss().getSheetByName(E);
  return s.getRange(2, 1, s.getLastRow() - 1, 6).getValues().filter(r => String(r[0]).trim().toLowerCase() === h.name.toLowerCase())
    .reverse().slice(0, 15).map(r => ({ on: r[3] ? new Date(r[3]).toISOString() : '', by: r[4], source: r[5], status: r[2] }));
}

// ---------- Areas = the points of an HQ's territory (areaId = HQ001-P2) ----------
function getAreasForHQ(hqId) {
  const h = find(hqId); if (!h) return [];
  return h.points.map((p, i) => ({ areaId: hqId + '-P' + (i + 1), hqId, hqName: h.name, areaName: 'Point ' + (i + 1), lat: p.lat, lng: p.lng, type: 'Point' }));
}
function addAreaToHQ(hqId, areaName, areaType, notes, lat, lng) {
  try {
    const h = find(hqId); if (!h) return { error: 'HQ not found: ' + hqId };
    const c = parseManualCoords(lat, lng); if (!c) return { error: 'Please provide valid coordinates.' };
    if (h.points.length >= MAXP) return { error: 'An HQ can have at most ' + MAXP + ' points.' };
    put(h.name, h.region, h.points.concat([c]));
    syncActiveAreas();
    return { areaId: hqId + '-P' + (h.points.length + 1), lat: c.lat, lng: c.lng, formattedAddress: areaName };
  } catch (e) { return fail('addAreaToHQ', e, hqId); }
}
function editArea(areaId, areaName, areaType, notes, lat, lng) { // moves one point = a new latest entry
  try {
    const m = /^(HQ\d+)-P(\d+)$/.exec(areaId), h = m && find(m[1]); if (!h) return { error: 'Area not found' };
    const c = parseManualCoords(lat, lng); if (!c) return { error: 'Please provide valid coordinates.' };
    const pts = h.points.slice(), i = +m[2] - 1; if (!pts[i]) return { error: 'Point not found' };
    pts[i] = c; put(h.name, h.region, pts);
    syncActiveAreas();
    return { success: true, areaId, areaName: 'Point ' + (i + 1) };
  } catch (e) { return fail('editArea', e, areaId); }
}
function removeArea(areaId) {
  try {
    const m = /^(HQ\d+)-P(\d+)$/.exec(areaId), h = m && find(m[1]); if (!h) return { error: 'Area not found' };
    const pts = h.points.slice(); pts.splice(+m[2] - 1, 1); put(h.name, h.region, pts);
    syncActiveAreas();
    return { success: true };
  } catch (e) { return fail('removeArea', e, areaId); }
}

// ---------- boundary, distances, map ----------
// The polygon is the points exactly as entered (your sheet defines explicit "New Map" vertices) - no hull, no padding.
function dists(p) {
  const n = p.length, o = [];
  for (let i = 0; i < n; i++) { const j = (i + 1) % n; if (n < 3 && j < i) break; o.push({ pointA: i + 1, pointB: j + 1, distance: hav(p[i], p[j]), type: 'EDGE' }); }
  for (let i = 0; i < n; i++) for (let j = i + 2; j < n; j++) { if (i === 0 && j === n - 1) continue; o.push({ pointA: i + 1, pointB: j + 1, distance: hav(p[i], p[j]), type: 'DIAGONAL' }); }
  return o;
}
function getDistancesForHQ(hqId) {
  try {
    const h = find(hqId); if (!h) return { error: 'HQ not found' };
    if (h.points.length < 3) return { error: h.name + ' needs at least 3 points' };
    const d = dists(h.points), edges = d.filter(x => x.type === 'EDGE'), dg = d.filter(x => x.type === 'DIAGONAL');
    return { hqId, hqName: h.name, pointCount: h.points.length, edges, diagonals: dg,
      perimeter: +edges.reduce((s, x) => s + x.distance, 0).toFixed(2), maxDiagonal: +dg.reduce((m, x) => Math.max(m, x.distance), 0).toFixed(2) };
  } catch (e) { return fail('getDistancesForHQ', e, hqId); }
}
function getMapData() {
  try {
    const hs = readCurrent().filter(h => h.points.length >= 3);
    if (!hs.length) return { error: 'No territories with 3+ points yet' };
    const data = hs.map((h, i) => {
      const c = centroid(h.points);
      return { hqId: h.id, hqName: h.name,
        hqAddress: [h.region, 'latest entry ' + (h.on ? h.on.slice(0, 10) + ' by ' + (h.by || '-') : h.source), h.entries + ' entr' + (h.entries > 1 ? 'ies' : 'y')].filter(Boolean).join(' | '),
        hqLatLng: c, boundaryPolygon: h.points, color: HQ_COLORS[i % HQ_COLORS.length], areaCount: h.points.length, osmLink: osm(c.lat, c.lng, 13),
        areaCentroids: h.points.map((p, j) => ({ name: 'Point ' + (j + 1), lat: p.lat, lng: p.lng, type: 'Point', index: j + 1 })) };
    });
    const cen = centroid(data.map(d => d.hqLatLng));
    return { data, center: cen, osmLink: osm(cen.lat, cen.lng, 11) };
  } catch (e) { return fail('getMapData', e); }
}

// ---------- Import tab (file upload): each HQ's rows become ONE new latest entry that replaces its old points ----------
function importParsedRows(rows) {
  try {
    initSheets();
    if (!Array.isArray(rows) || !rows.length) return { error: 'No data provided for import' };
    const sum = { hqsCreated: 0, areasCreated: 0, errors: [], warnings: [] }, g = {}, ord = [];
    const exist = new Set(readCurrent().map(h => h.name.toLowerCase()));
    rows.forEach((r, i) => {
      const n = String(r.hqName || '').trim(), c = parseManualCoords(r.lat, r.lng);
      if (!n || !c) { sum.warnings.push('Row ' + (i + 2) + ': skipped (missing HQ name or invalid coordinates)'); return; }
      if (!g[n]) { g[n] = { region: String(r.region || '').trim(), pts: [] }; ord.push(n); }
      g[n].pts.push(c);
    });
    ord.forEach(n => {
      const x = g[n];
      if (x.pts.length > MAXP) { sum.warnings.push(n + ': only the first ' + MAXP + ' points were kept'); x.pts = x.pts.slice(0, MAXP); }
      put(n, x.region, x.pts, 'Active', 'import');
      if (!exist.has(n.toLowerCase())) sum.hqsCreated++;
      sum.areasCreated += x.pts.length;
    });
    syncActiveAreas();
    return { success: true, summary: sum, message: 'Imported ' + sum.hqsCreated + ' new HQ(s) and ' + sum.areasCreated + ' point(s). Each HQ\'s newest entry replaces its old points.' };
  } catch (e) { return fail('importParsedRows', e); }
}

// ---------- Reading your existing tabs ----------
// A header row has "Latitude"/"Longitude" cells; the row above has "Point N" labels. Rows below it with 4+ numbers are
// entries, in sheet order. A name in the table labels every row of that table; the LAST row is the latest entry.
function firstTab() { return ss().getSheets().find(sh => !NON_SOURCE.includes(sh.getName())); }
function sourceSheets() { const f = firstTab(); return f ? [f] : []; }
function parseTab(sheet) {
  const v = sheet.getDataRange().getValues(), out = [], warn = [], tab = sheet.getName();
  let cols = null, block = [];
  const num = x => { if (typeof x === 'number') return x; const t = String(x).trim().replace(/,$/, ''); return /^-?\d+(\.\d+)?$/.test(t) ? parseFloat(t) : NaN; };
  const flush = () => {
    const labels = [...new Set(block.filter(b => b.lbl).map(b => b.lbl))];
    let last = '';
    block.forEach(b => {
      const nm = b.lbl || (labels.length === 1 ? labels[0] : last);
      if (b.lbl) last = b.lbl;
      if (!nm) { warn.push(tab + ' row ' + b.row + ': no area name found'); return; }
      const nums = Object.keys(b.pts).map(Number).sort((a, c) => a - c), pts = [];
      for (let k = 1; k <= nums[nums.length - 1] && pts.length < MAXP; k++) {
        const p = b.pts[k];
        if (!p || p.lat === undefined || p.lng === undefined) { warn.push(tab + ' row ' + b.row + ' (' + nm + '): Point ' + k + ' is missing its ' + (p && p.lat === undefined ? 'latitude' : 'longitude')); break; }
        pts.push(p);
      }
      if (pts.length < 3) { warn.push(tab + ' row ' + b.row + ' (' + nm + '): skipped, fewer than 3 complete points'); return; }
      out.push({ name: nm, points: pts, row: b.row, tab });
    });
    block = [];
  };
  for (let r = 0; r < v.length; r++) {
    const row = v[r], low = row.map(c => String(c).trim().toLowerCase());
    if (low.some(c => c === 'latitude' || c === 'longitude')) {
      flush();
      const lab = r > 0 ? v[r - 1].map(c => String(c).trim().toLowerCase()) : [];
      cols = [];
      low.forEach((c, i) => {
        if (c !== 'latitude' && c !== 'longitude') return;
        let pt = 1;
        for (let k = i; k >= 0; k--) { const m = /^point\s*(\d+)/.exec(lab[k] || ''); if (m) { pt = +m[1]; break; } }
        cols.push({ i, pt, k: c === 'latitude' ? 'lat' : 'lng' });
      });
      // a lone "Longitude" for Point 1: take its latitude from the cell just to the left, if that cell holds a number
      const first = cols.filter(c => c.pt === 1);
      if (first.length === 1 && first[0].k === 'lng' && first[0].i > 0) cols.push({ i: first[0].i - 1, pt: 1, k: 'lat' });
      continue;
    }
    if (!cols) continue;
    const minC = Math.min.apply(null, cols.map(c => c.i)), pts = {};
    let got = 0;
    cols.forEach(c => { const x = num(row[c.i]); if (!isNaN(x)) { (pts[c.pt] = pts[c.pt] || {})[c.k] = x; got++; } });
    const lbl = row.slice(0, minC).map(c => String(c).trim()).find(t => t && isNaN(num(t)) && !/^(point|diagonal|distance|new map)/i.test(t) && !/km$/i.test(t));
    if (got >= 4) block.push({ row: r + 1, pts, lbl: lbl || '' });
  }
  flush();
  return { entries: out, warn };
}
// Menu: Territory > Check what the map reads (also rebuilds Active Areas and verifies it)
function checkSheet() {
  const lines = [], warn = [];
  sourceSheets().forEach(sh => {
    const r = parseTab(sh);
    lines.push(sh.getName() + ': ' + r.entries.length + ' entr' + (r.entries.length === 1 ? 'y' : 'ies') + ' read');
    r.warn.forEach(w => warn.push(w));
  });
  const hs = readCurrent(), sync = syncActiveAreas();
  SpreadsheetApp.getUi().alert(lines.join('\n') + '\n\nAreas on the map (3+ points): ' + hs.filter(h => h.points.length >= 3).length + ' of ' + hs.length +
    '\nActive Areas sheet: ' + (sync.ok ? sync.rows + ' HQ(s), check passed' : 'CHECK FAILED - ' + sync.problems.join('; ')) +
    (warn.length ? '\n\nNot read (' + warn.length + '):\n' + warn.slice(0, 15).join('\n') + (warn.length > 15 ? '\n...' : '') : ''));
}