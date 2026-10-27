/** NBD generator
 * Usage: npm install exceljs; node generate.js path/to/checklist.xlsx [identities.json]
 * Output only (./output): <Sheet>.json, <Sheet>.validation.txt, validation.json. Never writes to MongoDB.
 *
 * identities.json (optional, keyed by sheet name) - values the checklist does not contain:
 * { "Shayota": {
 *     "person_id": "KOL_POC_1016", "full_name": "Brian Jacob Shayota",
 *     "scores": { "advocacy_score": {"value": 10, "avg": 2.5}, "bibliometric_score": {"value": 10, "avg": 1.8} },
 *     "events": [ { "event_start_date": "08/01/2026", "event_end_date": "08/30/2026" } ] } }
 */
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const workbookFile = process.argv[2] || path.join(__dirname, 'NBD_Zevra_Miplyffa_2_HCPs_v2.1.xlsx');
const identityFile = process.argv[3] || path.join(__dirname, 'identities.json');
const identities = fs.existsSync(identityFile) ? JSON.parse(fs.readFileSync(identityFile, 'utf8')) : {};
const outputDir = path.join(__dirname, 'output'); fs.mkdirSync(outputDir, {recursive: true});

const headings = {
  'insights':'insight', 'insight':'insight', 'how to connect':'how_to_connect',
  'where to connect':'when_to_connect', 'when to connect':'when_to_connect',
  'strategic recommendation':'strategic_recommendations', 'strategic recommendations':'strategic_recommendations',
  'what to discuss':'what_to_discuss', 'action':'action'
};
// Strategic Recommendation label -> score key + platform defaults (copied from the verified config; override per person in identities.json).
const SCORES = {
  centrality: {key:'advocacy_score', value:10, avg:2.5, tooltip:'The expert can be considered for investigator-sponsored or collaborative research given their trial leadership. '},
  reputation: {key:'bibliometric_score', value:10, avg:1.8, tooltip:'The expert has a high overall score and can be considered for scientific advisory board, one-on-one exchanges, and Expertise Consulting.'}
};
const MONTHS = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
const MON_RE = '(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*';
const B0='\u0001', B1='\u0002', I0='\u0003', I1='\u0004'; // private markers, turned into tags last

const clean = v => String(v ?? '').trim();
const nl = s => String(s ?? '').replace(/\r\n?/g, '\n');
const esc = s => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const anchor = (url, label = 'Source') => `<A href='${url.replace(/'/g, '%27')}' target='_blank'>${label}</A>`;
// "(url)", "Link: url" or a bare url -> one <a>Source</a> glued to the previous text (no space, no parentheses).
const URL_G = () => /[ \t]*\n?[ \t]*(?:Link:\s*)?(?:\((https?:\/\/[^\s)]+)\)|(https?:\/\/[^\s)]+))/g;

function cellText(cell) {
  const value = cell.value;
  if (value == null) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) return value.richText.map(p => p.text || '').join('');
    if ('text' in value) return String(value.text ?? '');
    if ('result' in value) return String(value.result ?? '');
  }
  return String(value);
}
function segments(cell) {
  const v = cell.value; if (v == null) return [];
  const cb = !!cell.font?.bold, ci = !!cell.font?.italic;
  if (Array.isArray(v.richText)) return v.richText.map(r => ({
    text: nl(r.text ?? ''), bold: Boolean(r.font ? r.font.bold : cb), italic: Boolean(r.font ? r.font.italic : ci)}));
  return [{text: nl(cellText(cell)), bold: cb, italic: ci}];
}
// Wrap each line's core text; leading/trailing spaces and the line breaks stay OUTSIDE the tag: <b>Title</b><br />
function wrapLines(text, open, close) {
  return text.replace(/[^\n]+/g, line => {
    const m = line.match(/^(\s*)([\s\S]*?)(\s*)$/);
    return m[2] ? m[1] + open + m[2] + close + m[3] : line;
  });
}
function marked(cell) {
  const segs = [];
  for (const s of segments(cell)) {
    const last = segs[segs.length - 1];
    if (last && last.bold === s.bold && last.italic === s.italic) last.text += s.text; else segs.push({...s});
  }
  return segs.map(s => {
    let t = esc(s.text);
    if (s.bold) t = wrapLines(t, B0, B1);
    if (s.italic) t = wrapLines(t, I0, I1);
    return t;
  }).join('');
}
function toHtml(m, trimEnd = true, linkLabel = 'Source') {
  let s = m.replace(URL_G(), (_, u1, u2) => anchor(u1 || u2, linkLabel));
  s = s.replace(/(?:[ \t]*\n[ \t]*)+/g, x => (x.match(/\n/g).length > 1 ? '<BR /><BR/>' : '<BR />'));
  s = s.replace(/^(?:\s|<BR ?\/>)+/, '');
  if (trimEnd) s = s.replace(/(?:\s|<BR ?\/>)+$/, '');
  s = s.replace(/(?:<BR ?\/>)+(?=<UL>)/g, '').replace(/(<\/UL>)(?:<BR ?\/>)+/g, '$1'); // no <BR> around lists
  return s.replace(/\u0001/g,'<B>').replace(/\u0002/g,'</B>').replace(/\u0003/g,'<I>').replace(/\u0004/g,'</I>');
}
const paragraphs = text => nl(text).split(/\n[ \t]*\n+/).map(p => p.trim()).filter(Boolean);

// Profile links start with "/main/..." so they resolve from the site root.
function toRelative(url, w) {
  try {
    const u = new URL(url);
    if (/neoexpert/i.test(u.hostname)) {
      if (!/\/summary\/?$/.test(u.pathname)) w.push(`How to Connect: link '${u.pathname}' does not end with /summary; confirm the target.`);
      return '/' + (u.pathname + u.search + u.hash).replace(/^\/+/, '');
    }
  } catch (e) {}
  return url.replace(/^\/*(?=main\/)/, '/');
}
// Doctors' names (Dr. First Last) -> <b>; skips text that is already bold.
function boldNames(m) {
  const NAME = /\bDr\.?\s+[A-Z][\w.'’-]*(?:\s+[A-Z][\w.'’-]*){0,3}/g;
  return m.split(/(\u0001[\s\S]*?\u0002)/).map(part => part.startsWith('\u0001') ? part : part.replace(NAME, x => `\u0001${x}\u0002`)).join('');
}
// Lines starting with "- ", "* " or a bullet become <ul><li>..</li></ul> (Insights/Action only).
function listify(m) {
  const out = [], lines = m.split('\n'); let items = [];
  const flush = () => { if (items.length) { out.push('<UL>' + items.map(x => `<LI>${x}</LI>`).join('') + '</UL>'); items = []; } };
  for (const ln of lines) { const b = ln.match(/^\s*[-*•]\s+(.+)$/); if (b) items.push(b[1].trim()); else { flush(); out.push(ln); } }
  flush(); return out.join('\n');
}
function parseHowTo(text, w) {
  const t = nl(text).replace(/\s+/g, ' ').trim();
  let m = t.match(/^(.*?)\s*\((https?:\/\/[^\s)]+)\)\s*(.*)$/);
  if (!m) {
    // Profile link is often not in the checklist: keep the name hyperlinked with an EMPTY href to fill in manually.
    const n = t.match(/^(Leverage\s+)?(.+?)\s+((?:who|to|and|as|with)\b.*)$/);
    if (!n) { w.push('How to Connect: could not find the connector name; message needs manual review.'); return esc(t); }
    w.push('How to Connect: no profile link in the checklist; href left empty (add /main/projects/<project>/kols/<id>/summary manually).');
    m = [t, (n[1] || '') + n[2], '', n[3]];
  }
  const lead = m[1].match(/^(Leverage\s+)(.+)$/);
  const intro = lead ? lead[1] : '', name = lead ? lead[2] : m[1];
  const sp = m[3].match(/^(.*?\bas they (?:have both participated in|both have|have both|both|have))\s+(.*)$/i);
  let tail;
  if (sp) {
    const items = sp[2].split(/;\s*/).map(x => x.replace(/^and\s+/i, '').trim()).filter(Boolean);
    tail = `${esc(sp[1])}<BR/><UL>${items.map(i => `<LI>${esc(i)}</LI>`).join('')}</UL>`;
  } else {
    w.push('How to Connect: could not detect the "as they have ..." list; text left unformatted.');
    tail = esc(m[3]);
  }
  const href = m[2] ? toRelative(m[2], w).replace(/'/g, '%27') : '';
  return `${esc(intro)}<A href='${href}' target='_blank'><B>${esc(name)}</B></A> ${tail}`;
}
function parseDates(sfx) {
  const mm = sfx.match(new RegExp(`\\b(${MON_RE})\\.?(?:\\s+(\\d{1,2})(?:\\s*[-–]\\s*(\\d{1,2}))?(?!\\d))?`, 'i'));
  const ym = sfx.match(/\b(20\d\d)\b/);
  if (!mm || !ym) return null;
  const mo = MONTHS.indexOf(mm[1].slice(0, 3).toLowerCase()) + 1, y = +ym[1], p = n => String(n).padStart(2, '0');
  const f = d => `${p(mo)}/${p(d)}/${y}`;
  // Docs: date and event_end_date = the actual date in the sheet; event_start_date = a date BEFORE it.
  if (mm[2] && mm[3]) return {start: f(+mm[2]), end: f(+mm[3]), estimated: false};        // range: first day -> last day
  if (mm[2]) { const b = new Date(y, mo - 1, +mm[2] - 1);                                  // single day: start = day before
    return {start: `${p(b.getMonth() + 1)}/${p(b.getDate())}/${b.getFullYear()}`, end: f(+mm[2]), estimated: false}; }
  return {start: f(1), end: f(new Date(y, mo, 0).getDate()), estimated: true};            // month only
}
const good = s => s.replace(/(?<!<B>)Good news!/g, '<B>Good news!</B>'); // docs: "Good news!" is always bold
function parseWhere(text, w, ov) {
  const events = [], other = [];
  for (const p0 of paragraphs(text)) {
    const p = p0.replace(/\s+/g, ' ');
    const um = p.match(/\s*(?:\((https?:\/\/[^\s)]+)\)|(https?:\/\/[^\s)]+))/);
    if (!um) { other.push(p); continue; }
    const url = um[1] || um[2], txt = p.replace(um[0], '').trim();
    const m = txt.match(new RegExp(`^(.*?)(\\b20\\d\\d\\b.+?)(\\s+(?:in|on|from|at)\\s+${MON_RE}\\b.*)$`, 'i'));
    let name, d = null;
    if (m) { name = `${esc(m[1])}<B>${esc(m[2])}</B>${esc(m[3])}`; d = parseDates(m[3]); }
    else { name = esc(txt); w.push(`Where to Connect: could not locate the event title/date in "${txt}"; dates left empty.`); }
    const ev = {name: good(name), date: d ? d.end : null, event_start_date: d ? d.start : null, event_part_date: d ? d.start : null,
                event_end_date: d ? d.end : null, url, updated_events: true};
    Object.assign(ev, (ov.events || [])[events.length] || {});
    if (d && d.estimated && !(ov.events || [])[events.length]?.event_end_date)
      w.push(`Where to Connect: only month/year given for "${txt.slice(0, 40)}..."; event_end_date set to month end (${d.end}). Verify, or set it in identities.json.`);
    events.push(ev);
  }
  const content = other.map(p => { const a = p.match(/^(.*?\bat)\s+(\d.*)$/); return good(a ? `${esc(a[1])}<BR/> ${esc(a[2])}` : esc(p)); }).join('<BR /><BR/>');
  return {content: content || null, events};
}
function parseStrategic(text, w, ov) {
  const out = [];
  for (const p of paragraphs(text)) {
    const lines = p.split('\n').map(x => x.trim()).filter(Boolean);
    if (lines.length < 2) w.push(`Strategic Recommendation: "${lines[0]}" has no description line.`);
    const label = lines[0], def = SCORES[label.toLowerCase()];
    if (!def) w.push(`Strategic Recommendation: label '${label}' has no score mapping (key/value/avg/tooltip left generic); add it to SCORES.`);
    const key = def ? def.key : label.toLowerCase().replace(/\W+/g, '_') + '_score';
    // Docs: only label and text change per sheet; key/value/avg/max/enabled/tooltip come from the copied 6171 NBD collection (defaults below, optional override).
    const o = (ov.scores || {})[key] || {};
    const value = o.value ?? (def ? def.value : null), avg = o.avg ?? (def ? def.avg : null);
    const text = boldNames(esc(lines.slice(1).join(' '))).replace(/\u0001/g, '<B>').replace(/\u0002/g, '</B>');
    out.push({key, value, label, avg, max: true, enabled: true, tooltip: def ? def.tooltip : '', text});
  }
  return out;
}

function generateSheet(sheetName, ws) {
  const rows = [];
  for (let i = 1; i <= ws.rowCount; i++) {
    const row = ws.getRow(i), cells = [];
    for (let j = 1; j <= ws.columnCount; j++) cells.push(cellText(row.getCell(j)));
    rows.push(cells);
  }
  const warnings = []; let headerRow = -1, colSection = -1, colContent = -1;
  rows.forEach((r, i) => { if (headerRow >= 0) return; const a = r.map(c => clean(c).toLowerCase()); const x = a.indexOf('section'), y = a.indexOf('content'); if (x >= 0 && y >= 0) { headerRow = i; colSection = x; colContent = y; } });
  if (headerRow < 0) throw new Error(`${sheetName}: Missing Section and Content headers`);
  const name = clean(rows[1]?.[colSection]) || sheetName;
  const override = identities[sheetName] || {};
  const personId = clean(override.person_id || rows[2]?.[colSection]);
  const fullName = clean(override.full_name || name.replace(/\s*\([^)]*\)\s*$/, '').replace(/,\s*(MD|MPH|PhD|DO|MBBS)(\s*,\s*(MD|MPH|PhD|DO|MBBS))*/gi, ''));
  if (!personId) warnings.push('Missing person_id. Supply in identities.json; do not deploy until resolved.');
  if (!override.full_name && /\b[A-Z]\.(\s|$)/.test(fullName)) warnings.push(`full_name '${fullName}' contains an initial; set the full name in identities.json if the platform uses it.`);
  const nbd = {insight:{Content:null,AdditionalData:null}, how_to_connect:{Content:null,AdditionalData:[]}, when_to_connect:{Content:null,AdditionalData:[]},
               strategic_recommendations:{Content:'',AdditionalData:[]}, what_to_discuss:{Content:null,AdditionalData:[]}, action:{Content:null,AdditionalData:null}};
  const seen = new Set();
  for (let i = headerRow + 1; i < rows.length; i++) {
    const label = clean(rows[i]?.[colSection]), content = clean(rows[i]?.[colContent]);
    if (!label) continue;
    const key = headings[label.toLowerCase()];
    if (!key) { warnings.push(`Unknown section '${label}' at row ${i + 1}`); continue; }
    if (seen.has(key)) warnings.push(`Repeated section '${label}' at row ${i + 1}`);
    seen.add(key);
    if (!content) { warnings.push(`Empty content for ${label}`); continue; }
    const cell = ws.getRow(i + 1).getCell(colContent + 1), m = marked(cell);
    if (key === 'insight') nbd[key].Content = toHtml(listify(boldNames(m)), true, 'LINK');
    if (key === 'action') nbd[key].Content = toHtml(listify(m));
    if (key === 'how_to_connect') nbd[key].AdditionalData = [{content: 'External Connection', additionalData: {message: parseHowTo(content, warnings)}}];
    if (key === 'when_to_connect') { const r = parseWhere(content, warnings, override); nbd[key].Content = r.content; nbd[key].AdditionalData = r.events; }
    if (key === 'strategic_recommendations') nbd[key].AdditionalData = parseStrategic(content, warnings, override);
    if (key === 'what_to_discuss') {
      nbd[key].AdditionalData = m.split(/\n[ \t]*\n\s*/).map(b => b.replace(/^\s+/, '').replace(/\n+$/, '')).filter(b => b.trim())
        .map(b => toHtml(b, false));
      nbd[key].AdditionalData.forEach((h, n) => { if (!/^<B>[^<]+<\/B><BR \/>/.test(h)) warnings.push(`What to Discuss: item ${n + 1} does not start with a bold title followed by a line break (check bold/line breaks in the checklist).`); });
    }
  }
  Object.keys(nbd).filter(k => !seen.has(k)).forEach(k => warnings.push(`Missing required section: ${k}`));
  const document = {person_id: personId || null, next_best_dialogue: nbd, full_name: fullName};
  // MongoDB generates _id on insert; never copy another document's _id.
  const stem = sheetName.replace(/[^\w-]+/g, '_');
  fs.writeFileSync(path.join(outputDir, `${stem}.json`), JSON.stringify(document, null, 2));
  fs.writeFileSync(path.join(outputDir, `${stem}.validation.txt`), warnings.length ? warnings.map(x => '- ' + x).join('\n') : 'No warnings.');
  return {sheet: sheetName, person_id: personId || null, full_name: fullName, warnings};
}
async function main() {
  if (!fs.existsSync(workbookFile)) throw new Error(`Excel file not found: ${workbookFile}`);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(workbookFile);
  const report = [];
  for (const worksheet of workbook.worksheets) {
    try { report.push(generateSheet(worksheet.name, worksheet)); }
    catch (e) { report.push({sheet: worksheet.name, error: e.message}); }
  }
  fs.writeFileSync(path.join(outputDir, 'validation.json'), JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report, null, 2));
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
