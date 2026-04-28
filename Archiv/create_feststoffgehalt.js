const XLSX = require('xlsx');
const fs   = require('fs');

const wb = XLSX.readFile('C:/Users/monak/projects/ligarolabui/solid_content.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const raw = XLSX.utils.sheet_to_json(ws, { header: 1 }).slice(1).filter(r => r[0]);

// Experiment-ID vom Rest des Sample-ID trennen
// z.B. "TEST-002 Kontrolle" → { expId: "TEST-002", probe: "Kontrolle" }
function parseSampleId(raw) {
  const s = String(raw).trim();
  const m = s.match(/^([A-Z]{2,5}-\d{2,3}(?:-\d{2,3})*)(.*)/);
  if (!m) return { expId: s, probe: '' };
  return { expId: m[1], probe: m[2].trim() };
}

function toCSV(arr) {
  const keys = Object.keys(arr[0]);
  const esc = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
  return [keys.join(','), ...arr.map(row => keys.map(k => esc(row[k])).join(','))].join('\r\n');
}

const rows = raw.map(r => {
  const { expId, probe } = parseSampleId(r[0]);
  return {
    Experiment_ID:  expId,
    Probe:          probe,
    Leergewicht_g:  r[1] != null ? r[1] : '',
    Einwaage_g:     r[2] != null ? r[2] : '',
    Endgewicht_g:   r[3] != null ? r[3] : '',
    Kommentar:      '',
  };
});

fs.writeFileSync(
  'C:/Users/monak/projects/ligarolabui/Feststoffgehalt.csv',
  '\uFEFF' + toCSV(rows), 'utf8'
);

console.log('Feststoffgehalt.csv:', rows.length, 'Einträge');
console.log('Spalten:', Object.keys(rows[0]).join(', '));
console.log('\nBeispiele:');
rows.slice(0, 6).forEach(r =>
  console.log(` ${r.Experiment_ID} | Probe: "${r.Probe}" | Leergewicht: ${r.Leergewicht_g} | Einwaage: ${r.Einwaage_g} | Endgewicht: ${r.Endgewicht_g}`)
);

// Duplikate prüfen
const multi = {};
rows.forEach(r => { const k = r.Experiment_ID + '|' + r.Probe; multi[k] = (multi[k]||0)+1; });
const dupes = Object.entries(multi).filter(([,v])=>v>1);
if (dupes.length) { console.log('\nMehrfachmessungen:'); dupes.forEach(([k])=>console.log(' ',k)); }
