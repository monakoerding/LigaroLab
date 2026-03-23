const XLSX = require('xlsx');
const fs = require('fs');

const wb = XLSX.readFile('C:/Users/monak/projects/ligarolabui/ExperimenteListe.xlsx');

// ── Hilfsfunktionen ──────────────────────────────────────────────────────────

function excelDate(s) {
  if (!s || typeof s !== 'number') return '';
  return new Date((s - 25569) * 86400000).toISOString().split('T')[0];
}

function toCSV(arr) {
  if (!arr || arr.length === 0) return '';
  const keys = Object.keys(arr[0]);
  const esc = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
  return [keys.join(','), ...arr.map(row => keys.map(k => esc(row[k])).join(','))].join('\r\n');
}

// ── Personen ─────────────────────────────────────────────────────────────────

const personenMap = {
  LL: { Vorname: 'Leander', Nachname: 'Lehmann' },
  TG: { Vorname: 'Tim',     Nachname: 'Gatz' },
  MK: { Vorname: 'Mona',    Nachname: 'Körding' },
  FA: { Vorname: 'Florentine', Nachname: 'Adam' },
  VM: { Vorname: 'Victor',  Nachname: 'Mayerhofer' },
};

const personen = Object.entries(personenMap).map(([kuerzel, p]) => ({
  Kuerzel:  kuerzel,
  Vorname:  p.Vorname,
  Nachname: p.Nachname,
}));

fs.writeFileSync('C:/Users/monak/projects/ligarolabui/Personen.csv', '\uFEFF' + toCSV(personen), 'utf8');
console.log('Personen.csv:', personen.length, 'Einträge');

// ── Experiment-ID Präfixe als Projektkürzel ──────────────────────────────────
// 306 unique "Projekt"-Freitexte → nicht als Lookup geeignet.
// Stattdessen: ID-Präfix als Projektkürzel, Projektname als freies Feld behalten.

const ws1 = wb.Sheets['Experimente'];
const expRaw = XLSX.utils.sheet_to_json(ws1, {header:1}).slice(1).filter(r => r[0]);

const prefixNamen = {};
expRaw.forEach(r => {
  const prefix = r[0].replace(/-.*/, '');
  if (!prefixNamen[prefix]) prefixNamen[prefix] = [];
  const proj = (r[3] && r[3] !== 'undefined') ? String(r[3]).trim() : '';
  if (proj) prefixNamen[prefix].push(proj);
});

// Manuelle Beschreibungen (vom Team bestätigt)
const PROJEKT_BESCHREIBUNGEN = {
  TEST: 'Allgemeine Tests',
  PHOK: 'Phosphorylierung Kartoffelprotein säuregefällt',
  PHOG: 'Phosphorylierung Protexceed',
  OTAM: 'Oxidiertes Tannin mit Metallen',
  ALHY: 'Alkalische Hydrolyse',
  AVGC: 'Anorganische Vernetzung Protexceed mit Calcium-Salzen',
  PHOH: 'Phosphatangereicherte Hefe',
  ENHY: 'Enzymatische Hydrolyse',
  CROX: 'Crosslinking mit oxidiertem Tannin',
  PHON: 'Phosphorylierung natives Protein (S020)',
};

const projekte = Object.entries(prefixNamen).map(([prefix, namen]) => {
  // Manuelle Beschreibung bevorzugen, sonst häufigsten Namen als Fallback
  let beschreibung = PROJEKT_BESCHREIBUNGEN[prefix];
  if (!beschreibung) {
    const freq = {};
    namen.forEach(n => { freq[n] = (freq[n] || 0) + 1; });
    const top = Object.entries(freq).sort((a, b) => b[1] - a[1])[0];
    beschreibung = top ? top[0] : '';
  }
  return {
    Projekt_Kuerzel: prefix,
    Beschreibung:    beschreibung,
  };
}).sort((a, b) => a.Projekt_Kuerzel.localeCompare(b.Projekt_Kuerzel));

fs.writeFileSync('C:/Users/monak/projects/ligarolabui/Projektkuerzel.csv', '\uFEFF' + toCSV(projekte), 'utf8');
console.log('Projektkuerzel.csv:', projekte.length, 'Einträge');

// ── Experimente ──────────────────────────────────────────────────────────────
// Optimierungen:
// - Datum: Excel-Seriennummer → ISO-Datum
// - Redundante Aggregatspalten bleiben für Initialimport, werden per Flow überschrieben
// - Spaltenname Standardabweichung → Dry_StdAbw / Wet_StdAbw
// - Projekt-Kürzel aus ID extrahiert als eigene Spalte

function parentId(id) {
  const parts = id.split('-');
  return parts.length > 2 ? parts.slice(0, -1).join('-') : '';
}

const experimente = expRaw.map(r => ({
  Experiment_ID:   r[0],
  Parent_ID:       parentId(r[0]),
  Projekt_Kuerzel: r[0].replace(/-.*/, ''),
  Datum:           excelDate(r[1]),
  Person_Kuerzel:  r[2] || '',
  Projekttitel:    (r[3] && r[3] !== 'undefined') ? String(r[3]).trim() : '',
  Beschreibung:    r[4] || '',
  Beobachtungen:   r[5] || '',
  Kommentar:       r[6] || '',
}));

fs.writeFileSync('C:/Users/monak/projects/ligarolabui/Experimente.csv', '\uFEFF' + toCSV(experimente), 'utf8');
console.log('\nExperimente.csv:', experimente.length, 'Einträge');
console.log('Datum-Beispiele:', experimente.slice(0,3).map(e => e.Datum));

// ── Komponenten ──────────────────────────────────────────────────────────────
// Optimierungen:
// - Rolle normalisieren: fehlerhafte Werte → 'Sonstiges' (mit Originalwert in Kommentar)
// - Menge in Zahl + Einheit trennen

const GUELTIGE_ROLLEN = new Set([
  'Protein', 'Vernetzer', 'Lösungsmittel', 'Puffer',
  'Crosslinker', 'Oxidationsmittel', 'Quench', 'Feststoff',
  'Kontrolle', 'Phosphorylierungsreagenz', 'Additiv', 'Sonstiges',
]);

// Manuelle Korrekturen: Schlüssel = "Experiment_ID|Komponente"
const ROLLEN_KORREKTUREN = {
  'TEST-003|Glycerol':              'Additiv',
  'TEST-004|Glycerol':              'Additiv',
  'TEST-012-04|Tannin 1':          'Crosslinker',
  'TEST-035-03|OTAM 2':            'Crosslinker',
  'TEST-019-07|Kaliumphosphatpuffer': 'Lösungsmittel',
};

function normRolle(rolle, expId, komponente) {
  // Manuelle Korrekturen zuerst prüfen
  const key = expId + '|' + komponente;
  if (ROLLEN_KORREKTUREN[key]) return ROLLEN_KORREKTUREN[key];
  if (!rolle) return '';
  // Synonyme zusammenführen
  if (rolle === 'Phosphorylierungsmittel') return 'Phosphorylierungsreagenz';
  if (rolle === 'Phosphoryliertes Protein') return 'Protein';
  if (rolle === 'höhere Flexibilität') return 'Additiv';
  if (GUELTIGE_ROLLEN.has(rolle)) return rolle;
  return 'Sonstiges';
}

// Experiment-ID-Muster: z.B. ETHG-002, OTAM-003-09, TEST-012-04-01
const EXP_ID_PATTERN = /^[A-Z]{2,5}-\d{2,3}(-\d{2,3})*$/;

// Alle bekannten Experiment-IDs als Set für schnelle Prüfung
const alleExpIds = new Set(expRaw.map(r => r[0]));

function klassifiziereKomponente(name) {
  if (!name) return 'text';
  const trimmed = String(name).trim();
  // Exakter Treffer in Experiment-Tabelle → Experiment-Referenz
  if (alleExpIds.has(trimmed)) return 'experiment';
  // Muster passt auf Experiment-ID-Format → Experiment-Referenz (auch wenn Eltern-ID fehlt)
  if (EXP_ID_PATTERN.test(trimmed)) return 'experiment';
  return 'chemikalie';
}

const ws2 = wb.Sheets['Komponenten'];
const kompRaw = XLSX.utils.sheet_to_json(ws2, {header:1}).slice(1).filter(r => r[0]);

const komponenten = kompRaw.map(r => {
  const name = r[1] ? String(r[1]).trim() : '';
  const typ = klassifiziereKomponente(name);
  const rolleOrig = r[6] ? String(r[6]).trim() : '';
  const rolleNorm = normRolle(rolleOrig, r[0], name);
  const rolleKomm = (rolleNorm === 'Sonstiges' && rolleOrig && rolleOrig !== 'Sonstiges') ? '[Orig. Rolle: ' + rolleOrig + ']' : '';
  return {
    Experiment_ID:      r[0],
    // Quelle-Typ: 'chemikalie' | 'experiment' | 'text'
    Quelle_Typ:         typ,
    // Genau eine der drei Quell-Spalten ist befüllt:
    Chemikalie_Name:    typ === 'chemikalie' ? name : '',
    Experiment_Ref:     typ === 'experiment' ? name : '',
    Komponente_Name:    typ === 'text'       ? name : '',
    Hersteller:         r[2] || '',
    Beschreibung:       r[3] || '',
    Menge:              r[4] != null ? r[4] : '',
    Einheit:            r[5] || '',
    Rolle:              rolleNorm,
    Kommentar:          rolleKomm,
  };
});

fs.writeFileSync('C:/Users/monak/projects/ligarolabui/Komponenten.csv', '\uFEFF' + toCSV(komponenten), 'utf8');
console.log('\nKomponenten.csv:', komponenten.length, 'Einträge');

// Quelle-Typ Verteilung
const typStat = {};
komponenten.forEach(k => { typStat[k.Quelle_Typ] = (typStat[k.Quelle_Typ] || 0) + 1; });
console.log('Quelle-Typ Verteilung:', JSON.stringify(typStat));

// Experiment-Referenzen auflisten
const expRefs = komponenten.filter(k => k.Quelle_Typ === 'experiment');
console.log('\nExperiment-Referenzen (' + expRefs.length + '):');
expRefs.forEach(k => console.log('  In', k.Experiment_ID, '→ Vorprodukt:', k.Experiment_Ref, '| Rolle:', k.Rolle));

// Rollen-Zusammenfassung
const rollenStat = {};
komponenten.forEach(k => { rollenStat[k.Rolle || '(leer)'] = (rollenStat[k.Rolle || '(leer)'] || 0) + 1; });
console.log('\nRollen:', JSON.stringify(rollenStat));

const sonstige = komponenten.filter(k => k.Rolle === 'Sonstiges');
if (sonstige.length > 0) {
  console.log('\nVerbleibende Sonstiges-Einträge:');
  sonstige.forEach(k => console.log('  Exp:', k.Experiment_ID, '| Name:', k.Chemikalie_Name || k.Experiment_Ref || k.Komponente_Name, '| Kommentar:', k.Kommentar));
}

// ── Materialprüfung ──────────────────────────────────────────────────────────
// Optimierungen:
// - Fläche entfernt (= Länge × Breite, redundant)
// - Datum: Excel-Seriennummer → ISO-Datum

const ws3 = wb.Sheets['Materialprüfung'];
const matRaw = XLSX.utils.sheet_to_json(ws3, {header:1}).slice(1).filter(r => r[0]);

// Jede Originalzeile enthält ggf. BEIDE Messwerte (trocken + nass) –
// das sind tatsächlich unterschiedliche Probekörper.
// → Aufspaltung: eine Zeile pro Probekörper.
// Lagerfolge_ID ersetzt Testbedingungen (Freitext).
// Default: LAFO-DRY-01 (trocken) / LAFO-WET-01 (nass) für Bestandsdaten.

const materialPruefung = [];
matRaw.forEach(r => {
  const basis = {
    Experiment_ID: r[0],
    // r[2] = Testbedingungen (meist leer) → in Kommentar falls vorhanden
    Laenge_mm:     r[3] != null ? r[3] : '',
    Breite_mm:     r[4] != null ? r[4] : '',
  };
  const testbedNote = r[2] ? '[Orig. Testbed.: ' + r[2] + '] ' : '';

  // Trockener Probekörper
  if (r[6] != null && r[6] !== '') {
    materialPruefung.push({
      ...basis,
      Lagerfolge_ID: 'LAFO-DRY-01',
      Kraft_N:       r[6],
      Holzbruch_pct: r[8] != null ? r[8] : '',
      Kommentar:     testbedNote + (r[12] || ''),
    });
  }

  // Nasser Probekörper (nur wenn Kraft vorhanden und nicht null/leer)
  if (r[9] != null && r[9] !== '') {
    materialPruefung.push({
      ...basis,
      Lagerfolge_ID: 'LAFO-WET-01',
      Kraft_N:       r[9],
      Holzbruch_pct: r[11] != null ? r[11] : '',
      Kommentar:     testbedNote + (r[12] || ''),
    });
  }
});

fs.writeFileSync('C:/Users/monak/projects/ligarolabui/Materialpruefung.csv', '\uFEFF' + toCSV(materialPruefung), 'utf8');
console.log('\nMaterialpruefung.csv:', materialPruefung.length, 'Einträge (aufgespalten aus', matRaw.length, 'Originalzeilen)');
const nurTrocken = materialPruefung.filter(m => m.Lagerfolge_ID === 'LAFO-DRY-01').length;
const nurNass    = materialPruefung.filter(m => m.Lagerfolge_ID === 'LAFO-WET-01').length;
console.log('  davon trocken (LAFO-DRY-01):', nurTrocken);
console.log('  davon nass    (LAFO-WET-01):', nurNass);

// ── Übersicht aller erzeugten Dateien ────────────────────────────────────────
console.log('\n=== Erzeugte CSV-Dateien ===');
['Personen', 'Projektkuerzel', 'Experimente', 'Komponenten', 'Materialpruefung'].forEach(name => {
  const path = 'C:/Users/monak/projects/ligarolabui/' + name + '.csv';
  const size = fs.statSync(path).size;
  console.log(' ', name + '.csv', '-', Math.round(size/1024) + ' KB');
});
