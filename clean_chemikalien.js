const XLSX = require('xlsx');
const fs = require('fs');

const wb = XLSX.readFile('C:/Users/monak/projects/ligarolabui/Chemikalienliste.xlsx');
const ws = wb.Sheets['Sheet1'];
const raw = XLSX.utils.sheet_to_json(ws, {header:1}).slice(1).filter(r => r[0]);

// ── Hilfsfunktionen ──────────────────────────────────────────────────────────

function normHersteller(h) {
  if (!h || h === '-') return '';
  return h.trim()
    .replace('Sigma-Alrich', 'Sigma-Aldrich')
    .replace('avebe', 'Avebe');
}

function parseMenge(m) {
  if (!m) return { zahl: '', einheit: '' };
  const s = String(m).trim().replace(',', '.');
  const match = s.match(/^([0-9.]+)\s*([a-zA-Z]+)/);
  if (!match) return { zahl: s, einheit: '' };
  const e = match[2].trim().toLowerCase();
  const einheitNorm = e === 'l' ? 'L' : e === 'ml' ? 'mL' : e === 'g' ? 'g' : e === 'kg' ? 'kg' : match[2].trim();
  return { zahl: parseFloat(match[1]), einheit: einheitNorm };
}

function mapVerfuegbarkeit(statusList) {
  const s = statusList.map(x => (x || '').toLowerCase());
  if (s.every(x => x === 'entsorgen')) return 'Entsorgen';
  if (s.every(x => x === 'sample' || x === 'entsorgen')) return 'Sample';
  if (s.some(x => x === 'noch nicht geliefert') && s.every(x => ['noch nicht geliefert', 'entsorgen'].includes(x))) return 'Bestellt (noch nicht geliefert)';
  return 'Vorrätig';
}

function mapHerkunft(status) {
  const s = (status || '').toLowerCase().trim();
  if (s === 'sample') return 'Sample';
  if (s === 'labor gegenüber') return 'Übernommen, Obstbau';
  if (s === 'übernommen') return 'Übernommen, Gemüsebau';
  if (s === 'entsorgen') return 'Entsorgen';
  if (s === 'noch nicht geliefert') return 'Gekauft';
  return 'Gekauft';
}

// ── Gruppen-Definition ───────────────────────────────────────────────────────
// Key = Name des Allgemein-Eintrags, Value = Array der zugehörigen spezifischen Namen
const GRUPPEN = {
  'Schwefelsäure (allgemein)':        ['Schwefelsäure', 'Schwefelsäure 45% (H2SO4)'],
  'Natriumhydroxid/NaOH (allgemein)': ['NaOH 100%', 'NaOH 50%', 'Natriumhydroxid'],
  'Chitosan (allgemein)':             ['Chitosan', 'Chitosan, from crab shells'],
  'Salzsäure/HCl (allgemein)':        ['Salzsäure', 'Salzsäure 1N (HCl)', 'Salzsäure 37 %'],
  'Tannin (allgemein)':               ['Tannin / Tannic acid', 'Tannin braun, wasser- und alkohollöslich', 'Tannin braun, wasserlöslich'],
  // Gluten, Lignin, Kartoffelprotein, Casein, Biertreber: keine Allgemein-Einträge
  // (chemisch nicht exakt definiert – auf Wunsch des Teams entfernt)
  'Calciumhydroxid (allgemein)':      ['Calciumhydroxid, gesumpft', 'Calciumhydroxid, pulver'],
  'Ortho-Phosphorsäure (allgemein)':  ['Ortho-Phosphorsäure'],
};

// Umgekehrte Zuordnung: Name → Gruppenname
const nameZuGruppe = {};
Object.entries(GRUPPEN).forEach(([gruppe, namen]) => {
  namen.forEach(n => { nameZuGruppe[n] = gruppe; });
});

// ── Rohdaten bereinigen ──────────────────────────────────────────────────────

const rows = raw.map(r => ({
  name:       (r[0] || '').trim(),
  menge:      r[1],
  hersteller: normHersteller(r[2]),
  ort:        (r[3] || '').trim(),
  status:     (r[4] || '').trim(),
  kommentar:  (r[5] || '').trim(),
}));

// Ort korrigieren: "Labor gegenüber" im Status → Ort setzen
rows.forEach(r => {
  if (r.status === 'Labor gegenüber') r.ort = 'Labor gegenüber';
});

// ── Chemikalien (Stammdaten, eindeutig nach Name) ────────────────────────────

const byName = {};
rows.forEach(r => {
  if (!byName[r.name]) byName[r.name] = [];
  byName[r.name].push(r);
});

const chemikalien = Object.entries(byName).map(([name, group]) => {
  const kommentare = [...new Set(group.map(g => g.kommentar).filter(Boolean))].join(' | ');
  const verfuegbarkeit = mapVerfuegbarkeit(group.map(g => g.status));
  const gruppe = nameZuGruppe[name] || '';
  return {
    Name:          name,
    Gruppe:        gruppe,
    Ist_Allgemein: 'Nein',
    Verfuegbarkeit: verfuegbarkeit,
    Kommentar:     kommentare,
  };
});

// Allgemein-Einträge hinzufügen (ein Eintrag pro Gruppe)
Object.entries(GRUPPEN).forEach(([allgemeinName, mitglieder]) => {
  // Verfügbarkeit: "Vorrätig" wenn mindestens ein Mitglied vorrätig ist
  const mitgliederRows = rows.filter(r => mitglieder.includes(r.name));
  const verfuegbarkeit = mitgliederRows.length > 0
    ? mapVerfuegbarkeit(mitgliederRows.map(r => r.status))
    : 'Vorrätig';
  chemikalien.push({
    Name:           allgemeinName,
    Gruppe:         allgemeinName,
    Ist_Allgemein:  'Ja',
    Verfuegbarkeit: verfuegbarkeit,
    Kommentar:      'Allgemeiner Eintrag – für unspezifische Verwendung in Experimenten',
  });
});

// Sortieren: Allgemein-Einträge zuerst, dann alphabetisch
chemikalien.sort((a, b) => {
  if (a.Ist_Allgemein !== b.Ist_Allgemein) return a.Ist_Allgemein === 'Ja' ? -1 : 1;
  return a.Name.localeCompare(b.Name, 'de');
});

// ── IUPAC / Formel Lookup ─────────────────────────────────────────────────────
// Nur chemisch eindeutig definierte Reinstoffe; komplexe Gemische, Polymere,
// Proteine, Lignine, Tannine, pflanzliche Extrakte bleiben leer.

const IUPAC_MAP = {
  // Alkohole
  '2-Propanol':                          { iupac: 'Propan-2-ol',                                   formel: 'CH3CH(OH)CH3' },
  'Isopropanol (iPrOH)':                 { iupac: 'Propan-2-ol',                                   formel: 'CH3CH(OH)CH3' },
  'Ethanol':                             { iupac: 'Ethanol',                                        formel: 'C2H5OH' },
  'Methanol':                            { iupac: 'Methanol',                                       formel: 'CH3OH' },
  'Glycerin':                            { iupac: 'Propan-1,2,3-triol',                             formel: 'HOCH2CH(OH)CH2OH' },
  'Resorcinol':                          { iupac: 'Benzol-1,3-diol',                                formel: 'C6H6O2' },
  // Carbonsäuren
  'Essigsäure':                          { iupac: 'Ethansäure',                                     formel: 'CH3COOH' },
  'Essigsäure ':                         { iupac: 'Ethansäure',                                     formel: 'CH3COOH' },
  'Citronensäure':                       { iupac: '2-Hydroxypropan-1,2,3-tricarbonsäure',           formel: 'C6H8O7' },
  'Salicylsäure':                        { iupac: '2-Hydroxybenzoesäure',                           formel: 'HOC6H4COOH' },
  'Ortho-Phosphorsäure':                 { iupac: 'Phosphorsäure',                                  formel: 'H3PO4' },
  // Ester / Ketone / Aldehyde
  'Essigsäureethylester':                { iupac: 'Ethylethanoat',                                  formel: 'CH3COOC2H5' },
  'Aceton':                              { iupac: 'Propan-2-on',                                    formel: 'CH3COCH3' },
  'Acetaldehyd':                         { iupac: 'Ethanal',                                        formel: 'CH3CHO' },
  'Kampher':                             { iupac: '(1R,4R)-1,7,7-Trimethylbicyclo[2.2.1]heptan-2-on', formel: 'C10H16O' },
  'Genipin':                             { iupac: 'Methyl (1R,2R,6S)-2-hydroxy-9-(hydroxymethyl)-3-oxabicyclo[4.3.0]nona-4,8-dien-5-carboxylat', formel: 'C11H14O5' },
  // Zucker
  'D(-)-Fructose':                       { iupac: '(3S,4R,5R)-1,3,4,5,6-Pentahydroxyhexan-2-on',  formel: 'C6H12O6' },
  'D(+)-Glucose Monohydrat':             { iupac: '(2R,3S,4R,5R)-2,3,4,5,6-Pentahydroxyhexanal-Monohydrat', formel: 'C6H12O6·H2O' },
  'D(+)-Saccharose':                     { iupac: 'alpha-D-Glucopyranosyl-(1→2)-beta-D-fructofuranosid', formel: 'C12H22O11' },
  // Stickstoff-organisch
  'TRIS':                                { iupac: 'Tris(hydroxymethyl)aminomethan',                 formel: 'C4H11NO3' },
  'Urea (Harnstoff)':                    { iupac: 'Harnstoff',                                      formel: '(NH2)2CO' },
  'Kaliumsorbat, Granulat':              { iupac: 'Kalium-(2E,4E)-hexa-2,4-dienoat',               formel: 'C6H7KO2' },
  'EDTA Dinatriumsalz-Lösung':          { iupac: 'Dinatriumethylendiamintetraacetat',              formel: 'C10H14N2Na2O8' },
  'SDS / Natriumlaurysulfat':            { iupac: 'Natriumdodecylsulfat',                          formel: 'C12H25NaO4S' },
  'Natriumsalicylat':                    { iupac: 'Natrium-2-hydroxybenzoat',                       formel: 'NaC7H5O3' },
  // Anorganische Säuren
  'Salpetersäure 65 %':                  { iupac: 'Salpetersäure',                                  formel: 'HNO3' },
  'Salzsäure':                           { iupac: 'Chlorwasserstoffsäure',                          formel: 'HCl' },
  'Salzsäure 1N (HCl)':                  { iupac: 'Chlorwasserstoffsäure',                          formel: 'HCl' },
  'Salzsäure 37 %':                      { iupac: 'Chlorwasserstoffsäure',                          formel: 'HCl' },
  'Schwefelsäure':                       { iupac: 'Schwefelsäure',                                  formel: 'H2SO4' },
  'Schwefelsäure 45% (H2SO4)':          { iupac: 'Schwefelsäure',                                  formel: 'H2SO4' },
  'Phosphorpentoxid':                    { iupac: 'Phosphor(V)-oxid',                               formel: 'P4O10' },
  'Wasserstoffperoxid 35 %':             { iupac: 'Wasserstoffperoxid',                             formel: 'H2O2' },
  // Natrium-Verbindungen
  'NaOH 100%':                           { iupac: 'Natriumhydroxid',                                formel: 'NaOH' },
  'NaOH 50%':                            { iupac: 'Natriumhydroxid',                                formel: 'NaOH' },
  'Natriumhydroxid':                     { iupac: 'Natriumhydroxid',                                formel: 'NaOH' },
  'Natriumchlorid':                      { iupac: 'Natriumchlorid',                                 formel: 'NaCl' },
  'Natriumcitrat (Na3-Citrate)':         { iupac: 'Trinatriumcitrat',                               formel: 'Na3C6H5O7' },
  'Natriumdisulfit':                     { iupac: 'Dinatriumdisulfit',                              formel: 'Na2S2O5' },
  'Natriumhexametaphosphat':             { iupac: 'Hexanatriumhexametaphosphat',                    formel: '(NaPO3)6' },
  'Natriumnitrat':                       { iupac: 'Natriumnitrat',                                  formel: 'NaNO3' },
  'Natriumnitrit':                       { iupac: 'Natriumnitrit',                                  formel: 'NaNO2' },
  'Natriumperjodat':                     { iupac: 'Natriumperiodat',                                formel: 'NaIO4' },
  'Natriumthiosulfat Pentahydrat (Na2S2O3*5H2O)': { iupac: 'Natriumthiosulfat-Pentahydrat',        formel: 'Na2S2O3·5H2O' },
  'Natriumtrimetaphosphat STMP':         { iupac: 'Natriumtrimetaphosphat',                         formel: '(NaPO3)3' },
  'Natriumtripolyphosphat STPP E451':    { iupac: 'Pentanatriumtriphosphat',                        formel: 'Na5P3O10' },
  'Natriummolybdat technisch':           { iupac: 'Dinatriumtetraoxomolybdat',                      formel: 'Na2MoO4' },
  // Kalium-Verbindungen
  'Di-Kaliumhydrogenphosphat':           { iupac: 'Dikaliumhydrogenphosphat',                       formel: 'K2HPO4' },
  'Kaliumchlorid (KCl)':                 { iupac: 'Kaliumchlorid',                                  formel: 'KCl' },
  'Kaliumdihydrogenphosphat':            { iupac: 'Kaliumdihydrogenphosphat',                       formel: 'KH2PO4' },
  'Kaliumnitrat':                        { iupac: 'Kaliumnitrat',                                   formel: 'KNO3' },
  'Kaliumsulfat (K2SO4)':                { iupac: 'Kaliumsulfat',                                   formel: 'K2SO4' },
  // Calcium-Verbindungen
  'Calciumchlorid-Dihydrat (CaCl2*2H2O)': { iupac: 'Calciumchlorid-Dihydrat',                      formel: 'CaCl2·2H2O' },
  'Calciumhydroxid, gesumpft':           { iupac: 'Calciumdihydroxid',                              formel: 'Ca(OH)2' },
  'Calciumhydroxid, pulver':             { iupac: 'Calciumdihydroxid',                              formel: 'Ca(OH)2' },
  'Calciumnitrat Tetrahydrate':          { iupac: 'Calciumnitrat-Tetrahydrat',                      formel: 'Ca(NO3)2·4H2O' },
  'Calciumnitrit Lösung 30 wt% in Wasser': { iupac: 'Calciumnitrit',                               formel: 'Ca(NO2)2' },
  // Magnesium-Verbindungen
  'Magnesiumhydroxid':                   { iupac: 'Magnesiumdihydroxid',                            formel: 'Mg(OH)2' },
  'Magnesiumoxid':                       { iupac: 'Magnesiumoxid',                                  formel: 'MgO' },
  'Magnesiumsulfat Heptahydrat':         { iupac: 'Magnesiumsulfat-Heptahydrat',                    formel: 'MgSO4·7H2O' },
  // Aluminium / weitere Metalle
  'Aluminiumchlorid-Hexahydrat':         { iupac: 'Aluminiumtrichlorid-Hexahydrat',                 formel: 'AlCl3·6H2O' },
  'Aluminiumnitrat Nonahydrat':          { iupac: 'Aluminiumnitrat-Nonahydrat',                     formel: 'Al(NO3)3·9H2O' },
  'Bariumchlorid Dihydrat':              { iupac: 'Bariumdichlorid-Dihydrat',                       formel: 'BaCl2·2H2O' },
  'Eisen(III)-chlorid Hexahydrat':       { iupac: 'Eisen(III)-chlorid-Hexahydrat',                  formel: 'FeCl3·6H2O' },
  'Kupfersulfat (CuSO4)':                { iupac: 'Kupfer(II)-sulfat',                              formel: 'CuSO4' },
  'Silbernitrat (AgNO3)':                { iupac: 'Silbernitrat',                                   formel: 'AgNO3' },
  'Ammoniumheptamolybdat':               { iupac: 'Hexaammoniumheptamolybdat-Tetrahydrat',          formel: '(NH4)6Mo7O24·4H2O' },
  'Ammoniummonovandat':                  { iupac: 'Ammoniummetavanadat',                            formel: 'NH4VO3' },
  'Ammoniumperoxodisulfat':              { iupac: 'Diammoniumperoxodisulfat',                       formel: '(NH4)2S2O8' },
};

// ── Chemikalienbestand (Gebinde) ─────────────────────────────────────────────

const bestand = rows.map(r => {
  const { zahl, einheit } = parseMenge(r.menge);
  const lookup = IUPAC_MAP[r.name] || {};
  return {
    Chemikalie_Name: r.name,
    IUPAC:           lookup.iupac || '',
    Formel:          lookup.formel || '',
    Hersteller:      r.hersteller,
    Menge:           zahl,
    Einheit:         einheit,
    Ort:             r.ort,
    Herkunft:        mapHerkunft(r.status),
    Kommentar:       r.kommentar,
  };
});

// ── CSV-Ausgabe ───────────────────────────────────────────────────────────────

function toCSV(arr) {
  const keys = Object.keys(arr[0]);
  const esc = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
  return [keys.join(','), ...arr.map(row => keys.map(k => esc(row[k])).join(','))].join('\r\n');
}

fs.writeFileSync('C:/Users/monak/projects/ligarolabui/Chemikalien_Stammdaten.csv', '\uFEFF' + toCSV(chemikalien), 'utf8');
fs.writeFileSync('C:/Users/monak/projects/ligarolabui/Chemikalienbestand.csv',    '\uFEFF' + toCSV(bestand),     'utf8');

// ── Ausgabe ──────────────────────────────────────────────────────────────────

console.log('=== Chemikalien_Stammdaten.csv ===');
console.log('Allgemein-Einträge:         ', chemikalien.filter(c => c.Ist_Allgemein === 'Ja').length);
console.log('Spezifische Einträge:       ', chemikalien.filter(c => c.Ist_Allgemein === 'Nein').length);
console.log('Gesamt:                     ', chemikalien.length);

console.log('\nVerfügbarkeiten (spezifisch):');
const vmap = {};
chemikalien.filter(c => c.Ist_Allgemein === 'Nein').forEach(c => { vmap[c.Verfuegbarkeit] = (vmap[c.Verfuegbarkeit] || 0) + 1; });
Object.entries(vmap).forEach(([k, v]) => console.log(' ', k + ':', v));

console.log('\nAllgemein-Einträge mit Gruppe:');
chemikalien.filter(c => c.Ist_Allgemein === 'Ja').forEach(c => console.log(' -', c.Name, '(' + c.Verfuegbarkeit + ')'));

console.log('\n=== Chemikalienbestand.csv ===');
console.log('Gebinde gesamt:', bestand.length);
console.log('\nHerkunft-Verteilung:');
const hmap = {};
bestand.forEach(b => { hmap[b.Herkunft] = (hmap[b.Herkunft] || 0) + 1; });
Object.entries(hmap).forEach(([k, v]) => console.log(' ', k + ':', v));

console.log('\nBeispiel Schwefelsäure-Gruppe:');
chemikalien.filter(c => c.Gruppe === 'Schwefelsäure (allgemein)').forEach(c =>
  console.log(' ', c.Ist_Allgemein === 'Ja' ? '[ALLGEMEIN]' : '[SPEZIFISCH]', c.Name)
);

console.log('\nBeispiel NaOH-Gruppe:');
chemikalien.filter(c => c.Gruppe === 'Natriumhydroxid/NaOH (allgemein)').forEach(c =>
  console.log(' ', c.Ist_Allgemein === 'Ja' ? '[ALLGEMEIN]' : '[SPEZIFISCH]', c.Name)
);
