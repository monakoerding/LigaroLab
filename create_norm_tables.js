const fs = require('fs');

function toCSV(arr) {
  const keys = Object.keys(arr[0]);
  const esc = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
  return [keys.join(','), ...arr.map(row => keys.map(k => esc(row[k])).join(','))].join('\r\n');
}

// ── 1. Spanplatten_Produktionsparameter ──────────────────────────────────────
// Rezeptur-Stammdaten für Spanplatten, ohne Klebstoff-Referenz.
// Klebstoff wird erst bei der Prüfung (EN319_Pruefung) angegeben.

const spanplatten = [
  {
    Spanplatten_ID:         'SPAN-EGGR-001',
    Dichte_kg_m3:           650,
    Anteil_Klebstoff_pct:   12,
    Presstemperatur_C:      200,
    Pressdruck_MPa:         3.5,
    Typ_Spaene:             'Fichtenspäne (industriell, 0.2–0.4 mm)',
    Kommentar:              'Erste Referenz-Rezeptur; Schichtaufbau symmetrisch',
  },
];

fs.writeFileSync(
  'C:/Users/monak/projects/ligarolabui/Spanplatten_Produktionsparameter.csv',
  '\uFEFF' + toCSV(spanplatten), 'utf8'
);
console.log('Spanplatten_Produktionsparameter.csv:', spanplatten.length, 'Einträge');
console.log('Spalten:', Object.keys(spanplatten[0]).join(', '));

// ── 2. Lagerfolgen ───────────────────────────────────────────────────────────
// Kopfdaten jeder Konditionierungsprozedur.
// Gilt für EN 12765 (Holzverklebung) UND EN 319 (Spanplatten) – eine gemeinsame Tabelle.

const lagerfolgen = [
  // ── Bisherige Praxis (Bestandsdaten) ────────────────────────────────────────
  {
    Lagerfolge_ID:  'LAFO-DRY-01',
    Name:           'Trockenlagerung 24h (bisherige Praxis)',
    Norm:           'intern',
    Anwendung:      'Lagerung 24 h bei 21 °C / 65 % rH, anschließend trockene Zugprüfung',
    Anzahl_Schritte: 1,
  },
  {
    Lagerfolge_ID:  'LAFO-WET-01',
    Name:           'Nasstest nach 24h Wasser (bisherige Praxis)',
    Norm:           'intern',
    Anwendung:      'Lagerung 24 h bei 21 °C / 65 % rH, dann 24 h Wasserlagerung RT, nasse Zugprüfung',
    Anzahl_Schritte: 2,
  },
  // ── EN 12765:2016 Lagerungsfolgen (Tabelle 2) ─────────────────────────────
  // Jede LAFO entspricht einer nummerierten Lagerungsfolge aus der Norm.
  // C1 erfordert LF1; C2 erfordert LF1+LF2; C3 erfordert LF1+LF2+LF3; C4 erfordert LF1+LF2+LF4.
  {
    Lagerfolge_ID:  'LAFO-C1-1',
    Name:           'EN 12765 Lagerungsfolge 1 (C1)',
    Norm:           'EN 12765:2016',
    Anwendung:      'Holzverklebung Klasse C1 – trockener Innenbereich (Holzfeuchte ≤ 15 %). 7 Tage Normalklima, Prüfung trocken.',
    Anzahl_Schritte: 1,
  },
  {
    Lagerfolge_ID:  'LAFO-C2-1',
    Name:           'EN 12765 Lagerungsfolge 2 (C2)',
    Norm:           'EN 12765:2016',
    Anwendung:      'Holzverklebung Klasse C2 – Innenbereich mit gelegentlicher Feuchteexposition (Holzfeuchte bis 18 %). 7 Tage Normalklima + 1 Tag Wasser (20 °C), Prüfung nass.',
    Anzahl_Schritte: 2,
  },
  {
    Lagerfolge_ID:  'LAFO-C3-1',
    Name:           'EN 12765 Lagerungsfolge 3 (C3)',
    Norm:           'EN 12765:2016',
    Anwendung:      'Holzverklebung Klasse C3 – Innenbereich mit häufiger Feuchteexposition oder Außenbereich geschützt. 7 Tage Normalklima + 3 h Wasser (67 °C) + 2 h Wasser (20 °C), Prüfung nass.',
    Anzahl_Schritte: 3,
  },
  {
    Lagerfolge_ID:  'LAFO-C4-1',
    Name:           'EN 12765 Lagerungsfolge 4 (C4)',
    Norm:           'EN 12765:2016',
    Anwendung:      'Holzverklebung Klasse C4 – Außenbereich wetterexponiert mit Oberflächenschutz. 7 Tage Normalklima + 3 h kochendes Wasser + 2 h Wasser (20 °C), Prüfung nass.',
    Anzahl_Schritte: 3,
  },
];

fs.writeFileSync(
  'C:/Users/monak/projects/ligarolabui/Lagerfolgen.csv',
  '\uFEFF' + toCSV(lagerfolgen), 'utf8'
);
console.log('\nLagerfolgen.csv:', lagerfolgen.length, 'Einträge');

// ── 3. Lagerfolgen_Schritte ──────────────────────────────────────────────────
// Eine Zeile pro Schritt. Medium: Luft, Wasser, Wasserbad, etc.
// RH_pct = relative Luftfeuchtigkeit; leer bei Wasserlagerung.

const schritte = [
  // LAFO-DRY-01 – bisherige Trockentestpraxis
  { Lagerfolge_ID: 'LAFO-DRY-01', Schritt_Nr: 1, Behandlung: 'Klimatisierung', Medium: 'Luft',    Temperatur_C: 21, RH_pct: 65, Dauer_h: 24,  Kommentar: '24 h bei 21 °C / 65 % rH; danach trockene Zugprüfung' },
  // LAFO-WET-01 – bisherige Nasstestpraxis
  { Lagerfolge_ID: 'LAFO-WET-01', Schritt_Nr: 1, Behandlung: 'Klimatisierung', Medium: 'Luft',    Temperatur_C: 21, RH_pct: 65, Dauer_h: 24,  Kommentar: '24 h bei 21 °C / 65 % rH' },
  { Lagerfolge_ID: 'LAFO-WET-01', Schritt_Nr: 2, Behandlung: 'Wasserlagerung', Medium: 'Wasser',  Temperatur_C: 21, RH_pct: '',  Dauer_h: 24,  Kommentar: '24 h in Wasser bei RT; danach nasse Zugprüfung' },
  // EN 12765 LF1 (C1) – 7 Tage Normalklima, trocken
  { Lagerfolge_ID: 'LAFO-C1-1', Schritt_Nr: 1, Behandlung: 'Klimatisierung', Medium: 'Luft',      Temperatur_C: 20, RH_pct: 65, Dauer_h: 168, Kommentar: '7 Tage (168 h) bei 20 °C / 65 % rH (Normalklima); danach trockene Zugprüfung' },
  // EN 12765 LF2 (C2) – 7 Tage Normalklima + 1 Tag Wasser RT, nass
  { Lagerfolge_ID: 'LAFO-C2-1', Schritt_Nr: 1, Behandlung: 'Klimatisierung', Medium: 'Luft',      Temperatur_C: 20, RH_pct: 65, Dauer_h: 168, Kommentar: '7 Tage (168 h) bei 20 °C / 65 % rH (Normalklima)' },
  { Lagerfolge_ID: 'LAFO-C2-1', Schritt_Nr: 2, Behandlung: 'Wasserlagerung', Medium: 'Wasser',    Temperatur_C: 20, RH_pct: '',  Dauer_h: 24,  Kommentar: '1 Tag (24 h) in Wasser bei (20 ± 5) °C; danach nasse Zugprüfung' },
  // EN 12765 LF3 (C3) – 7 Tage Normalklima + 3 h 67 °C + 2 h 20 °C, nass
  { Lagerfolge_ID: 'LAFO-C3-1', Schritt_Nr: 1, Behandlung: 'Klimatisierung',    Medium: 'Luft',   Temperatur_C: 20, RH_pct: 65, Dauer_h: 168, Kommentar: '7 Tage (168 h) bei 20 °C / 65 % rH (Normalklima)' },
  { Lagerfolge_ID: 'LAFO-C3-1', Schritt_Nr: 2, Behandlung: 'Heißwasserlagerung', Medium: 'Wasser', Temperatur_C: 67, RH_pct: '',  Dauer_h: 3,   Kommentar: '3 h in Wasser bei (67 ± 2) °C' },
  { Lagerfolge_ID: 'LAFO-C3-1', Schritt_Nr: 3, Behandlung: 'Abkühlung',         Medium: 'Wasser', Temperatur_C: 20, RH_pct: '',  Dauer_h: 2,   Kommentar: '2 h in Wasser bei (20 ± 5) °C; danach nasse Zugprüfung' },
  // EN 12765 LF4 (C4) – 7 Tage Normalklima + 3 h kochend + 2 h 20 °C, nass
  { Lagerfolge_ID: 'LAFO-C4-1', Schritt_Nr: 1, Behandlung: 'Klimatisierung',    Medium: 'Luft',   Temperatur_C: 20, RH_pct: 65, Dauer_h: 168, Kommentar: '7 Tage (168 h) bei 20 °C / 65 % rH (Normalklima)' },
  { Lagerfolge_ID: 'LAFO-C4-1', Schritt_Nr: 2, Behandlung: 'Kochwasserlagerung', Medium: 'Wasser', Temperatur_C: 100, RH_pct: '', Dauer_h: 3,   Kommentar: '3 h in kochendem Wasser (100 °C)' },
  { Lagerfolge_ID: 'LAFO-C4-1', Schritt_Nr: 3, Behandlung: 'Abkühlung',         Medium: 'Wasser', Temperatur_C: 20, RH_pct: '',  Dauer_h: 2,   Kommentar: '2 h in Wasser bei (20 ± 5) °C; danach nasse Zugprüfung' },
];

fs.writeFileSync(
  'C:/Users/monak/projects/ligarolabui/Lagerfolgen_Schritte.csv',
  '\uFEFF' + toCSV(schritte), 'utf8'
);
console.log('\nLagerfolgen_Schritte.csv:', schritte.length, 'Einträge');
console.log('Spalten:', Object.keys(schritte[0]).join(', '));

// ── 4. EN319_Pruefung ────────────────────────────────────────────────────────
// Eine Zeile pro Probekörper.
// Querzugfestigkeit_MPa = Kraft_N / (Laenge_mm × Breite_mm) – berechnet in Power Apps.
// Bruchbild-Werte nach EN 319: kohäsiv (Platte), adhäsiv (Klebefuge), Mischbruch.

const en319 = [
  {
    Experiment_ID:      'TEST-041-03',    // Klebstoff-Experiment
    Spanplatten_ID:     'SPAN-EGGR-001',  // Spanplatten-Rezeptur
    Lagerfolge_ID:      'LAFO-C1-1',
    Laenge_mm:          50,
    Breite_mm:          50,
    Kraft_N:            1250,
    // Querzugfestigkeit = 1250 / (50×50) = 0.50 MPa – berechnet in Power Apps
    Bruchbild:          'kohäsiv (Spanplatte)',
    Kommentar:          'Beispieleintrag',
  },
];

fs.writeFileSync(
  'C:/Users/monak/projects/ligarolabui/EN319_Pruefung.csv',
  '\uFEFF' + toCSV(en319), 'utf8'
);
console.log('\nEN319_Pruefung.csv:', en319.length, 'Einträge');
console.log('Spalten:', Object.keys(en319[0]).join(', '));

// ── Hinweis: Materialprüfung anpassen ────────────────────────────────────────
console.log('\n── Empfehlung ──────────────────────────────────────────────────────────');
console.log('Materialprüfung.csv: Spalte "Testbedingungen" (Freitext) ersetzen durch');
console.log('"Lagerfolge_ID" (Lookup → Lagerfolgen), damit EN 12765 und EN 319');
console.log('dieselbe Lagerfolgen-Tabelle nutzen.');
