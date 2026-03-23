const XLSX = require('xlsx');
const wb = XLSX.readFile('C:/Users/monak/projects/ligarolabui/ExperimenteListe.xlsx');
const ws2 = wb.Sheets['Komponenten'];
const kompRaw = XLSX.utils.sheet_to_json(ws2, {header:1}).slice(1).filter(r => r[0]);
const EXP_ID_PATTERN = /^[A-Z]{2,5}-\d{2,3}(-\d{2,3})*$/;
const ws1 = wb.Sheets['Experimente'];
const expRaw = XLSX.utils.sheet_to_json(ws1, {header:1}).slice(1).filter(r => r[0]);
const alleExpIds = new Set(expRaw.map(r => r[0]));

kompRaw.forEach(r => {
  const name = r[1] ? String(r[1]).trim() : '';
  const isExp = alleExpIds.has(name) || EXP_ID_PATTERN.test(name);
  const isChem = name.length > 0 && !isExp;
  // "text" = non-empty, not exp, but also not matching any chemikalie
  // For now just find entries that look unusual
  if (name && !isExp) {
    // check if it looks like a non-chemical (very short, numeric, etc.)
    if (/^\d/.test(name) || name.length < 3) {
      console.log('UNUSUAL:', JSON.stringify(r));
    }
  }
});

// Find the 2 "text" entries by checking what's not matching either pattern
// "text" type = doesn't match exp pattern AND is empty or very unusual
kompRaw.forEach(r => {
  const name = r[1] ? String(r[1]).trim() : '';
  if (!name) console.log('LEER:', JSON.stringify(r));
});
