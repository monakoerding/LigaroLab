const XLSX = require('xlsx');
const wb = XLSX.readFile('C:/Users/monak/projects/ligarolabui/ExperimenteListe.xlsx');
const ws = wb.Sheets['Experimente'];
const exp = XLSX.utils.sheet_to_json(ws, {header:1}).slice(1).filter(r => r[0]);
const ids = exp.map(r => r[0]);

const d1 = ids.filter(id => /^[A-Z]+-\d+$/.test(id));
const d2 = ids.filter(id => /^[A-Z]+-\d+-\d+$/.test(id));
const d3 = ids.filter(id => /^[A-Z]+-\d+-\d+-\d+$/.test(id));
const dX = ids.filter(id => !/^[A-Z]+-\d+(-\d+)*$/.test(id));

console.log('Format AAAA-000:       ', d1.length);
console.log('Format AAAA-000-00:    ', d2.length);
console.log('Format AAAA-000-00-00: ', d3.length);
console.log('Andere Formate:        ', dX.length);
if (dX.length) console.log('Andere:', dX);

console.log('\nBeispiele Tiefe 3:', d3.slice(0, 5));

// Eltern-Beziehung prüfen: Hat jede Tiefe-2/3-ID eine existierende Eltern-ID?
const idSet = new Set(ids);
let missingParents = 0;
d2.forEach(id => {
  const parent = id.replace(/-\d+$/, '');
  if (!idSet.has(parent)) {
    missingParents++;
    console.log('KEIN PARENT:', id, '->', parent);
  }
});
d3.forEach(id => {
  const parent = id.replace(/-\d+$/, '');
  const grandparent = parent.replace(/-\d+$/, '');
  if (!idSet.has(parent)) console.log('KEIN PARENT:', id, '->', parent);
  if (!idSet.has(grandparent)) console.log('KEIN GRANDPARENT:', id, '->', grandparent);
});

console.log('\nVerzweigte Familien (Eltern mit mehreren Kindern):');
const families = {};
d2.forEach(id => {
  const parent = id.replace(/-\d+$/, '');
  if (!families[parent]) families[parent] = [];
  families[parent].push(id);
});
Object.entries(families)
  .filter(([, children]) => children.length > 1)
  .slice(0, 8)
  .forEach(([parent, children]) => console.log(' ', parent, '->', children.join(', ')));
