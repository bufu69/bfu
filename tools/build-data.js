// Încorporează datele unui fișier Excel în Dashboard_Analiza_5C.html (datele afișate la prima deschidere).
// Utilizare: npm i xlsx@0.18.5 && node tools/build-data.js "Analiza 5C.xlsx" [Dashboard_Analiza_5C.html]
const fs = require('fs'), path = require('path'), XLSX = require('xlsx');
const [xlsxPath, htmlArg] = process.argv.slice(2);
if (!xlsxPath) { console.error('Utilizare: node tools/build-data.js <fisier.xlsx> [dashboard.html]'); process.exit(1); }
const htmlPath = htmlArg || path.join(__dirname, '..', 'Dashboard_Analiza_5C.html');
let html = fs.readFileSync(htmlPath, 'utf8');
const src = html.slice(html.indexOf('/*PARSER-START*/'), html.indexOf('/*PARSER-END*/'));
const parse5C = new Function(src + '\nreturn parse5C;')();
const model = parse5C(XLSX.readFile(xlsxPath, { cellDates: false }), XLSX, path.basename(xlsxPath));
const json = JSON.stringify(model).replace(/</g, '\\u003c');
html = html.replace(/\/\*DEFAULT-DATA\*\/[\s\S]*?\/\*END-DEFAULT-DATA\*\//, () => `/*DEFAULT-DATA*/${json}/*END-DEFAULT-DATA*/`);
fs.writeFileSync(htmlPath, html);
console.log(`OK: ${model.weeks.length} săptămâni, ${model.records.length} rânduri, ${model.directory.length} persoane în director -> ${htmlPath}`);
