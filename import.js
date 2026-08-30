const sqlite3 = require('sqlite3').verbose();
const xlsx = require('xlsx');
const path = require('path');

const db = new sqlite3.Database(path.join(__dirname, 'icard.db'));
const workbook = xlsx.readFile(path.join(__dirname, 'CPL I CARD MAKER.xlsx'), { cellDates: false });

const sheetName = workbook.SheetNames.includes('MASTERLIST') ? 'MASTERLIST' : workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });

console.log(`Loading Sheet: ${sheetName} from CPL I CARD MAKER.xlsx`);

// Excel Date Serial Number (e.g. 30317) ko DD/MM/YYYY me convert karne ka function
function formatExcelDate(val) {
  if (!val) return '';
  
  // Agar Excel number format me hai (e.g. 30317, 46155)
  if (!isNaN(val) && typeof val === 'number' || (!isNaN(val) && String(val).trim().length === 5)) {
    const num = Number(val);
    if (num > 1000) { // Valid serial number check
      const date = new Date(Math.round((num - 25569) * 86400 * 1000));
      const day = String(date.getUTCDate()).padStart(2, '0');
      const month = String(date.getUTCMonth() + 1).padStart(2, '0');
      const year = date.getUTCFullYear();
      return `${day}/${month}/${year}`;
    }
  }

  // Agar already text/string hai (jaise "19-04-1987" ya "Transferred from...")
  let str = String(val).trim();
  str = str.replace(/-/g, '/'); // 19-04-1987 ko 19/04/1987 standard karega
  return str;
}

db.serialize(() => {
  db.run(`DROP TABLE IF EXISTS employees`);

  db.run(`CREATE TABLE employees (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    code TEXT,
    father TEXT,
    trade TEXT,
    dob TEXT,
    doe TEXT,
    idmark TEXT
  )`);

  const stmt = db.prepare(`INSERT INTO employees (name, code, father, trade, dob, doe, idmark) VALUES (?, ?, ?, ?, ?, ?, ?)`);

  let count = 0;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length < 4) continue;

    const code = String(r[1] || '').trim();
    const trade = String(r[2] || '').trim();
    const name = String(r[3] || '').trim();
    const dob = formatExcelDate(r[4]);                 // Column E (Date of Birth formatted)
    const father = String(r[5] || '').trim();
    const doe = formatExcelDate(r[22]);                // Column W (Date of Enrolment formatted)
    const idmark = '';                                 // Blank

    if (name && name.toUpperCase() !== 'NAME' && name.toUpperCase() !== 'NAME OF CPL' && isNaN(name)) {
      stmt.run(name, code, father, trade, dob, doe, idmark);
      count++;
    }
  }

  stmt.finalize(() => {
    console.log(`✅ Success! ${count} Employees Cleanly Imported with Proper Dates.`);
  });
});