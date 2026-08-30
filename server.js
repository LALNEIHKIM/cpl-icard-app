const express = require('express');
const cors = require('cors');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

// Advanced Excel Date Formatter (Serial Number, Text, Date Object Sabhi Ke Liye)
function formatExcelDate(value) {
  if (!value && value !== 0) return '';
  
  // Agar pure number (Excel Serial Date) ho
  if (!isNaN(value) && typeof value !== 'boolean') {
    const num = Number(value);
    if (num > 1000) {
      // Excel epoch fix (1900 leap year bug accounted for)
      const date = new Date(Math.round((num - 25569) * 86400 * 1000));
      const d = String(date.getUTCDate()).padStart(2, '0');
      const m = String(date.getUTCMonth() + 1).padStart(2, '0');
      const y = date.getUTCFullYear();
      return `${d}/${m}/${y}`;
    }
  }

  // Agar already text/string date format me ho
  const str = String(value).trim();
  const parts = str.split(/[-/]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      // YYYY-MM-DD to DD/MM/YYYY
      return `${parts[2].padStart(2, '0')}/${parts[1].padStart(2, '0')}/${parts[0]}`;
    }
    return `${parts[0].padStart(2, '0')}/${parts[1].padStart(2, '0')}/${parts[2]}`;
  }

  return str;
}

app.get('/api/employees', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'CPL I CARD MAKER.xlsx');
    const workbook = xlsx.readFile(filePath);

    // MASTERLIST sheet select karein
    const sheetName = workbook.SheetNames.find(s => s.toUpperCase().includes('MASTER')) || workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Array of Arrays read karein (raw values false taaki dates formatted string me milein)
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const employees = [];

    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.length === 0) continue;

      const code = String(r[1] || '').trim();     // Col B (Code)
      const trade = String(r[2] || '').trim();    // Col C (Trade)
      const name = String(r[3] || '').trim();     // Col D (Name)
      const dobRaw = r[4];                        // Col E (DOB)
      const father = String(r[5] || '').trim();   // Col F (Father's Name)
      const doeRaw = r[22];                       // Col W (Date of Enrolment)

      // Agar valid name record hai
      if (name && !name.toUpperCase().includes('NAME') && isNaN(name)) {
        employees.push({
          name: name.toUpperCase(),
          code: code,
          trade: trade.toUpperCase(),
          father: father.toUpperCase(),
          dob: formatExcelDate(dobRaw),
          doe: formatExcelDate(doeRaw),
          idmark: '',
          rmpl: '403'
        });
      }
    }

    res.json(employees);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server live at http://localhost:${PORT}`);
});