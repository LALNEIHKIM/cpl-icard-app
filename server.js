const express = require('express');
const cors = require('cors');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

// Date Formatter (Handles Excel Serial & Text Dates like 19-04-1987 or 21-Nov-25)
function formatExcelDate(value) {
  if (!value) return '';
  if (typeof value === 'number') {
    const date = new Date(Math.round((value - 25569) * 86400 * 1000));
    const d = String(date.getUTCDate()).padStart(2, '0');
    const m = String(date.getUTCMonth() + 1).padStart(2, '0');
    const y = date.getUTCFullYear();
    return `${d}/${m}/${y}`;
  }
  return String(value).trim();
}

app.get('/api/employees', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'CPL I CARD MAKER.xlsx');
    const workbook = xlsx.readFile(filePath);

    // MASTERLIST sheet select karein
    const sheetName = workbook.SheetNames.find(s => s.toUpperCase().includes('MASTER')) || workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(sheet, { defval: '' });

    const formattedData = rawData.map(row => {
      // Keys matching helper
      const getVal = (...keys) => {
        for (let k of keys) {
          for (let prop in row) {
            const cleanProp = prop.replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
            if (cleanProp.includes(k.toLowerCase()) && row[prop] !== '') {
              return String(row[prop]).trim();
            }
          }
        }
        return '';
      };

      const name = getVal('name');
      const code = getVal('code');
      const trade = getVal('trade');
      const father = getVal('father', "s/o", 'husband');
      const dob = formatExcelDate(getVal('birth', 'dob'));
      const doe = formatExcelDate(getVal('enrolment', 'doe', 'enrol'));

      return {
        name,
        code,
        trade,
        father,
        dob,
        doe,
        idmark: '',
        rmpl: '403'
      };
    }).filter(emp => emp.name !== '' && isNaN(emp.name));

    res.json(formattedData);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server live at http://localhost:${PORT}`);
});