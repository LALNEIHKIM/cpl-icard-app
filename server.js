const express = require('express');
const cors = require('cors');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

// Excel Serial Date to DD/MM/YYYY Formatter
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

// API endpoint to fetch employees directly from Excel MASTERLIST
app.get('/api/employees', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'CPL I CARD MAKER.xlsx');
    const workbook = xlsx.readFile(filePath);
    
    // MASTERLIST sheet dhoondo ya pehli sheet lo
    const sheetName = workbook.SheetNames.find(s => s.toUpperCase().includes('MASTER')) || workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(sheet);

    // Clean and Format Fields
    const formattedData = rawData.map(emp => {
      // Keys case-insensitive mapping
      const getVal = (...keys) => {
        for (let k of keys) {
          const foundKey = Object.keys(emp).find(x => x.trim().toLowerCase() === k.toLowerCase());
          if (foundKey && emp[foundKey] !== undefined) return emp[foundKey];
        }
        return '';
      };

      const dobVal = getVal('DOB', 'DATE OF BIRTH', 'BIRTH');
      const doeVal = getVal('DOE', 'DATE OF ENROLMENT', 'ENROLMENT', 'DOJ');

      return {
        ...emp,
        NAME: getVal('NAME', 'CPL NAME', 'EMPLOYEE NAME'),
        CODE: getVal('CODE', 'CODE NO', 'CPL NO', 'REGN NO'),
        FNAME: getVal("FATHER'S NAME", "FATHERS NAME", "FNAME", "HUSBAND NAME"),
        DESIG: getVal('TRADE / DESIG', 'TRADE', 'DESIGNATION', 'DESIG'),
        DOB: formatExcelDate(dobVal),
        DOE: formatExcelDate(doeVal),
        IDMARK: getVal('IDENTIFICATION MARK', 'ID MARK', 'MARK'),
        RMPL: getVal('RMPL', 'RMPL NO')
      };
    });

    res.json(formattedData);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server live at http://localhost:${PORT}`);
});