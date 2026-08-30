const express = require('express');
const cors = require('cors');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

// Date Formatter (Handles Excel Serial & Text Dates)
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

    // Dhoondo wo sheet jisme actual employee list ho
    let targetSheetName = workbook.SheetNames.find(name => 
      name.toUpperCase().includes('MASTER') || 
      name.toUpperCase().includes('DATA') || 
      name.toUpperCase().includes('CPL') ||
      name.toUpperCase().includes('CARD')
    );

    // Agar na mile toh sabse badi sheet uthao
    if (!targetSheetName) {
      let maxRows = 0;
      workbook.SheetNames.forEach(name => {
        const sheet = workbook.Sheets[name];
        const rows = xlsx.utils.sheet_to_json(sheet);
        if (rows.length > maxRows) {
          maxRows = rows.length;
          targetSheetName = name;
        }
      });
    }

    const sheet = workbook.Sheets[targetSheetName || workbook.SheetNames[0]];
    const rawData = xlsx.utils.sheet_to_json(sheet);

    // Universal Property Matcher
    const formattedData = rawData.map(emp => {
      const getVal = (...keys) => {
        for (let k of keys) {
          for (let prop in emp) {
            if (prop.trim().toLowerCase() === k.toLowerCase() && emp[prop] !== undefined && emp[prop] !== null) {
              return String(emp[prop]).trim();
            }
          }
        }
        return '';
      };

      const name = getVal('NAME', 'CPL NAME', 'NAME OF CPL', 'EMPLOYEE NAME', 'NAME OF EMPLOYEE');
      const code = getVal('CODE', 'CODE NO', 'CPL NO', 'REGN NO', 'REG NO', 'REGISTRATION NO');
      const father = getVal('FATHER', "FATHER'S NAME", "FATHERS NAME", "S/O", "HUSBAND NAME", "FNAME");
      const trade = getVal('TRADE', 'TRADE / DESIG', 'DESIGNATION', 'DESIG', 'TRADE/DESIG');
      const dob = formatExcelDate(getVal('DOB', 'DATE OF BIRTH', 'BIRTH'));
      const doe = formatExcelDate(getVal('DOE', 'DATE OF ENROLMENT', 'ENROLMENT', 'DOJ', 'DATE OF ENGAGEMENT'));
      const idmark = getVal('IDMARK', 'IDENTIFICATION MARK', 'IDENTIFICATION', 'ID MARK');
      const rmpl = getVal('RMPL', 'RMPL NO') || '403';

      return { name, code, father, trade, dob, doe, idmark, rmpl };
    }).filter(emp => emp.name !== ''); // Khali rows filter out

    res.json(formattedData);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server live at http://localhost:${PORT}`);
});