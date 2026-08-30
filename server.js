const express = require('express');
const cors = require('cors');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static(path.join(__dirname, 'public')));

// Excel Date to DD/MM/YYYY Formatter
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

// API endpoint to fetch employees directly from Excel
app.get('/api/employees', (req, res) => {
  try {
    const filePath = path.join(__dirname, 'CPL I CARD MAKER.xlsx');
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rawData = xlsx.utils.sheet_to_json(sheet);

    // Format fields cleanly
    const formattedData = rawData.map(emp => ({
      ...emp,
      DOB: formatExcelDate(emp.DOB),
      DOE: formatExcelDate(emp.DOE)
    }));

    res.json(formattedData);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server live at http://localhost:${PORT}`);
});