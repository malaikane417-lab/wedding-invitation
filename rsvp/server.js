const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
const PORT = 3001;
const XLSX_PATH = path.join(__dirname, 'details.xlsx');

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

function appendToExcel(data) {
  let workbook;
  let ws;

  if (fs.existsSync(XLSX_PATH)) {
    workbook = XLSX.readFile(XLSX_PATH);
    ws = workbook.Sheets[workbook.SheetNames[0]];
  } else {
    workbook = XLSX.utils.book_new();
    ws = XLSX.utils.aoa_to_sheet([['Date', 'Full Name', 'Email', 'Phone', 'Guests', 'Attendance', 'Dietary', 'Message']]);
    XLSX.utils.book_append_sheet(workbook, ws, 'RSVPs');
  }

  const existingData = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const now = new Date().toLocaleString();
  existingData.push([now, data.name, data.email, data.phone, data.guests, data.attendance, data.dietary, data.message]);

  const newWs = XLSX.utils.aoa_to_sheet(existingData);
  workbook.Sheets[workbook.SheetNames[0]] = newWs;
  XLSX.writeFile(workbook, XLSX_PATH);
}

app.post('/rsvp', (req, res) => {
  try {
    const data = req.body;
    appendToExcel(data);
    console.log('RSVP received:', data.name, '-', data.attendance);
    res.json({ success: true, message: 'RSVP saved successfully' });
  } catch (err) {
    console.error('Error saving RSVP:', err);
    res.status(500).json({ success: false, message: 'Failed to save RSVP' });
  }
});

app.listen(PORT, () => {
  console.log('RSVP server running on http://localhost:' + PORT);
});
