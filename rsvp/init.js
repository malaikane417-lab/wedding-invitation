const XLSX = require('xlsx');
const path = require('path');

const ws = XLSX.utils.aoa_to_sheet([['Date', 'Full Name', 'Email', 'Phone', 'Guests', 'Attendance', 'Dietary', 'Message']]);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'RSVPs');
XLSX.writeFile(wb, path.join(__dirname, 'details.xlsx'));
console.log('Created details.xlsx');
