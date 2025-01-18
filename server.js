const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const cors = require('cors');
const path = require('path');

const app = express();
const port = 5000;

// Middleware
app.use(cors());
app.use(express.json());

// Set up multer for file upload
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, './uploads');
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + path.extname(file.originalname));
  },
});

const upload = multer({ storage });

// Upload endpoint
app.post('/upload', upload.single('file'), (req, res) => {
  const filePath = path.join(__dirname, 'uploads', req.file.filename);

  // Read the uploaded Excel file
  const workbook = xlsx.readFile(filePath);

  // Log sheet names to ensure we are reading the correct sheet
  console.log('Sheet Names:', workbook.SheetNames);

  // Extract the first sheet
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  // Log raw sheet data to understand its structure
  console.log('Raw Sheet Data:', sheet);

  // Convert sheet to JSON with header row
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

  console.log('Parsed Data:', data); // Log the parsed data

  if (!data || data.length === 0) {
    return res.status(400).json({ error: 'No data found in the Excel sheet' });
  }

  // Clean the data (optional)
  const cleanedData = sanitizeData(data);
  console.log('Cleaned Data:', cleanedData); // Log cleaned data

  res.json(cleanedData);
});

// Function to clean non-printable characters and remove unnecessary spaces from the headers
const sanitizeData = (data) => {
  const headers = data[0].map(header => header.trim().replace(/\s+/g, '_')); // Sanitize headers
  const cleanedData = data.map((row, index) => {
    if (index === 0) return headers; // Replace headers with cleaned ones
    return row.map(cell => {
      if (typeof cell === 'string') {
        // Remove non-printable characters (optional)
        return cell.replace(/[^\x20-\x7E]/g, '');
      }
      return cell;
    });
  });
  return cleanedData;
};

// Start server
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
