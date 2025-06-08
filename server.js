const express = require('express');
const mongoose = require('mongoose');
const multer = require('multer');
const cors = require('cors');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const authRoutes = require('./routes/auth');
require('dotenv').config();

const app = express();

// Middleware
app.use(cors());
app.use(express.json());

// Routes
app.use('/api/auth', authRoutes);

// MongoDB Connection
mongoose.connect(process.env.MONGODB_URI, {
  useNewUrlParser: true,
  useUnifiedTopology: true,
})
.then(() => console.log('MongoDB connected successfully'))
.catch(err => console.error('MongoDB connection error:', err));

// Mongoose Schema for Excel Data
const excelDataSchema = new mongoose.Schema({
  filename: String,
  uploadDate: { type: Date, default: Date.now },
  headers: [String],
  data: mongoose.Schema.Types.Mixed, // To store array of JSON objects
  rowCount: Number
});

const ExcelData = mongoose.model('ExcelData', excelDataSchema);

// Set up multer for file uploads
const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    const uploadDir = 'uploads/';
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir);
    }
    cb(null, uploadDir);
  },
  filename: function (req, file, cb) {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const fileFilter = (req, file, cb) => {
  // Accept excel files only
  if (
    file.mimetype === 'application/vnd.ms-excel' ||
    file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  ) {
    cb(null, true);
  } else {
    cb(new Error('Only Excel files are allowed!'), false);
  }
};

const upload = multer({ 
  storage: storage,
  fileFilter: fileFilter,
  limits: { fileSize: 1024 * 1024 * 5 } // 5MB limit
});

// Route to handle file upload
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please upload an Excel file' });
    }

    const filePath = req.file.path;
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetNames[0]]);

    // Get column headers (keys of the first object)
    const headers = jsonData.length > 0 ? Object.keys(jsonData[0]) : [];

    // Save to MongoDB
    const newExcelData = new ExcelData({
      filename: req.file.originalname,
      headers: headers,
      data: jsonData,
      rowCount: jsonData.length
    });

    await newExcelData.save();

    // Optionally, delete the file from uploads/ after saving to DB
    // fs.unlinkSync(filePath); 

    return res.status(200).json({
      message: 'File uploaded and data saved to MongoDB successfully',
      filename: req.file.originalname,
      headers: headers,
      data: jsonData,
      rowCount: jsonData.length
    });
  } catch (error) {
    console.error('Error processing file:', error);
    return res.status(500).json({ message: 'Error processing file', error: error.message });
  }
});

// Start server
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});