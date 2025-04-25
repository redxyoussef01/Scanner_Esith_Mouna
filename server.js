const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;

const app = express();
const port = 5000;

app.use(cors());
app.use(express.json()); // Middleware to parse JSON request bodies



let latestBarcode = null;

app.post('/api/receive-barcode', (req, res) => {
  const { barcode } = req.body;
  console.log('Received barcode:', barcode);
  latestBarcode = barcode; // Store the latest barcode
  res.json({ message: 'Barcode received successfully!' });
});

app.get('/api/get-latest-barcode', (req, res) => {
  res.json({ barcode: latestBarcode });
});


// New endpoint to nullify the QR code (in your data store)
app.post('/api/nullify-barcode', (req, res) => { const { barcode } = req.body;
console.log('Nullifying barcode:', barcode);
latestBarcode = null; // Set the latest barcode to null
// In a real application, you would likely update your database here
// to mark this barcode as processed or null.
res.json({ message: `Barcode "${barcode}" has been nullified.` });
});

const excelFilePath = path.join(__dirname, 'excel', 'transactionLog.xlsx');

// Create directory and Excel file if they don't exist
async function ensureFileExists() {
  try {
    // Ensure directory exists
    await fs.mkdir(path.dirname(excelFilePath), { recursive: true });
    
    // Check if file exists
    try {
      await fs.access(excelFilePath);
      // File exists, no need to create it
    } catch (error) {
      // File doesn't exist, create it with headers
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Transactions');
      
      // Add and style headers
      const headerRow = worksheet.addRow(['Type', 'Date', 'Time', 'Produit', 'Nombre']);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, size: 12 };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFD3D3D3' } // Light gray background
        };
        cell.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
        cell.alignment = { horizontal: 'center' };
      });
      
      await workbook.xlsx.writeFile(excelFilePath);
    }
  } catch (error) {
    console.error('Error ensuring file exists:', error);
  }
}

// Initialize the app by ensuring file exists
ensureFileExists().catch(console.error);

app.post('/api/update-log', async (req, res) => {
  const { type, product, quantity } = req.body;
  const currentDate = new Date().toLocaleDateString();
  const currentTime = new Date().toLocaleTimeString();
  
  try {
    // Ensure the directory and file exist
    await ensureFileExists();
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);
    
    const worksheet = workbook.getWorksheet('Transactions');
    if (!worksheet) {
      return res.status(500).json({ error: 'Worksheet not found after creation' });
    }
    
    // Add the new transaction with basic styling
    const dataRow = worksheet.addRow([type, currentDate, currentTime, product, quantity]);
    dataRow.eachCell((cell) => {
      cell.alignment = { horizontal: 'center' };
    });
    
    // Write the updated workbook back to the file
    await workbook.xlsx.writeFile(excelFilePath);
    
    res.json({ message: 'Transaction log updated successfully on the server.' });
  } catch (error) {
    console.error('Error updating transaction log:', error);
    res.status(500).json({ error: `Failed to update transaction log: ${error.message}` });
  }
});

app.get('/api/transaction-log', async (req, res) => {
  try {
    // Check if file exists before trying to read it
    try {
      await fs.access(excelFilePath);
    } catch (error) {
      // File doesn't exist, return empty array
      return res.json([]);
    }
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);
    
    const worksheet = workbook.getWorksheet('Transactions');
    if (!worksheet) {
      return res.json([]); // Return empty array if the worksheet doesn't exist
    }
    
    const data = [];
    
    // We start from row 2 to skip the header row
    worksheet.eachRow({ start: 2, includeEmpty: false }, (row) => {
      // Directly map to the expected format
      const rowData = {
        type: row.getCell(1).value,
        product: Number(row.getCell(4).value),
        quantity: Number(row.getCell(5).value),
        timestamp: null // Will be set below
      };
      
      // Safely create timestamp from date and time cells
      try {
        const dateValue = row.getCell(2).value;
        const timeValue = row.getCell(3).value;
        
        // Handle different date/time formats
        let dateStr = typeof dateValue === 'string' ? dateValue : 
                    dateValue instanceof Date ? dateValue.toLocaleDateString() : '';
        let timeStr = typeof timeValue === 'string' ? timeValue : 
                    timeValue instanceof Date ? timeValue.toLocaleTimeString() : '';
        
        if (dateStr && timeStr) {
          // Create a new Date object safely
          const dateObj = new Date(`${dateStr} ${timeStr}`);
          rowData.timestamp = !isNaN(dateObj.getTime()) ? dateObj.toISOString() : null;
        }
      } catch (dateError) {
        console.error('Error parsing date/time:', dateError);
        // Use current timestamp as fallback
        rowData.timestamp = new Date().toISOString();
      }
      
      data.push(rowData);
    });
    
    res.json(data);
  } catch (error) {
    console.error('Error fetching transaction log:', error);
    res.status(500).json({ error: `Failed to fetch transaction log: ${error.message}` });
  }
});

const inventoryExcelFilePath = path.join(__dirname, 'excel', 'productInventory.xlsx');

// Ensure the /excel directory exists
fs.mkdir(path.dirname(inventoryExcelFilePath), { recursive: true }).catch(console.error);

app.post('/api/add-inventory', async (req, res) => {
  const { documentInventaire, magazin } = req.body;

  if (!documentInventaire || typeof magazin !== 'number') {
    return res.status(400).json({ error: 'Document Inventaire and Magazin (number) are required.' });
  }

  try {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
      await workbook.xlsx.readFile(inventoryExcelFilePath);
      worksheet = workbook.getWorksheet('Inventory') || workbook.addWorksheet('Inventory');
    } catch (error) {
      worksheet = workbook.addWorksheet('Inventory');
      const headerRow = worksheet.addRow(['Document Inventaire', 'Magazin']);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, size: 12 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD3D3D3' } };
        cell.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
        cell.alignment = { horizontal: 'center' };
      });
    }

    const existingRow = worksheet.getRow(1).values.includes('Document Inventaire') ?
      worksheet.addRow([documentInventaire, magazin]) :
      worksheet.addRow(['Document Inventaire', 'Magazin']).addRow([documentInventaire, magazin]);

    existingRow.eachCell((cell) => {
      cell.alignment = { horizontal: 'center' };
    });

    await workbook.xlsx.writeFile(inventoryExcelFilePath);
    res.json({ message: 'Product inventory added successfully.' });

  } catch (error) {
    console.error('Error adding product to inventory:', error);
    res.status(500).json({ error: 'Failed to add product to inventory.' });
  }
});

app.get('/api/inventory-data', async (req, res) => {
  try {
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.readFile(inventoryExcelFilePath);
    } catch (readError) {
      console.error('Error reading inventory Excel file:', readError);
      return res.status(500).json({ error: 'Could not read inventory data file.' });
    }

    const worksheet = workbook.getWorksheet('Inventory');
    if (!worksheet) {
      return res.json([]);
    }

    const headerRow = worksheet.getRow(1);
    const headers = headerRow.values.filter(value => value !== null);
    const data = [];

    worksheet.eachRow({ start: 2, includeEmpty: false }, (row, rowNumber) => {
      const rowData = {};
      row.eachCell((cell, colNumber) => {
        if (headers[colNumber - 1]) {
          rowData[headers[colNumber - 1].toLowerCase().replace(/ /g, '')] = cell.value;
        }
      });
      data.push(rowData);
    });

    res.json(data);

  } catch (error) {
    console.error('Error fetching inventory data:', error);
    res.status(500).json({ error: 'Failed to fetch inventory data.' });
  }
});

app.listen(port, () => {
  console.log(`Server listening on port ${port}`);
});