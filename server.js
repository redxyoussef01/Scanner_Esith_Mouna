const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;

const app = express();
const port = 5000;

// Enable CORS for all origins, methods, and headers
app.use(cors());
app.use(express.json()); // Middleware to parse JSON request bodies

let latestBarcode = null;

app.post('/barcodes', (req, res) => { // Changed endpoint to /barcodes
  const { barcode } = req.body;
  console.log('Received barcode:', barcode);
  latestBarcode = barcode; // Store the latest barcode
  res.json({ message: 'Barcode received successfully!' });
});

app.get('/api/get-latest-barcode', (req, res) => {
  res.json({ barcode: latestBarcode });
});

// New endpoint to nullify the QR code (in your data store)
app.post('/api/nullify-barcode', (req, res) => {
  const { barcode } = req.body;
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
  const logs = req.body;

  if (!Array.isArray(logs) || logs.length === 0) {
      return res.status(400).json({ error: 'Un tableau d\'entrées de journal est requis.' });
  }

  try {
      // Ensure the directory and file exist
      await ensureFileExists();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(excelFilePath);

      let worksheet = workbook.getWorksheet('Transactions');
      if (!worksheet) {
          worksheet = workbook.addWorksheet('Transactions');
          worksheet.addRow(['Type', 'Date', 'Time', 'Product', 'Quantity']); // Add headers if the sheet is new
          worksheet.getRow(1).eachCell((cell) => {
              cell.font = { bold: true };
              cell.alignment = { horizontal: 'center' };
          });
      }

      for (const log of logs) {
          const { type, product, quantity } = log;
          const currentDate = new Date().toLocaleDateString();
          const currentTime = new Date().toLocaleTimeString();

          if (!type || !product || quantity === undefined || isNaN(quantity)) {
              console.warn('Invalid log entry:', log);
              continue; // Skip invalid log entries
          }

          // Add the new transaction with basic styling
          const dataRow = worksheet.addRow([type, currentDate, currentTime, product, quantity]);
          dataRow.eachCell((cell) => {
              cell.alignment = { horizontal: 'center' };
          });
      }

      // Write the updated workbook back to the file
      await workbook.xlsx.writeFile(excelFilePath);

      res.json({ message: 'Journal des transactions mis à jour avec succès sur le serveur.' });

  } catch (error) {
      console.error('Erreur lors de la mise à jour du journal des transactions en lot :', error);
      res.status(500).json({ error: `Échec de la mise à jour du journal des transactions en lot: ${error.message}` });
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
        product: row.getCell(4).value,
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

async function ensureInventoryFileExists() {
    try {
        await fs.access(inventoryExcelFilePath);
    } catch (error) {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Inventory');
        // Add headers including the new columns
        worksheet.addRow(['ProductID', 'Name', 'Quantity', 'Daily Transactions', 'Last Transaction Date']);
        await workbook.xlsx.writeFile(inventoryExcelFilePath);
    }
}

// Initialize inventory file if it doesn't exist
ensureInventoryFileExists().catch(console.error);

// Function to reset daily transactions if the date has changed
async function resetDailyTransactionsIfNeeded(worksheet) {
    const today = new Date().toDateString();
    let headerRow = worksheet.getRow(1);
    const lastTransactionDateColumn = headerRow.values.indexOf('Last Transaction Date') + 1;

    worksheet.eachRow({ start: 2, includeEmpty: false }, (row) => {
        const lastDateValue = row.getCell(lastTransactionDateColumn).value;
        const lastTransactionDate = lastDateValue instanceof Date ? lastDateValue.toDateString() : lastDateValue;

        if (lastTransactionDate !== today) {
            row.getCell(4).value = 0; // Reset daily transactions
            row.getCell(lastTransactionDateColumn).value = today;
        }
    });
}

app.get('/api/inventory', async (req, res) => {
  try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(inventoryExcelFilePath);
      const worksheet = workbook.getWorksheet('Inventory');
      if (!worksheet) {
          return res.json([]);
      }

      await resetDailyTransactionsIfNeeded(worksheet);
      await workbook.xlsx.writeFile(inventoryExcelFilePath); // Save changes after reset

      const data = [];
      worksheet.eachRow({ start: 2, includeEmpty: false }, (row) => {
          data.push({
              productId: row.getCell(1).value,
              name: row.getCell(2).value,
              quantity: Number(row.getCell(3).value),
              dailyTransactions: Number(row.getCell(4).value) || 0, // Get daily transactions
          });
      });
      res.json(data);
  } catch (error) {
      console.error('Erreur lors de la récupération des données d\'inventaire :', error);
      res.status(500).json({ error: 'Échec de la récupération des données d\'inventaire.' });
  }
});

// POST /api/inventory - Add or update a product in the inventory
app.post('/api/inventory', async (req, res) => {
    const { productId, name, quantity } = req.body; // Expecting productId, name, and quantity

    if (!productId || !name || quantity === undefined || isNaN(quantity)) {
        return res.status(400).json({ error: 'L\'ID du produit, le nom et la quantité sont requis.' });
    }

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(inventoryExcelFilePath);
        const worksheet = workbook.getWorksheet('Inventory');

        await resetDailyTransactionsIfNeeded(worksheet); // Ensure daily transactions are reset if needed

        let productFound = false;
        worksheet.eachRow({ start: 2 }, (row) => {
            if (row.getCell(1).value === productId) {
                const existingQuantity = Number(row.getCell(3).value) || 0;
                row.getCell(3).value = existingQuantity + Number(quantity);
                row.getCell(4).value = (Number(row.getCell(4).value) || 0) + Number(quantity); // Increment daily transactions
                row.getCell(5).value = new Date().toDateString(); // Update last transaction date
                productFound = true;
            }
        });

        if (!productFound) {
            worksheet.addRow([productId, name, Number(quantity), Number(quantity), new Date().toDateString()]); // Add new row with initial daily transaction and date
        }

        await workbook.xlsx.writeFile(inventoryExcelFilePath);
        res.json({ message: 'L\'inventaire du produit a été mis à jour avec succès.' });
    } catch (error) {
        console.error('Erreur lors de l\'ajout/mise à jour du produit :', error);
        res.status(500).json({ error: 'Échec de l\'ajout/mise à jour du produit dans l\'inventaire.' });
    }
});

app.post('/api/update-inventory', async (req, res) => {
    const updates = req.body;

    if (!Array.isArray(updates) || updates.length === 0) {
        return res.status(400).json({ error: 'Un tableau de mises à jour de l\'inventaire est requis.' });
    }

    const errors = [];

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(inventoryExcelFilePath);
        const worksheet = workbook.getWorksheet('Inventory');

        await resetDailyTransactionsIfNeeded(worksheet); // Ensure daily transactions are reset if needed

        for (const update of updates) {
            const { type, product, quantity } = update;

            if (!type || !product || quantity === undefined || isNaN(quantity)) {
                errors.push({ product, error: 'Le type (Entrée/Sortie), l\'ID du produit et la quantité sont requis.' });
                continue; // Move to the next update
            }

            let productFound = false;
            worksheet.eachRow({ start: 2 }, (row) => {
                if (String(row.getCell(1).value) === String(product)) {
                    const currentQuantity = Number(row.getCell(3).value) || 0;
                    const currentDailyTransactions = Number(row.getCell(4).value) || 0;

                    if (type === 'Entree') {
                        row.getCell(3).value = currentQuantity + Number(quantity);
                        row.getCell(4).value = currentDailyTransactions + Number(quantity);
                        row.getCell(5).value = new Date().toDateString(); // Update last transaction date
                    } else if (type === 'Sortie') {
                        const newQuantity = currentQuantity - Number(quantity);
                        if (newQuantity >= 0) {
                            row.getCell(3).value = newQuantity;
                            row.getCell(4).value = currentDailyTransactions - Number(quantity);
                            row.getCell(5).value = new Date().toDateString(); // Update last transaction date
                        } else {
                            errors.push({ product, error: `Quantité insuffisante pour le produit ${product}.` });
                        }
                    }
                    productFound = true;
                }
            });

            if (!productFound && !errors.some(err => err.product === product)) {
                errors.push({ product, error: `Produit avec l'ID '${product}' non trouvé dans l'inventaire.` });
            }
        }

        await workbook.xlsx.writeFile(inventoryExcelFilePath);

        if (errors.length > 0) {
            return res.status(400).json({ errors, message: 'Certains produits n\'ont pas pu être mis à jour.' });
        }

        res.json({ message: 'Inventaire mis à jour avec succès pour tous les produits.' });

    } catch (error) {
        console.error('Erreur lors de la mise à jour de l\'inventaire en lot :', error);
        res.status(500).json({ error: 'Échec de la mise à jour de l\'inventaire en lot.' });
    }
});
app.listen(port, '0.0.0.0', () => {
  console.log(`Server listening on port ${port}`);
});