const express = require('express');
const { fetchDataAndExportToExcel } = require('./service'); // Adjust the path as needed

const app = express();
const PORT = 3005;

app.get('/', (req, res) => {
  res.send('Server is running');
});

// Automatically generate the Excel file when the server starts
async function generateExcelOnStart() {
  try {
    await fetchDataAndExportToExcel();
    process.exit();
  } catch (error) {
    console.error('Error exporting data on startup:', error);
  }
}

// Start the server and call the function
app.listen(PORT, async () => {
  console.log(`Server is running on port ${PORT}`);
  await generateExcelOnStart(); // Call the function here
});