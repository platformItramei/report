const { createPool } = require('@vercel/postgres');
const XLSX = require('xlsx');
const fs = require('fs');

const pool = createPool({
  connectionString: "postgres://default:jJWUhdwZl9v4@ep-restless-scene-a4jiwlc1-pooler.us-east-1.aws.neon.tech:5432/verceldb?sslmode=require"
});

async function createItrameiWaitlist(client) {
  const result = await client.query(
    "SELECT fullname,email,event,consent,createdAt,updatedAt FROM users WHERE email NOT IN ('Neda@itramei.com', 'ap425039@gmail.com', 'kristoffer.lettegard@outlook.com', 'kristoffer.bengt.lettegard@gmail.com', 'patil.ashish.19it5008@gmail.com', 'ashish.patil5@mail.dcu.ie') AND event='waitinglist'"
  );
  const data = result.rows;

  const worksheet = XLSX.utils.json_to_sheet(data);
  
  // Set custom column widths
  worksheet['!cols'] = [
    { wch: 20 }, 
    { wch: 25 }, 
    { wch: 15 }, 
    { wch: 10 }, 
    { wch: 20 }, 
    { wch: 20 }, 
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Users');

  const filePath = './itramei_waitinglist.xlsx';
  XLSX.writeFile(workbook, filePath);

  console.log(`Excel file created successfully at ${filePath}`);
}

async function createLaunchEventWaitlist(client) {
  const result = await client.query(
    "SELECT fullname,firstname,email,companyname,position,event,phonenumber,createdat,updatedat FROM users WHERE email NOT IN ('Neda@itramei.com', 'ap425039@gmail.com', 'kristoffer.lettegard@outlook.com', 'patil.ashish.19it5008@gmail.com', 'ashish.patil5@mail.dcu.ie') AND event='launch-waitlist'"
  );
  const data = result.rows;

  const worksheet = XLSX.utils.json_to_sheet(data);
  
  worksheet['!cols'] = [
    { wch: 20 },
    { wch: 20 }, 
    { wch: 25 }, 
    { wch: 25 }, 
    { wch: 20 }, 
    { wch: 15 }, 
    { wch: 15 }, 
    { wch: 20 },
    { wch: 20 }, 
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Users');

  const filePath = './itramei_launch_event_waitinglist.xlsx';
  XLSX.writeFile(workbook, filePath);

  console.log(`Excel file created successfully at ${filePath}`);
}

async function fetchDataAndExportToExcel() {
  const client = await pool.connect();
  try {
    console.log('Connected to PostgreSQL');
    
    await createItrameiWaitlist(client);
    await createLaunchEventWaitlist(client);
    
  } catch (error) {
    console.error('Error fetching data and exporting to Excel:', error);
    throw error;
  } finally {
    client.release(); // Release the client only once after both functions have completed
  }
}

module.exports = { pool, fetchDataAndExportToExcel };