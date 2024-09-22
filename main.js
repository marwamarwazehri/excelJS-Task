// const express = require('express')
// const app = express()

// app.get('/', function (req, res) {
//   res.send('Hello World ghh')
// })

// app.listen(3000);


/*
const express = require('express');:
This line uses the require function to import the Express module into your file.

require is a built-in Node.js function that allows you to include external 
modules (like libraries or your own code) in your application.

The express module is a web application framework for Node.js that simplifies 
the process of building web servers and APIs.

const app = express();:
This line creates an instance of an Express application.
The express() function returns an Express application object, which you can 
use to set up middleware, define routes, and handle HTTP requests.

*/

//////////////////////////////////////////////////////////////////////////////////////////////////

const express = require('express')
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const app = express();

app.use(bodyParser.json());//


//add data to the excel file
app.post('/add', async (req, res) => {

  
  const { name, age } = req.body;

  
  if (!name || !age) {
      return res.status(400).json({ error: 'Name and age are required' });
  }

  // Load or create the workbook
  const filePath = 'data.xlsx';  /*The value 'data.xlsx' indicates that the file will be created in the current working directory of your application unless you specify a different path (like './folder/data.xlsx'). */
  const workbook = new ExcelJS.Workbook();/*When you create an instance of ExcelJS.Workbook():
You're preparing to either read from an existing workbook or create a new one.*/
  
  let worksheet;
  try {
      await workbook.xlsx.readFile(filePath);
      worksheet = workbook.getWorksheet(1); // Get the first worksheet 
  } catch (error) {
      worksheet = workbook.addWorksheet('Sheet1'); // Create a new worksheet if it doesn't exist
      worksheet.columns = [{ header: 'Name', key: 'name' }, { header: 'Age', key: 'age' }];
  }

  

  // Add a new row with the data
  worksheet.addRow({ name, age });

  // Save the workbook
  try {
    await workbook.xlsx.writeFile(filePath);
    console.log('Data written successfully');
} catch (writeError) {
    console.error('Error writing file:', writeError);
    return res.status(500).json({ error: 'Failed to write to the Excel file' });
}
/*
  workbook.xlsx.writeFile(filePath):This method is provided by the exceljs library and is
   used to write the contents of the workbook to an Excel file.
 The filePath variable specifies the name (and location) of the file where the workbook will be saved. In this case, it's 'data.xlsx'.
   
 If data.xlsx already exists, the method will overwrite it with the current contents of 
 the workbook.

 If it doesn't exist, it will create a new file named data.xlsx and save the workbook's 
 data in that file.

 -The workbook itself is not an Excel file; it is an object in memory that represents the structure of an Excel file
  while you are manipulating it using the exceljs library.

  When you use await workbook.xlsx.writeFile(filePath);, you are instructing the
   library to take the in-memory representation (the workbook) and write it to 
  an actual file on disk (in this case, data.xlsx).
 */

  res.status(200).json({ message: 'Data added successfully!' });
});



//get data from the excel file
app.get('/data', async (req, res) => {
  const filePath = 'data.xlsx';
  const workbook = new ExcelJS.Workbook();

  try {
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1); // Get the first worksheet
      const data = [];

      // Loop through each row and push to the data array
      worksheet.eachRow((row, rowNumber) => {
          // Convert row values to an object
          data.push({
              name: row.getCell(1).value,
              age: row.getCell(2).value,
          });
      });

      res.status(200).json(data); // Send the data as a JSON response
  } catch (error) {
      console.error('Error reading file:', error);
      res.status(500).json({ error: 'Failed to read the Excel file' });
  }
});

//update age
app.put('/update', async (req, res) => {
  const { name, age } = req.body;

  if (!name || age === undefined) {
      return res.status(400).json({ error: 'Name and age are required' });
  }

  const filePath = 'data.xlsx';
  const workbook = new ExcelJS.Workbook();
  let worksheet;

  try {
      // Read the existing Excel file
      await workbook.xlsx.readFile(filePath);
      worksheet = workbook.getWorksheet(1); // Get the first worksheet
  } catch (error) {
      return res.status(404).json({ error: 'Data file not found' });
  }

  // Flag to check if the name was found
  let found = false;

  // Loop through the rows to find the matching name
  worksheet.eachRow((row, rowNumber) => {
      if (row.getCell(1).value === name) {
          // Update the age if the name matches
          row.getCell(2).value = age;
          found = true;
      }
  });

  if (!found) {
      return res.status(404).json({ error: 'Name not found' });
  }

  // Save the updated workbook
  await workbook.xlsx.writeFile(filePath);

  res.status(200).json({ message: 'Age updated successfully!' });
});

//delte row according to name
app.delete('/delete', async (req, res) => {
  const { name } = req.body;

  if (!name) {
      return res.status(400).json({ error: 'Name is required' });
  }

  const filePath = 'data.xlsx';
  const workbook = new ExcelJS.Workbook();
  let worksheet;

  try {
      // Read the existing Excel file
      await workbook.xlsx.readFile(filePath);
      worksheet = workbook.getWorksheet(1); // Get the first worksheet
  } catch (error) {
      return res.status(404).json({ error: 'Data file not found' });
  }

  // Flag to check if the name was found
  let found = false;

  // Loop through the rows to find the matching name
  worksheet.eachRow((row, rowNumber) => {
      if (row.getCell(1).value === name) {
          // Delete the row if the name matches
          worksheet.spliceRows(rowNumber, 1);
          found = true;
      }
  });

  if (!found) {
      return res.status(404).json({ error: 'Name not found' });
  }

  // Save the updated workbook
  await workbook.xlsx.writeFile(filePath);

  res.status(200).json({ message: 'Data deleted successfully!' });
});





const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});





