const express = require('express');
const bodyParser = require('body-parser');
const AdmZip = require('adm-zip');
const Handlebars = require('handlebars');
const fs = require('fs');

const app = express();

// Middleware to parse JSON requests
app.use(bodyParser.json());

// Function to fill the Word template
function fillWordTemplate(templatePath, outputPath, data) {
  const zip = new AdmZip(templatePath);
  const zipEntries = zip.getEntries();

  // Find the document.xml file in the Word archive
  const docEntry = zipEntries.find(entry => entry.entryName === 'word/document.xml');
  if (!docEntry) {
      throw new Error('Could not find document.xml in the Word template.');
  }

  // Read the document.xml content
  const docContent = docEntry.getData().toString('utf-8');
  // console.log('Original document.xml content: ', docContent);  // Debugging line

  // Compile and update the document content with Handlebars
  const template = Handlebars.compile(docContent);
  const updatedContent = template(data);

  // Debug the updated content before replacing it in the ZIP
  // console.log('Updated document.xml content: ', updatedContent);  // Debugging line

  // Replace the content in the ZIP archive
  zip.updateFile('word/document.xml', Buffer.from(updatedContent, 'utf-8'));

  // Save the updated document as a new file
  zip.writeZip(outputPath);
  console.log("end")
}


// API endpoint to generate the Word document
app.post('/generate', (req, res) => {
    const data = req.body; // Dynamic data from the request body
    const templatePath = './sample.docx'; // Path to the template file
    const outputPath = './output.docx'; // Path for the output file

    try {
        fillWordTemplate(templatePath, outputPath, data);

        // Set proper headers and send the file
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename="output.docx"');

        // Send the file for download
        console.log('Sending file to client');
        res.sendFile(outputPath, { root: __dirname }, (err) => {
            if (err) {
                console.error('Error sending file:', err);
                res.status(500).send('Could not generate document.');
            }

          
        });
    } catch (error) {
        console.error('Error generating document:', error);
        res.status(500).send('Error generating document.');
    }
});



// Start the server
const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});