const express = require('express');
const bodyParser = require('body-parser');
const AdmZip = require('adm-zip');
const Handlebars = require('handlebars');
const fs = require('fs');

const app = express();

app.use(bodyParser.json());

function fillWordTemplate(templatePath, outputPath, data) {
  const zip = new AdmZip(templatePath);
  const zipEntries = zip.getEntries();

  const docEntry = zipEntries.find(entry => entry.entryName === 'word/document.xml');
  if (!docEntry) {
      throw new Error('Could not find document.xml in the Word template.');
  }

  const docContent = docEntry.getData().toString('utf-8');

  const template = Handlebars.compile(docContent);
  const updatedContent = template(data);
  zip.updateFile('word/document.xml', Buffer.from(updatedContent, 'utf-8'));

  zip.writeZip(outputPath);
  console.log("end")
}

app.post('/generate', (req, res) => {
    const data = req.body;
    const templatePath = './sample.docx'; 
    const outputPath = './output.docx';

    try {
        fillWordTemplate(templatePath, outputPath, data);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename="output.docx"');

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


const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});