import express from 'express';
import multer from 'multer';
import { filterExcel } from './filterExcel.js';
import fs from 'fs';
import path from 'path';
import cors from 'cors';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const app = express();
const upload = multer({ dest: 'uploads/' });
const PORT = 5000;
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

app.use(express.static(path.join(__dirname, "../client/dist")))

app.get("/*", (req,res)=>{
  res.sendFile(path.join(__dirname, "../client/dist/index.html"))
})


app.post('/api/filter', upload.fields([
  { name: 'file', maxCount: 1 },
  { name: 'custData', maxCount: 1 }
]), async (req, res) => {
  if (!req.files || !req.files['file'] || !req.files['custData']) {
    return res.status(400).send('Both files must be uploaded');
  }
  const mainFilePath = req.files['file'][0].path;
  const custDataFilePath = req.files['custData'][0].path;
  try {
    const filteredBuffer = await filterExcel(mainFilePath, custDataFilePath);
    res.setHeader('Content-Disposition', 'attachment; filename="filtered.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(filteredBuffer);
  } catch (err) {
    res.status(500).send('Error filtering file: ' + err.message);
  } finally {
    // Clean up uploaded files
    if (mainFilePath) fs.unlink(mainFilePath, () => {});
    if (custDataFilePath) fs.unlink(custDataFilePath, () => {});
  }
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
}); 