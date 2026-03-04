const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

const ROOT = __dirname;
const outputPdf = path.join(ROOT, 'Word_to_PDF_converter_IOT_Hitesh3517.pdf');

const imageFiles = ['s1.png', 's2.png', 's3.png', 's4.png', 's5.png', 's6.png', 's7.png']
  .map((name) => path.join(ROOT, name))
  .filter((p) => fs.existsSync(p));

if (imageFiles.length === 0) {
  throw new Error('No screenshots found. Expected s1.png to s7.png in project root.');
}

function mimeType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.jpg' || ext === '.jpeg') return 'image/jpeg';
  if (ext === '.webp') return 'image/webp';
  return 'image/png';
}

function toDataUri(filePath) {
  const base64 = fs.readFileSync(filePath).toString('base64');
  return `data:${mimeType(filePath)};base64,${base64}`;
}

const captions = [
  'Home interface of the Word to PDF Converter.',
  'Sample source content prepared before conversion.',
  'File picker dialog used to select .docx input.',
  'Selected file displayed and ready for conversion.',
  'Successful conversion status shown after processing.',
  'Generated PDF opened in browser viewer.',
  'Additional workflow verification screenshot.'
];

const images = imageFiles.map((file, i) => ({
  src: toDataUri(file),
  caption: `Figure ${i + 1}: ${captions[i] || 'Project screenshot.'}`
}));

function renderRows(items) {
  return items.map((item) => `
    <div class="shot-row">
      <img src="${item.src}" alt="${item.caption}">
      <div class="cap">${item.caption}</div>
    </div>
  `).join('');
}

const rows12 = renderRows(images.slice(0, 2));
const rows34 = renderRows(images.slice(2, 4));
const rows56 = renderRows(images.slice(4, 6));
const row7 = renderRows(images.slice(6, 7));

const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Word to PDF Converter Report</title>
  <style>
    @page { size: A4; margin: 0; }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Times New Roman", Times, serif;
      color: #111;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
      background: #fff;
    }
    .page {
      width: 210mm;
      height: 297mm;
      padding: 12mm;
      page-break-after: always;
      background: #fff;
    }
    .page:last-child { page-break-after: auto; }
    .frame {
      width: 100%;
      height: 100%;
      border: 2px solid #2d2d2d;
      padding: 9mm 10mm;
      overflow: hidden;
      background: #fff;
    }
    .cover .frame {
      border-color: #3f5f91;
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      padding: 12mm;
    }
    .uni-heading {
      text-align: center;
      color: #123c7a;
      margin-bottom: 8mm;
      line-height: 1.25;
    }
    .uni-heading .u1 { font-size: 20pt; font-weight: 700; letter-spacing: 0.2px; }
    .uni-heading .u2 { font-size: 15pt; font-weight: 700; margin-top: 2mm; }
    .cover-rule {
      width: 92%;
      border-top: 2px solid #5e7fae;
      margin: 0 0 11mm;
    }
    .cover-title {
      text-align: center;
      color: #1f4a8a;
      line-height: 1.2;
      margin-bottom: 10mm;
    }
    .cover-title .a { font-size: 18pt; font-weight: 700; }
    .cover-title .b { font-size: 14pt; font-weight: 700; margin-top: 3px; }
    .cover-title .c { font-size: 16pt; font-weight: 700; margin-top: 4px; }
    .cover-table {
      width: 94%;
      border-collapse: collapse;
      font-size: 12pt;
    }
    .cover-table td {
      border: 1.5px solid #7d8ca3;
      padding: 7px 9px;
      vertical-align: top;
    }
    .cover-table td:first-child {
      width: 38%;
      background: #e6e9ef;
      font-weight: 700;
    }
    h2 {
      margin: 0 0 3mm;
      font-size: 16pt;
      border-bottom: 1px solid #333;
      padding-bottom: 1.5mm;
    }
    h3 {
      margin: 0 0 2mm;
      font-size: 13pt;
    }
    p {
      margin: 2.1mm 0;
      font-size: 12pt;
      line-height: 1.42;
      text-align: justify;
    }
    ul {
      margin: 1.5mm 0 0 6mm;
      padding: 0;
      font-size: 12pt;
      line-height: 1.42;
    }
    li { margin: 1.3mm 0; }
    .toc {
      margin-top: 3mm;
      width: 100%;
      border-collapse: collapse;
      font-size: 12pt;
      border: 1.5px solid #59657a;
    }
    .toc th, .toc td {
      border: 1px solid #7b879b;
      padding: 6px 8px;
    }
    .toc th {
      background: #e8edf5;
      text-align: left;
      font-size: 12pt;
    }
    .toc td:last-child, .toc th:last-child {
      text-align: center;
      width: 16%;
      font-weight: 700;
    }
    .shot-row {
      border: 1px solid #b6bcc7;
      padding: 2.5mm;
      background: #fbfcff;
      margin-bottom: 4mm;
    }
    .shot-row img {
      width: 100%;
      height: auto;
      max-height: 105mm;
      object-fit: contain;
      border: 1px solid #c9cdd6;
      display: block;
      background: #fff;
    }
    .cap {
      margin-top: 1.8mm;
      font-size: 10.2pt;
      font-style: italic;
      text-align: center;
    }
  </style>
</head>
<body>
  <section class="page cover">
    <div class="frame">
      <div class="uni-heading">
        <div class="u1">ANNA UNIVERSITY</div>
        <div class="u2">MIT CAMPUS</div>
      </div>
      <div class="cover-rule"></div>
      <div class="cover-title">
        <div class="a">PROJECT REPORT</div>
        <div class="b">ON</div>
        <div class="c">Word to PDF Converter</div>
      </div>
      <div class="cover-rule" style="margin: 0 0 8mm;"></div>
      <table class="cover-table">
        <tr><td>Student Name</td><td>V.Hitesh</td></tr>
        <tr><td>Roll Number</td><td>2024503517</td></tr>
        <tr><td>Subject</td><td>IOT Project</td></tr>
        <tr><td>Department</td><td>Computer Science</td></tr>
        <tr><td>Technology Stack</td><td>Node.js, Express.js, HTML, CSS, JavaScript, Multer, Mammoth, Puppeteer</td></tr>
        <tr><td>Academic Year</td><td>2025-2026</td></tr>
        <tr><td>Date</td><td>March 4, 2026</td></tr>
      </table>
    </div>
  </section>

  <section class="page">
    <div class="frame">
      <h2>CONTENTS</h2>
      <table class="toc">
        <tr><th>Section</th><th>Page No.</th></tr>
        <tr><td>1. Abstract</td><td>3</td></tr>
        <tr><td>2. Introduction</td><td>3</td></tr>
        <tr><td>3. Problem Statement</td><td>3</td></tr>
        <tr><td>4. Objectives</td><td>3</td></tr>
        <tr><td>5. Technology Stack</td><td>3</td></tr>
        <tr><td>6. System Architecture</td><td>3</td></tr>
        <tr><td>7. Implementation Details</td><td>4</td></tr>
        <tr><td>8. Workflow</td><td>4</td></tr>
        <tr><td>9. API/Module Description</td><td>4</td></tr>
        <tr><td>10. Testing and Results</td><td>4</td></tr>
        <tr><td>11. Screenshots with Explanations</td><td>5-8</td></tr>
        <tr><td>12. Challenges and Limitations</td><td>8</td></tr>
        <tr><td>13. Future Enhancements</td><td>8</td></tr>
        <tr><td>14. Conclusion</td><td>8</td></tr>
        <tr><td>15. References</td><td>8</td></tr>
      </table>
    </div>
  </section>

  <section class="page">
    <div class="frame">
      <h2>1. Abstract</h2>
      <p>This project implements a web-based Word to PDF Converter that allows users to upload .doc or .docx files and receive downloadable PDFs quickly. The solution is designed for simple interaction, clean output, and practical usage in academic and professional scenarios.</p>
      <h2>2. Introduction</h2>
      <p>Word documents are easy to edit, while PDFs are ideal for sharing. This project bridges that gap by providing a browser-accessible converter that handles document upload, conversion, and download in a single workflow.</p>
      <h2>3. Problem Statement</h2>
      <p>Many conversion tools require local software installation, subscriptions, or complex steps. The objective is to create a lightweight online utility that converts standard Word files to PDF efficiently.</p>
      <h2>4. Objectives</h2>
      <ul>
        <li>Accept .doc/.docx uploads and convert them to PDF.</li>
        <li>Provide drag-and-drop plus file browser upload modes.</li>
        <li>Show clear conversion feedback and downloadable output.</li>
        <li>Maintain reliable backend processing and cleanup.</li>
      </ul>
      <h2>5. Technology Stack</h2>
      <ul>
        <li>Frontend: HTML, CSS, JavaScript</li>
        <li>Backend: Node.js, Express.js</li>
        <li>Upload handling: Multer</li>
        <li>Word parsing: Mammoth</li>
        <li>PDF rendering: Puppeteer</li>
      </ul>
      <h2>6. System Architecture</h2>
      <p>Client sends document to the server, the server validates input, Mammoth converts Word to HTML, Puppeteer renders HTML to PDF, and the PDF is returned as download.</p>
    </div>
  </section>

  <section class="page">
    <div class="frame">
      <h2>7. Implementation Details</h2>
      <ul>
        <li>Unique filenames are generated to avoid collisions.</li>
        <li>Type and size validations prevent invalid uploads.</li>
        <li>PDF is generated in A4 format with print-friendly margins.</li>
        <li>Temporary files are cleaned after conversion.</li>
      </ul>
      <h2>8. Workflow</h2>
      <ul>
        <li>Open the converter page.</li>
        <li>Select or drag a Word file.</li>
        <li>Click "Convert to PDF".</li>
        <li>Receive and open the generated PDF.</li>
      </ul>
      <h2>9. API/Module Description</h2>
      <p><b>Endpoint:</b> POST /api/convert</p>
      <p><b>Input:</b> multipart/form-data with field <code>file</code></p>
      <p><b>Output:</b> downloadable PDF or structured error response.</p>
      <h2>10. Testing and Results</h2>
      <ul>
        <li>Upload and selection behavior verified.</li>
        <li>Conversion requests processed successfully.</li>
        <li>PDF download confirmed and output reviewed.</li>
      </ul>
    </div>
  </section>

  <section class="page">
    <div class="frame">
      <h2>11. Screenshots with Explanations (Page 1)</h2>
      ${rows12}
    </div>
  </section>

  <section class="page">
    <div class="frame">
      <h2>11. Screenshots with Explanations (Page 2)</h2>
      ${rows34}
    </div>
  </section>

  <section class="page">
    <div class="frame">
      <h2>11. Screenshots with Explanations (Page 3)</h2>
      ${rows56}
    </div>
  </section>

  <section class="page">
    <div class="frame">
      <h2>11. Screenshots with Explanations (Page 4)</h2>
      ${row7}
      <h2 style="margin-top:3mm;">12. Challenges and Limitations</h2>
      <ul>
        <li>Complex Word formatting may vary in rendered output.</li>
        <li>Insecure HTTP deployments can trigger browser download warnings.</li>
      </ul>
      <h2>13. Future Enhancements</h2>
      <ul>
        <li>Batch conversion and user history.</li>
        <li>Cloud storage integration and PDF customization options.</li>
      </ul>
      <h2>14. Conclusion</h2>
      <p>The project delivers a complete and practical Word-to-PDF conversion workflow with a user-friendly interface and stable backend conversion pipeline.</p>
      <h2>15. References</h2>
      <ul>
        <li>Node.js, Express.js, Multer, Mammoth, and Puppeteer documentation.</li>
      </ul>
    </div>
  </section>
</body>
</html>`;

(async () => {
  const browser = await puppeteer.launch({
    headless: true,
    executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu']
  });
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle0' });
  await page.pdf({
    path: outputPdf,
    format: 'A4',
    printBackground: true,
    preferCSSPageSize: true,
    margin: { top: '0', right: '0', bottom: '0', left: '0' }
  });
  await browser.close();
  console.log(outputPdf);
})();
