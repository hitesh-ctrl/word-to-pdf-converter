const express = require("express");
const multer = require("multer");
const { v4: uuidv4 } = require("uuid");
const path = require("path");
const fs = require("fs");
const mammoth = require("mammoth");
const puppeteer = require("puppeteer");

const app = express();
const PORT = process.env.PORT || 3000;

const UPLOADS_DIR = path.join(__dirname, "uploads");
const CONVERTED_DIR = path.join(__dirname, "converted");

// Ensure temp directories exist
fs.mkdirSync(UPLOADS_DIR, { recursive: true });
fs.mkdirSync(CONVERTED_DIR, { recursive: true });

// Clean up leftover files from previous runs
function cleanDir(dir) {
  for (const file of fs.readdirSync(dir)) {
    fs.unlinkSync(path.join(dir, file));
  }
}
cleanDir(UPLOADS_DIR);
cleanDir(CONVERTED_DIR);

// Launch browser once at startup for faster conversions
let browser;
(async () => {
  
browser = await puppeteer.launch({
  headless: true,
  executablePath: '/snap/bin/chromium',
  args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu']
});
  console.log("Puppeteer browser launched");
})();

// Multer config: accept only .docx, max 50MB
const storage = multer.diskStorage({
  destination: UPLOADS_DIR,
  filename: (_req, file, cb) => {
    const uniqueName = `${uuidv4()}${path.extname(file.originalname)}`;
    cb(null, uniqueName);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (_req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext !== ".docx" && ext !== ".doc") {
      return cb(new Error("Only .doc and .docx files are allowed"));
    }
    cb(null, true);
  },
});

// Serve static Web UI
app.use(express.static(path.join(__dirname, "public")));

// Convert endpoint
app.post("/api/convert", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "No file uploaded" });
  }

  const inputPath = req.file.path;
  const baseName = path.basename(inputPath, path.extname(inputPath));
  const outputPath = path.join(CONVERTED_DIR, `${baseName}.pdf`);

  try {
    await convertToPdf(inputPath, outputPath);

    const originalName = path.basename(
      req.file.originalname,
      path.extname(req.file.originalname)
    );

    res.download(outputPath, `${originalName}.pdf`, (err) => {
      cleanup(inputPath, outputPath);
      if (err && !res.headersSent) {
        res.status(500).json({ error: "Failed to send file" });
      }
    });
  } catch (err) {
    cleanup(inputPath, outputPath);
    res.status(500).json({ error: err.message || "Conversion failed" });
  }
});

async function convertToPdf(inputPath, outputPath) {
  // Step 1: Convert .docx to HTML using mammoth
  const { value: html } = await mammoth.convertToHtml({ path: inputPath });

  // Step 2: Wrap HTML in a styled page
  const fullHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        body {
          font-family: Arial, Helvetica, sans-serif;
          font-size: 12pt;
          line-height: 1.6;
          color: #000;
          max-width: 100%;
          padding: 0;
          margin: 0;
        }
        table {
          border-collapse: collapse;
          width: 100%;
          margin: 1em 0;
        }
        td, th {
          border: 1px solid #999;
          padding: 6px 10px;
          text-align: left;
        }
        img {
          max-width: 100%;
          height: auto;
        }
        h1 { font-size: 20pt; }
        h2 { font-size: 16pt; }
        h3 { font-size: 14pt; }
        p { margin: 0.5em 0; }
      </style>
    </head>
    <body>${html}</body>
    </html>
  `;

  // Step 3: Render HTML to PDF using puppeteer
  const page = await browser.newPage();
  try {
    await page.setContent(fullHtml, { waitUntil: "networkidle0" });
    await page.pdf({
      path: outputPath,
      format: "A4",
      margin: { top: "20mm", right: "15mm", bottom: "20mm", left: "15mm" },
      printBackground: true,
    });
  } finally {
    await page.close();
  }
}

function cleanup(...files) {
  for (const f of files) {
    try {
      if (fs.existsSync(f)) fs.unlinkSync(f);
    } catch {
      // ignore cleanup errors
    }
  }
}

// Multer error handling
app.use((err, _req, res, _next) => {
  if (err instanceof multer.MulterError) {
    if (err.code === "LIMIT_FILE_SIZE") {
      return res.status(413).json({ error: "File too large (max 50MB)" });
    }
    return res.status(400).json({ error: err.message });
  }
  if (err) {
    return res.status(400).json({ error: err.message });
  }
});

// Graceful shutdown
process.on("SIGINT", async () => {
  if (browser) await browser.close();
  process.exit(0);
});

app.listen(PORT, () => {
  console.log(`Word-to-PDF service running at http://localhost:${PORT}`);
});
