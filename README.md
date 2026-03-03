# Word to PDF Converter

A simple web service that converts Word documents (.doc/.docx) into PDF files.

## How It Works

The conversion happens in three steps:

1. **You upload a Word file** — either through the web page (drag and drop) or by calling the API directly.
2. **mammoth** reads the Word file and converts its content into HTML.
3. **puppeteer** (a headless Chrome browser) takes that HTML and prints it to a PDF, just like using "Save as PDF" from a browser's print dialog.

The PDF is sent back to you as a download, and the temp files are deleted immediately.

```
.docx  ──►  mammoth  ──►  HTML  ──►  puppeteer  ──►  .pdf
```

## Quick Start

```bash
npm install
node server.js
```

Open http://localhost:3000 in your browser. That's it.

## API Usage

```bash
curl -X POST -F "file=@document.docx" http://localhost:3000/api/convert -o output.pdf
```

- **Endpoint**: `POST /api/convert`
- **Field name**: `file`
- **Accepted formats**: `.doc`, `.docx`
- **Max file size**: 50 MB
- **Response**: The converted PDF file

## Project Structure

```
word-to-pdf/
├── server.js         # Express server + conversion logic
├── public/
│   └── index.html    # Web UI (drag-and-drop upload page)
├── uploads/          # Temp storage for uploaded files (auto-cleaned)
└── converted/        # Temp storage for PDFs (auto-cleaned)
```

## What Each Dependency Does

| Package     | Purpose                                              |
|-------------|------------------------------------------------------|
| express     | Web server that handles HTTP requests                |
| multer      | Handles file uploads from the browser/API            |
| mammoth     | Reads .docx files and converts them to clean HTML    |
| puppeteer   | Runs a headless Chrome to render HTML into PDF       |
| uuid        | Generates unique filenames so uploads don't collide  |

## Deployment

See [DEPLOY.md](DEPLOY.md) for step-by-step instructions to deploy on an Azure Ubuntu VM using PM2.

## Limitations

- Best suited for simple documents (text, headings, tables, images, basic formatting).
- Complex Word features like tracked changes, embedded charts, or macros won't carry over.
- mammoth focuses on semantic content, not pixel-perfect layout matching.
