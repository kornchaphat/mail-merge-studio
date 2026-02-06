# Mail Merge Studio

A Google Apps Script web app for generating personalized documents at scale. Connect a Google Sheets data source to a Google Docs or Slides template, and batch-generate PDFs, DOCX, or PPTX files — with conditional templates, subfolder organization, and configurable filename patterns.

> 🟢 **[Live Demo](https://script.google.com/macros/s/AKfycbx9Y5CikKwmz6rQmMIYrkALqtL8IP7UaFOL3amVhrngGi1HEShHNgMT0kJZm3kWb57tIQ/exec)** — Try it with your own Google Sheets data

## Screenshots

<!-- Upload screenshots to repo root, then replace filenames below -->
<p align="center">
  <img src="https://github.com/kornchaphat/mail-merge-studio/blob/61f13ddf39deb1dda3f437c83f08d532f6d26438/Screenshot%202026-02-07%20000525.png" width="700" alt="Step 1 - Data Source">
  <br><em>Step 1 — Define data source, select sheet, set header row, and filter values</em>
</p>
<p align="center">
  <img src="https://github.com/kornchaphat/mail-merge-studio/blob/61f13ddf39deb1dda3f437c83f08d532f6d26438/Screenshot%202026-02-07%20000658.png" width="700" alt="Step 2 - Templates">
  <br><em>Step 2 — Select one or multiple Google Docs/Slides templates with conditional rules</em>
</p>
<p align="center">
  <img src="https://github.com/kornchaphat/mail-merge-studio/blob/61f13ddf39deb1dda3f437c83f08d532f6d26438/Screenshot%202026-02-07%20000903.png" width="700" alt="Step 3 - Mapping">
  <br><em>Step 3 — Map template placeholders to spreadsheet column headers</em>
</p>
<p align="center">
  <img src="https://github.com/kornchaphat/mail-merge-studio/blob/61f13ddf39deb1dda3f437c83f08d532f6d26438/Screenshot%202026-02-07%20000950.png" width="700" alt="Step 4 - Configuration">
  <br><em>Step 4 — Configure output format, file naming pattern, and folder organization</em>
</p>
<p align="center">
  <img src="https://github.com/kornchaphat/mail-merge-studio/blob/61f13ddf39deb1dda3f437c83f08d532f6d26438/Screenshot%202026-02-07%20001024.png" width="700" alt="Step 5 - Preview & Generate">
  <br><em>Step 5 — Preview settings and generate documents</em>
</p>

## Features

### Data Source
- Connect any Google Sheets spreadsheet as data source
- Browse recent files or paste a spreadsheet URL
- Sheet and header row selection
- Row filtering with column-based conditions
- Live data preview with row count

### Templates
- Google Docs and Google Slides template support
- Auto-detect `{{placeholders}}` from template content
- Visual placeholder → column mapping interface
- Conditional template rules (use different templates per row based on column values)
- Multi-template support within a single merge

### Document Generation
- Output formats: PDF, DOCX, PPTX (based on template type)
- Configurable filename patterns with column variables and auto-numbering
- Replace existing files option
- Subfolder organization by column values (e.g., group by department)
- Output to any Google Drive folder
- ZIP download of all generated files

### Presets & History
- Save merge configurations as reusable presets
- Generation logging with timestamp, row count, and output location
- Usage statistics dashboard

### UX
- Step-by-step wizard interface (Data → Template → Configure → Generate)
- Glassmorphism UI with animated background
- Real-time progress tracking during generation
- Google SSO authentication (auto-detects logged-in user)

## Use Cases

- **HR**: Offer letters, employment contracts, training certificates, payslips
- **Sales**: Proposals, invoices, client reports
- **Education**: Certificates, report cards, personalized communications
- **Operations**: Shipping labels, work orders, compliance documents

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | HTML, CSS, JavaScript (1,850 lines, single-file) |
| Backend | Google Apps Script (1,100 lines) |
| Data | Google Sheets API |
| Templates | Google Docs API, Google Slides API |
| Storage | Google Drive API |
| Auth | Google SSO (Session.getActiveUser) |

## Architecture

```
┌──────────────────────────────────────────────────┐
│  Browser                                          │
│  ┌──────────────────────────────────────────────┐ │
│  │  MailMergeApp.html                           │ │
│  │  - 4-step wizard (Data → Template → Config)  │ │
│  │  - Glassmorphism UI + animated background    │ │
│  │  - google.script.run ←→ Backend API          │ │
│  └──────────────────────────────────────────────┘ │
└─────────────────────┬────────────────────────────┘
                      │ google.script.run
┌─────────────────────▼────────────────────────────┐
│  Google Apps Script Backend (Code.gs)             │
│  - Spreadsheet data reading + parsing             │
│  - Template placeholder extraction                │
│  - Document generation (Docs/Slides → PDF/DOCX)   │
│  - Drive folder management + ZIP creation         │
│  - Preset storage + generation logging            │
└─────────────────────┬────────────────────────────┘
                      │ Google APIs
┌─────────────────────▼────────────────────────────┐
│  Google Workspace                                 │
│  Sheets (data) │ Docs/Slides (templates)          │
│  Drive (output + ZIP) │ Script Properties (config)│
└──────────────────────────────────────────────────┘
```

## Project Structure

```
├── Code.gs              # Backend (1,100 lines)
├── MailMergeApp.html    # Frontend (1,850 lines)
└── README.md
```

## Setup

1. Create a new [Google Apps Script](https://script.google.com) project
2. Replace `Code.gs` with the backend code
3. Create `MailMergeApp.html` and paste the frontend code
4. **Deploy → New Deployment → Web app**
   - Execute as: User accessing the web app
   - Who has access: Anyone (or within your organization)
5. Open the deployment URL

No additional configuration needed — the app auto-creates its database sheet for presets and logs on first run.

## How It Works

1. **Select data source** — Pick a Google Sheets file, choose the sheet and header row
2. **Choose template** — Select a Google Docs or Slides file with `{{placeholder}}` markers
3. **Map fields** — Connect template placeholders to spreadsheet columns
4. **Generate** — Batch-create personalized documents, download as ZIP or find in Drive

## Author

Built by **Kornchaphat Piyatakoolkan**

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue)](https://www.linkedin.com/in/kornchaphat)
