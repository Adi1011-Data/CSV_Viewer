# File Viewer

A browser-based app to open CSV files, view them in an editable table, adjust layout, and download your changes.

## Features

- **Drag & drop** or click to open `.csv, .xlsx, .xls` files (multiple files open in separate tabs)
- **Tabs** — each file opens in its own tab; switch between files without leaving the page
- **Table view** with row numbers and editable cells (click to edit)
- **Resize columns** by dragging the right border of any cell (like a spreadsheet)
- **Resize rows** by dragging the bottom border of any cell
- **Font size** slider in the header
- **Download** saves the edited table as a new CSV file (`*_edited.csv`)
- **JLC Format** — choose **BOM** or **Pick & Place** in the toolbar; a dialog opens with **Download JLC CSV** (auto column mapping; see `JLC_Format_docs` samples)

## Quick start
Use the web app directly here:  
[Open Web App](https://adi1011-data.github.io/FileViewer/)

No download or installation is required.

No build step required. Open `index.html` in a modern browser, or serve the folder locally:

```bash
# Python
cd csv_viewer
python -m http.server 8080

# Node (npx)
npx serve .
```

Then visit `http://localhost:8080`.

## Usage

1. Drop a file onto the upload area (or click to browse).
2. Edit cell values directly in the table.
3. Use the font size slider in the header if needed.
4. Drag a column’s right border or a row’s bottom border to resize.
5. Click **Download CSV** to save your edited file.
6. Choose **BOM** or **Pick & Place** from **JLC Format** — in the dialog, click **Download JLC CSV**.
7. Click **New tab** beside your tabs (or drop files on the workspace) to open more CSVs.
8. Close a tab with **×**; when all tabs are closed, the upload screen returns.

## Project structure

```
FileViewer/
├── index.html
├── css/styles.css
├── js/main.js
├── js/jlc-format.js
├── JLC_Format_docs/     # BOM & Pick-and-Place sample templates
└── README.md
```

## Browser support

Works in current Chrome, Firefox, Edge, and Safari. Requires JavaScript enabled.
