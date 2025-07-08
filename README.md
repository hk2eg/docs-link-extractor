# Link Extractor

A lightweight, cross-platform utility that extracts hyperlinks from PDF, Excel, and CSV files. Designed to run either as a standard GUI app for Windows.

## Supported Formats:

- **PDF** The program is built assuming that the first PDF page containing the links is tabular. (manual table‐region selection)
- **Excel** (`.xlsx` with embedded hyperlinks or `HYPERLINK` formulas)
- **CSV** (plain-text URLs or hyperlinks)

When launched normally, it opens a Tkinter GUI. When launched with `--silent`, it runs only in the system tray, remaining hidden until the user chooses “Open Link Extractor” from the tray menu.

---

## Table of Contents

1. [Features](#features)
2. [Screenshots](#screenshots)
3. [Dependencies](#dependencies)  
4. [Installation](#installation)  
5. [Usage](#usage)  
   - [Standard GUI Mode](#standard-gui-mode)  
   - [Silent Tray Mode](#silent-tray-mode)  
6. [Project Structure](#project-structure)  
7. [Configuration & Saving](#configuration--saving)  
8. [Known Limitations](#known-limitations)  
9. [Future Improvements](#future-improvements)  
10. [License](#license)

---

## Features

- **PDF Table Extraction**  
  - Render the first page of any PDF at 150 DPI.  
  - Draw red rectangles around each “table” region.  
  - Intersect user-drawn regions with embedded hyperlink annotations and extract only those URLs belonging to each region.

- **Excel (`.xlsx`) Link Extraction**  
  - Detect links in cell values (plain text or `=HYPERLINK(...)` formulas).  
  - Detect “right-click → Insert Hyperlink” targets stored in `cell.hyperlink.target`.

- **CSV Link Extraction**  
  - Scan every cell of a CSV for patterns matching `https?://…` and collect all found URLs.

- **System Tray Integration**  
  - Uses `pystray` + `Pillow` to show a tray icon.  
  - Tray menu: “Open Link Extractor” and “Quit.”  
  - Can start minimized (hidden GUI) when invoked with `--silent`.

- **Drag-and-Drop Support**  
  - If `tkinterdnd2` is installed, drag a `.pdf`, `.xlsx`, or `.csv` onto the main window to open it immediately.

- **Auto-Versioned File Saving**  
  - By default, outputs a text file named `<source_filename>_extracted_links.txt` in the same folder of the executable.  
  - If that file already exists, it automatically increments: `… (1).txt`, `… (2).txt`, etc.  
  - User can override location via a “Choose Save Location” dialog.

---

## Screenshots

**Main Menu**  
![Main Menu](<docs/screenshots/Main Menu.png>)

**Dropping a Document File**  
![Dropping a Doc File](<docs/screenshots/Dropping a doc file.png>)

**Example Result**  
![Example Result](<docs/screenshots/Example result.png>)

**Highlighting Table Regions**  
![Highlighting Table Regions](<docs/screenshots/Highlighting table regions.png>)

---

## Dependencies

- **Python 3.10+**  
- **PyMuPDF** (`fitz`) — `pip install pymupdf`  
- **openpyxl** — `pip install openpyxl`  
- **Pillow** — `pip install pillow`  
- **pystray** — `pip install pystray`  
- **tkinterdnd2** (optional, for drag-and-drop) — `pip install tkinterdnd2`

> Note: `tkinter` is included in most Python installations. If using a distribution without Tk support, install the corresponding OS package (e.g., `python3-tk` on Debian/Ubuntu).

---

## Installation

1. **Clone or download** this repository to your local machine.  
2. **Install dependencies**:
   ```bash
   pip install pymupdf openpyxl pillow pystray tkinterdnd2
   ```
3. **Build the executable (Windows)**:
   ```bash
   pip install pyinstaller
   pyinstaller --onefile --noconsole pdf_spreadsheet_link_extractor.py
   ```
   - The resulting `pdf_spreadsheet_link_extractor.exe` will appear in the `dist/` folder.
4. **Prepare the bundle**:
   - Copy `pdf_spreadsheet_link_extractor.exe` into a new folder (e.g., `Link_Extractor_Bundle`).  
   - Create `add_to_startup.bat` and `README.md` (see “Project Structure” below).  
   - Compress the folder into `Link_Extractor_Bundle.zip` for distribution.

---

## Usage

### Standard GUI Mode
1. **Run the executable** (double-click `pdf_spreadsheet_link_extractor.exe`).  
2. The main window appears, showing:
   - “Browse File” button
   - “Choose Save Location” button
   - A scrollable text area for output
   - (If installed) drag-and-drop instructions at the top.

3. **Open any of**:
   - PDF (`.pdf`)
   - Excel workbook (`.xlsx`)
   - CSV (`.csv`)

4. **PDF**:  
   - A second window appears, showing the first page as an image.
   - Draw red rectangles around each table you want to extract links from.
   - Click **Done – Extract Links**.
   - In the main window, all extracted URLs appear.
   - Choose “Yes” to save results (auto-versioned) or “No” to skip.

5. **Excel / CSV**:  
   - Extracts all detected URLs and shows them in the output area.
   - Choose “Yes” to save, or “No” to skip.

6. **Close** the main window → hides to tray (remains running).

### Silent Tray Mode
To start the app minimized to the system tray (no window on launch), create a shortcut or `add_to_startup.bat` with the `--silent` flag:

```
pdf_spreadsheet_link_extractor.exe --silent
```

- No GUI appears on startup, only a tray icon.  
- Right-click **Link Extractor** tray icon:
  - **Open Link Extractor** → shows main window.  
  - **Quit** → exits the app completely.

---

## Project Structure

```
Link_Extractor_Bundle/
├── pdf_spreadsheet_link_extractor.exe    # Compiled executable (Windows)
├── add_to_startup.bat                    # Add app to Windows startup
├── README.md                             # Usage instructions
```

**Original source code** (for reference and rebuilding as `.exe`):

```
pdf_spreadsheet_link_extractor.py
```

---

## Configuration & Saving

- **Default save directory**: same folder as the executable.  
- **Custom save**: click “Choose Save Location” to pick any folder and filename.  
- **Auto-versioning**: if `filename.txt` exists, output becomes `filename (1).txt`, etc.

---

## Known Limitations

- **Single-page PDFs**: currently only the first page is scanned for links; multi-page support is not implemented (no pagination or scrolling within PDF selection window).  
- **Table detection**: entirely manual—requires drawing bounding boxes. No automatic table detection.  
- **Drag-and-drop**: only works on Windows when `tkinterdnd2` is installed.  
- **Tray icon graphics**: uses a programmatically drawn icon; you can replace it with a custom `.ico` if you prefer.

---

## Future Improvements

1. **Multi-page PDF Support**  
   - Loop through all pages and allow table selection on each page.  
   - Provide a “Next Page” and “Previous Page” button in the selection window.

2. **Automatic Table (Region) Detection**  
   - Integrate a lightweight OCR or image-analysis approach (e.g., edge detection) to auto-detect table boundaries.  
   - Offer a “Suggest Regions” button that pre-highlights possible table areas.

3. **Clickable Links in GUI Output**  
   - Display extracted URLs as clickable hyperlinks using a custom widget (e.g., a `tkinter.Text` with tag bindings).  
   - Enable “Open in Browser” on click.

4. **Progress & Status Bar**  
   - Show a progress indicator for large Excel/CSV files (e.g., “Reading row 3500 of 12,000”).  
   - Display an overall “Extracting…” splash screen on very large documents.

5. **Cross-Platform Packaging**  
   - Provide a macOS `.app` and a Linux AppImage or `.deb` with appropriate tray integration.  
   - Use libraries like `pyinstaller` + `electron-builder` or `cx_Freeze` for native bundling.

6. **Automated Scheduling / Watch Mode**  
   - Monitor a “Drop Zone” folder—whenever a PDF/CSV/XLSX is added, automatically extract links and save to a paired `.txt`.  
   - Useful for recurring ETL pipelines where new report files arrive daily.
---

## License

This project is released under the **MIT License**. Feel free to modify, redistribute, or incorporate it into your own workflows.
