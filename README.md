# Resume Extractor - React UI

A complete web application for extracting resumes from Excel files with a beautiful React interface.

## Features

✅ Drag-and-drop file upload  
✅ Auto-detects resume column in Excel files  
✅ Processes multiple files automatically  
✅ Real-time progress tracking  
✅ Beautiful, responsive UI  
✅ Organized folder structure for results  

## Setup

All dependencies are already installed. Just run:

### 1. Start the Flask Backend
```powershell
c:/Users/DELL/Desktop/Excel/.venv/Scripts/python.exe app.py
```

The server will run on `http://localhost:5000`

### 2. Open the UI
Open [index.html](index.html) in your web browser or:
```
File > Open File > index.html
```

Then open the file in your default browser.

## How to Use

1. **Upload Excel Files**
   - Click the upload area or drag-and-drop Excel files
   - You can select multiple files at once

2. **Extract Resumes**
   - Click "Upload & Extract" button
   - Watch the real-time progress

3. **View Results**
   - Results are organized in folders by filename
   - Files are copied to `extracted_resumes/` folder

## File Structure

```
Excel/
├── app.py                    # Flask backend API
├── extract_resumes.py        # Original extraction script
├── index.html                # React UI
├── extracted_resumes/        # Output folder
│   ├── your_file1/
│   │   ├── resume1.pdf
│   │   ├── resume2.pdf
│   └── your_file2/
│       ├── resume3.pdf
└── .venv/                    # Python virtual environment
```

## Configuration

Edit [app.py](app.py) to customize:
- `BASE_DESTINATION`: Output folder name
- Port number (change `5000` to another port if needed)

## Notes

- The Resume Link column is auto-detected in your Excel files
- File paths in Excel should be absolute paths (C:\path\to\resume.pdf)
- All metadata (dates, permissions) is preserved during copy
