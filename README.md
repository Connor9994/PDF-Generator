# PDF-Generator

![GitHub stars](https://img.shields.io/github/stars/Connor9994/PDF-Generator?style=social) ![GitHub forks](https://img.shields.io/github/forks/Connor9994/PDF-Generator?style=social) ![GitHub issues](https://img.shields.io/github/issues/Connor9994/PDF-Generator)

![output](https://github.com/Connor9994/PDF-Generator/assets/39637206/610e6c8c-bb20-4ce2-8e4e-866baacf5d15)

## Features

- **Date Selection**: Interactive calendar interface for selecting payment dates
- **Excel Data Extraction**: Automatically reads patient names and credit amounts from active Excel workbooks
- **Automated PDF Download**: Logs into Zirmed portal and downloads payment PDFs based on extracted data
- **Batch Processing**: Processes multiple records in sequence
- **Error Reporting**: Displays any records where PDFs couldn't be found
- **Stealth Mode**: Optional browser visibility toggle

## Requirements

- Windows OS
- PowerShell
- Microsoft Excel
- Internet Explorer
- Zirmed portal account

## Usage

1. **Select Date**: Click "Select Date" to choose the payment date from the calendar
2. **Load Names**: Click "Load Names" to extract patient names and credit amounts from your active Excel workbook
3. **Download PDFs**: Click "Download PDFs" to automatically:
   - Log into Zirmed portal
   - Search for payments matching the extracted data
   - Download corresponding PDFs
   - Report any missing PDFs

### Excel Data Format
The script expects Excel data in this format:
- Column A: Dates
- Column C: Patient Names  
- Column E: Credit Amounts

### Authentication
The script includes hardcoded credentials for Zirmed portal access.
```powershell
$Username="your_username"
$Password="your_password"
```

## Interface

The application features a simple GUI with:
- Three main action buttons with status indicators
- Date selection calendar popup
- Browser visibility toggle
- Error reporting for missing PDFs

SEE "Instructions" Folder for PDF examples (Blurred due to the PHI-nature of the info)

![1](https://github.com/Connor9994/PDF-Generator/blob/main/Photos/1.png)
![2](https://github.com/Connor9994/PDF-Generator/blob/main/Photos/2.png)

## Technical Details

- Built with PowerShell Windows Forms
- Uses Excel COM automation
- Internet Explorer automation for web interaction
- Fixed dialog window (230x205 pixels)
- Custom icon and font styling
