# POMtools

## Overview
POMtools is a PyRevit extension developed at POMERLEAU INC. for exporting Revit views to IFC and schedules to JSON. It simplifies and automates exports from Revit, especially for batch processing multiple views and files.

## Key Features
- Export views from active Revit document to IFC
- Export views from multiple external Revit files to IFC
- Export views to JSON format 
- Export schedules to JSON format
- Support for both IFC 2x3 Coordination View 2.0 and IFC4 Reference View
- Custom export configurations via JSON files

## Supported Export Formats
- **IFC 2x3 - Coordination View 2.0** (industry standard format)
- **IFC 4 - Reference View** (newer format with better geometry support)
- **Custom Configuration** (using JSON configuration files)
- **JSON** (for views and schedule data)

## Requirements
- Revit 2020 or higher
- pyRevit (latest version recommended)
- Admin rights for installation

## Usage

### IFC Export (Active Document)
1. Open the Revit file, click "Export Views to IFC (Document Actif)"
2. Select export format, destination folder, and views
3. IFC files will be created in a subfolder named after the Revit file

### IFC Export (External Files)
1. Click "Export Views to IFC (Autres Fichiers)"
2. Select export format, destination folder, and Revit files
3. For each file, select views to export
4. Files will be processed sequentially

### JSON Export (Views and Schedules)
1. Open the Revit file, click "Export Views to JSON" or "Export Schedules to JSON"
2. Select views or schedules and destination folder
3. JSON files will be created in the selected folder

## "Active View Only" Option
This option significantly reduces exported file size by including only elements visible in the selected view and excluding linked files and 2D elements.

## Troubleshooting
- **"Document is a linked file" error**: Open files directly and use the "Document Actif" script
- **Export errors**: Check for corrupt elements and folder access rights
- **Large files**: Use the "Active View Only" option

## Author
Yazid Ben Said | POMERLEAU INC.
