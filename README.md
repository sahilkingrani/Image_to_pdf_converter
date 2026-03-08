# Image + Word to PDF Converter

A simple desktop GUI application built with Tkinter that converts images and Word documents to PDF.  
Features drag-and-drop support and a clean interface.

## Features

- Drag & drop files directly into the app
- Select multiple files via a file dialog
- Supported input formats:
  - Images: `.png`, `.jpg`, `.jpeg`
  - Documents: `.docx`
- Image conversion:
  - Combines all selected images into a single PDF
  - Automatically scales images to fit an A4-like page (612x792 points)
- Word conversion:
  - Converts each `.docx` file into its own PDF
  - Opens the output folder after a successful conversion
- Status messages and error handling via message boxes

## Requirements

- Python 3.x
- Windows OS (recommended, especially for Word → PDF)
- Microsoft Word installed (required by `docx2pdf` on Windows)

Python libraries (installed via `requirements.txt`):

- `tkinterdnd2`
- `reportlab`
- `docx2pdf`
- `Pillow`


## Installation

1. Clone or download this project into a folder, for example:

   ```text
   c:\Users\sahil\OneDrive\Desktop\semester 5\python\python_project\image_to_pdf


- tkinterdnd2 – drag & drop support for Tkinter
- reportlab – creates the PDF from images
- Pillow – image loading/resizing ( from PIL import Image )
- docx2pdf – converts .docx files to PDF via Word
