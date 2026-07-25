# Automated-Cataloguing-System-Task
This project is a Python-based tool that automatically processes a folder of PDF files and generates a structured Excel catalog (catalog.xlsx). It is designed to simplify digital library management by extracting metadata from PDF books and organizing it into a searchable catalog.

Author: Rashi Bongirwar 
Date: 28/08/2025  
Task: Automated Cataloguing System

-----------------------------------------------------
1. Description
-----------------------------------------------------
This project is a Python script that automatically processes a folder of PDF files 
and generates a structured Excel catalog (catalog.xlsx).  

The catalog contains the following columns for each book:  
- Book Title  
- Author  
- Editor  
- Year of Publishing  
- Publisher  
- Language  
- Number of Pages  
- Format  

If any information is missing or unavailable, the script fills it with "Unknown".

-----------------------------------------------------
2. Requirements
-----------------------------------------------------
Install the following dependencies before running:

    pip install PyMuPDF, PyPDF2, pandas, openpyxl, langdetect.
-----------------------------------------------------
3. Usage
-----------------------------------------------------
1. Place the script file (`python pdf_catalog_generator`) in any folder.  
2. Create a folder named `PDF_Books` in the same location as the script.  
3. Add all your PDF files into the `PDF_Books` folder.  
4. Run the script from terminal/command prompt:
    
  First step:
       cd Desktop
  
  Second step:
      # Now your working location is Desktop.
      # If it’s inside another folder, say:
  
    python python pdf_catalog_generator.py

  Third step:
      Enter the path to the directory containing PDF files:
     # just type:
      PDF_Books

5. When prompted, enter the folder path (`PDF_Books`).  
6. The output file `catalog.xlsx` will be generated in the same folder as the script.

-----------------------------------------------------
4. Output
-----------------------------------------------------
- catalog.xlsx → Excel file containing extracted book metadata.  
- Each row corresponds to one PDF file.  

-----------------------------------------------------
5. Notes
-----------------------------------------------------
- If the PDF does not contain metadata or text for certain fields, 
  "Unknown" will be written in the catalog.  
- The script works best with text-based PDFs. 
  For scanned PDFs, OCR integration (using pytesseract) may be added.  
- 'en' represent English in excel sheet
=====================================================
End of README
=====================================================
