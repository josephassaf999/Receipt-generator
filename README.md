#Excel to PDF Receipt Generator

This Python tool lets you generate receipt PDFs from Excel data using a Word template. It works on Windows and allows you to create individual PDFs per row or a merged PDF for all rows. You can pause, resume, or cancel the process anytime.

How to Use

Make sure Python 3.x is installed on your Windows machine.

Install the required packages by running:

pip install pandas openpyxl python-docx PyPDF2 pywin32


Prepare your Excel file with the data and a Word template with placeholders.

Placeholders must match the Excel column names exactly, surrounded by double curly braces.
Example: Excel column Name → Word placeholder {{Name}}

Run the script:

python receipt_generator.py


In the GUI:

Upload your Excel file.

Upload your Word template.

Select an output folder.

(Optional) Check “Remember Last Used Files/Folders”.

Click Start Generating Receipts.

Use Pause, Resume, or Cancel buttons as needed.

Generated PDFs will be saved in the output folder.

If the process is cancelled, PDFs generated so far are merged into All_Receipts.pdf.
