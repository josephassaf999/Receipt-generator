# Excel to PDF Receipt Generator

<img width="780" height="631" alt="image" src="https://github.com/user-attachments/assets/3d159bd5-1296-44de-a4a8-72219c443cff" />

This tool lets you generate receipt PDFs from Excel data using a Word template. It works on Windows and allows you to create individual PDFs per row or a merged PDF for all rows. You can pause, resume, or cancel the process anytime.

## How to Use

1. Extract the zip file to a folder on your computer.
2. Double-click `AutoReceiptPro.exe` to launch the program.
3. In the GUI:
   - Upload your Excel file.
   - Upload your Word template.
   - Select an output folder (the folder name should be one word).
   - (Optional) Check “Remember Last Used Files/Folders”.
4. Click **Start Generating Receipts**.
5. Use **Pause**, **Resume**, or **Cancel** buttons as needed.

Generated PDFs will be saved in the output folder. If the process is cancelled, PDFs generated so far are merged into `All_Receipts.pdf`.

## Notes
- Your Excel columns should match the placeholders in your Word template exactly.  
  Example: Excel column `Name` → Word placeholder `{{Name}}`
- No Python or special software is required—just run the `.exe`.
