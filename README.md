Fix Excel 2013 (32-bits) bug of random crashes with below description.

xlsm files are displaying "Microsoft Excel has stopped working." randomly
after file access by other users.

- Files are shared and accessed by different users in a shared drive.
- Error pops out whenever VBA codes are compiled. 
  - Pressing "Enable Macro" button
  - Inputting any shortcuts which require a compilation (e.g. Ctrl+S, Ctrl+C...)

Temporary fixes:
Copy the target file





Feedback from Microsoft:
"Files containing macros and that linked to other worksheets crashed Excel in vbe7.dll."
https://answers.microsoft.com/en-us/msoffice/forum/msoffice_excel-mso_win10-mso_o365b/excel-64-bit-crashes-in-vbe7dll-when-saving-files/ae7d5132-988b-4300-aceb-8e9fb314bd54
