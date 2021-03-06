Fix Excel 2013 (32-bits) bug of random crashes with below description.

  xlsm files are displaying "Microsoft Excel has stopped working." whenever being opened.
  It occurs randomly after file access by other users.

  Files are shared and accessed by different users in a shared drive.
  Error pops up whenever VBA codes are compiled. 
  - Pressing "Enable Macro" button
  - Inputting any shortcuts which require a Macro compilation (e.g. Ctrl+S, Ctrl+C...)

Feedback from Microsoft:

  "Files containing macros and that linked to other worksheets crashed Excel in vbe7.dll."
  https://answers.microsoft.com/en-us/msoffice/forum/msoffice_excel-mso_win10-mso_o365b/excel-64-bit-crashes-in-vbe7dll-when-saving-files/ae7d5132-988b-4300-aceb-8e9fb314bd54

Temporary Fix:

  1) Copy the target file. ("test.xlsm" -> "test - Copy.xlsm"). 
  2) Open the copied file without Enabling Macro.
  3) Go to vba editor (Alt+F11).
  4) Make changes without impacting the code. (e.g. adding a " " and delete it)
  5) Save it as original name. (Replacing "test.xlsm")
  
Files will then be recovered temporarily, but files may crash again a while later.

Permanent Fix: (any one of below)
1) Add registry by regedit.exe
2) Running the attached ForceVBALoadFromSource.reg
3) Running below command line.
Reg.exe add "HKCU\Software\Microsoft\Office\16.0\Excel\Options" /v "ForceVBALoadFromSource" /t REG_DWORD /d "1" /f


Varying by the Excel verions, registry dir may vary.

