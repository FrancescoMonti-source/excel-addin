# Run the following script in powershell. 
# Requires Windows + Excel installed + “Trust access to the VBA project 
# object model” enabled in Excel’s Trust Center.

# Make sure the script is unblocked OR use Bypass
Unblock-File .\build_xlam.ps1

powershell -NoProfile -ExecutionPolicy Bypass -File .\build_xlam.ps1 `
  -SrcDir "C:\Users\franc\Documents\Git\excel-addin\src" `
  -TemplateXlam "C:\Users\franc\Documents\Git\excel-addin\addin_template.xlam" `
  -OutXlam "C:\Users\franc\Documents\Git\excel-addin\dist\my_addin.xlam" `
  -CustomUIXml "C:\Users\franc\Documents\Git\excel-addin\customUI.xml"
