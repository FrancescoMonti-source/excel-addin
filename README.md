# Run the following script in powershell

# Make sure the script is unblocked OR use Bypass
Unblock-File .\build_xlam.ps1

powershell -NoProfile -ExecutionPolicy Bypass -File .\build_xlam.ps1 `
  -SrcDir "C:\Users\franc\Documents\Git\excel-addin\src" `
  -TemplateXlam "C:\Users\franc\Documents\Git\excel-addin\addin_template.xlam" `
  -OutXlam "C:\Users\franc\Documents\Git\excel-addin\dist\my_addin.xlam" `
  -CustomUIXml "C:\Users\franc\Documents\Git\excel-addin\customUI.xml"
