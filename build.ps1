# build.ps1 — build the add-in from src, keeping Ribbon baked into addin_template.xlam
param(
  [string]$SrcDir       = "$PWD\src",
  [string]$TemplateXlam = "$PWD\addin_template.xlam",
  [string]$OutXlam      = "$PWD\dist\my_addin.xlam",
  [switch]$AutoLoad     # if set, copy to XLSTART for per-user auto-load
)

function Ensure-Folder($p){
  if(-not (Test-Path $p)){ New-Item -ItemType Directory -Force -Path $p | Out-Null }
}

# --- Preflight ---
if(-not (Test-Path $SrcDir)){ throw "SrcDir not found: $SrcDir" }
$srcFiles = Get-ChildItem -Path $SrcDir -File | Where-Object { $_.Extension -in ".bas",".cls",".frm" }
if(!$srcFiles){ throw "No .bas/.cls/.frm files in $SrcDir" }

$OutDir = Split-Path -Parent $OutXlam
Ensure-Folder $OutDir
Ensure-Folder (Split-Path -Parent $TemplateXlam)

# --- Build (Excel COM) ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  if (-not (Test-Path $TemplateXlam)) {
    Write-Host "Template not found; creating blank add-in at $TemplateXlam"
    $wbNew = $excel.Workbooks.Add()
    $wbNew.SaveAs([IO.Path]::GetFullPath($TemplateXlam), 55)  # 55 = .xlam
    $wbNew.Close($true)
    Write-Host "Reminder: open addin_template.xlam in Office RibbonX Editor and add your customUI once."
  }

  $wb = $excel.Workbooks.Open([IO.Path]::GetFullPath($TemplateXlam))

  # Strip non-document modules (keep ThisWorkbook/Sheet objects), then import fresh
  $vbproj = $wb.VBProject
  $toRemove = @()
  foreach($c in $vbproj.VBComponents){ if($c.Type -ne 100){ $toRemove += $c } }  # 100=document
  foreach($c in $toRemove){ $vbproj.VBComponents.Remove($c) }

  foreach($f in $srcFiles){ [void]$vbproj.VBComponents.Import($f.FullName) }

  # Clean, correct SaveAs path
  $outFull = [IO.Path]::GetFullPath((Join-Path $OutDir (Split-Path -Leaf $OutXlam)))
  $wb.SaveAs($outFull, 55)   # .xlam
  $wb.Close($true)
}
finally {
  $excel.Quit()
  # Release COM cleanly (avoid pipeline noise)
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
}

if(-not (Test-Path $OutXlam)){ throw "Failed to build output add-in: $OutXlam" }

# Unblock the output (no-op if not blocked)
Unblock-File -Path $OutXlam -ErrorAction SilentlyContinue

Write-Host ("Built: {0}  ({1:N0} bytes)" -f $OutXlam, (Get-Item $OutXlam).Length)

# Optional: copy to XLSTART for guaranteed auto-load
if($AutoLoad){
  $xlStart = Join-Path $env:APPDATA "Microsoft\Excel\XLSTART"
  Ensure-Folder $xlStart
  $target = Join-Path $xlStart (Split-Path -Leaf $OutXlam)
  Copy-Item -Force $OutXlam $target
  Write-Host "Installed for auto-load: $target"
}
