param(
  [Parameter(Mandatory=$true)][string]$SrcDir,
  [Parameter(Mandatory=$true)][string]$TemplateXlam,
  [Parameter(Mandatory=$true)][string]$OutXlam,
  [string]$CustomUIXml = ""
)

function Ensure-Folder($path) {
  if (-not (Test-Path $path)) { New-Item -ItemType Directory -Force -Path $path | Out-Null }
}

# Preflight
if (-not (Test-Path $SrcDir)) { throw "SrcDir not found: $SrcDir" }
$srcFiles = Get-ChildItem -Path $SrcDir -File | Where-Object { $_.Extension -in ".bas",".cls",".frm" }
if ($srcFiles.Count -eq 0) { throw "No .bas/.cls/.frm files in $SrcDir" }

# Ensure output folder exists
$OutDir = Split-Path -Parent $OutXlam
Ensure-Folder $OutDir

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  # Create template if missing
  if (-not (Test-Path $TemplateXlam)) {
    Write-Host "Template not found; creating a blank add-in at $TemplateXlam"
    Ensure-Folder (Split-Path -Parent $TemplateXlam)
    $wbNew = $excel.Workbooks.Add()
    $wbNew.SaveAs($TemplateXlam, 55)  # 55 = xlOpenXMLAddIn (.xlam)
    $wbNew.Close($true)
  }

  # Open template
  $wb = $excel.Workbooks.Open($TemplateXlam)

  # Remove non-document components (keep ThisWorkbook / sheets)
  $vbproj = $wb.VBProject
  $toRemove = @()
  foreach ($comp in $vbproj.VBComponents) {
    if ($comp.Type -ne 100) { $toRemove += $comp }  # 100 = vbext_ct_Document
  }
  foreach ($comp in $toRemove) { $vbproj.VBComponents.Remove($comp) }

  # Import fresh modules
  foreach ($f in $srcFiles) { [void]$vbproj.VBComponents.Import($f.FullName) }

  # Save output directly (no Resolve-Path here)
  $wb.SaveAs($OutXlam, 55)
  $wb.Close($true)
}
finally {
  $excel.Quit()
  [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
}

# If the file doesn’t exist, stop here with an actionable error
if (-not (Test-Path $OutXlam)) {
  throw "Failed to build output add-in. Excel couldn't save to: $OutXlam"
}

# Inject customUI.xml (optional)
if ($CustomUIXml -and (Test-Path $CustomUIXml)) {
  Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

  $tempPath = [System.IO.Path]::GetTempFileName()
  Remove-Item $tempPath
  Copy-Item $OutXlam $tempPath

  $fs = [System.IO.File]::Open($tempPath, 'Open', 'ReadWrite', 'None')
  $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update)
  try {
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update)

    $entry = $zip.GetEntry("customUI/customUI.xml")
    if ($entry) { $entry.Delete() }
    $newEntry = $zip.CreateEntry("customUI/customUI.xml")
    $writer = New-Object System.IO.StreamWriter($newEntry.Open())
    try { $writer.Write((Get-Content -Raw -Path $CustomUIXml)) } finally { $writer.Dispose() }

    $ctEntry = $zip.GetEntry("[Content_Types].xml")
    if ($ctEntry) {
      $reader = New-Object System.IO.StreamReader($ctEntry.Open())
      $xml = [xml]$reader.ReadToEnd()
      $reader.Dispose()

      $hasOverride = $false
      foreach ($ov in $xml.Types.Override) {
        if ($ov.PartName -eq "/customUI/customUI.xml") { $hasOverride = $true; break }
      }
      if (-not $hasOverride) {
        $newOv = $xml.CreateElement("Override", "http://schemas.openxmlformats.org/package/2006/content-types")
        $newOv.SetAttribute("PartName", "/customUI/customUI.xml")
        $newOv.SetAttribute("ContentType", "application/vnd.ms-office.customUI+xml")

        [void]$xml.Types.AppendChild($newOv)

        # Replace content types part
        $ctEntry.Delete()
        $ctNew = $zip.CreateEntry("[Content_Types].xml")
        $w = New-Object System.IO.StreamWriter($ctNew.Open())
        try { $w.Write($xml.OuterXml) } finally { $w.Dispose() }
      }
    }

    $zip.Dispose()
  } finally {
    $fs.Dispose()
  }

  Move-Item -Force $tempPath $OutXlam
}

Write-Host "Built: $OutXlam"
