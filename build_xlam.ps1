param(
  [Parameter(Mandatory=$true)][string]$SrcDir,          # e.g. C:\...\excel-addin\src
  [Parameter(Mandatory=$true)][string]$TemplateXlam,    # e.g. C:\...\excel-addin\addin_template.xlam
  [Parameter(Mandatory=$true)][string]$OutXlam,         # e.g. C:\...\excel-addin\dist\my_addin.xlam
  [string]$CustomUIXml = ""                             # e.g. C:\...\excel-addin\customUI.xml (optional)
)

function Ensure-Folder($path) {
  if (-not (Test-Path $path)) { New-Item -ItemType Directory -Force -Path $path | Out-Null }
}

# ---------- Preflight ----------
if (-not (Test-Path $SrcDir)) { throw "SrcDir not found: $SrcDir" }
$srcFiles = Get-ChildItem -Path $SrcDir -File | Where-Object { $_.Extension -in ".bas",".cls",".frm" }
if ($srcFiles.Count -eq 0) { throw "No .bas/.cls/.frm files in $SrcDir" }

$OutDir = Split-Path -Parent $OutXlam
Ensure-Folder $OutDir
Ensure-Folder (Split-Path -Parent $TemplateXlam)

# ---------- Start Excel ----------
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  # Create template if missing
  if (-not (Test-Path $TemplateXlam)) {
    Write-Host "Template not found; creating blank add-in at $TemplateXlam"
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

  # Save output .xlam
  $wb.SaveAs($OutXlam, 55)
  $wb.Close($true)
}
finally {
  $excel.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}

# ---------- Stop here if build failed ----------
if (-not (Test-Path $OutXlam)) { throw "Failed to build output add-in. Excel couldn't save to: $OutXlam" }

# ---------- Inject customUI.xml (optional) + ensure ContentTypes + relationship ----------
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

$needsRibbonPatch = $false
if ($CustomUIXml -and (Test-Path $CustomUIXml)) { $needsRibbonPatch = $true }

# Work on temp copy, then swap back (avoids file locks)
$tmp = [IO.Path]::GetTempFileName()
Remove-Item $tmp
Copy-Item $OutXlam $tmp

$fs  = [IO.File]::Open($tmp, 'Open', 'ReadWrite', 'None')
$zip = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Update)

# 1) Write /customUI/customUI.xml (only if provided)
if ($needsRibbonPatch) {
  $entry = $zip.GetEntry("customUI/customUI.xml")
  if ($entry) { $entry.Delete() }
  $newEntry = $zip.CreateEntry("customUI/customUI.xml")
  $w = New-Object IO.StreamWriter($newEntry.Open())
  $w.Write([IO.File]::ReadAllText($CustomUIXml))
  $w.Dispose()
}

# 2) Ensure [Content_Types].xml override for customUI
$ct = $zip.GetEntry("[Content_Types].xml")
if (-not $ct) { throw "[Content_Types].xml not found inside the add-in." }
$reader = New-Object IO.StreamReader($ct.Open())
[xml]$ctXml = $reader.ReadToEnd()
$reader.Dispose()

$hasOverride = $false
foreach ($ov in $ctXml.Types.Override) {
  if ($ov.PartName -eq "/customUI/customUI.xml") { $hasOverride = $true; break }
}
if ($needsRibbonPatch -and -not $hasOverride) {
  $ov = $ctXml.CreateElement("Override", "http://schemas.openxmlformats.org/package/2006/content-types")
  $ov.SetAttribute("PartName", "/customUI/customUI.xml")
  $ov.SetAttribute("ContentType", "application/vnd.ms-office.customUI+xml")

  [void]$ctXml.Types.AppendChild($ov)

  # Replace content types part
  $ct.Delete()
  $ctNew = $zip.CreateEntry("[Content_Types].xml")
  $w2 = New-Object IO.StreamWriter($ctNew.Open())
  $w2.Write($ctXml.OuterXml)
  $w2.Dispose()
}

# 3) Ensure /_rels/.rels relationship to customUI/customUI.xml
if ($needsRibbonPatch) {
  $relsPath = "_rels/.rels"
  $rels = $zip.GetEntry($relsPath)
  if (-not $rels) {
    $rels = $zip.CreateEntry($relsPath)
    $w = New-Object IO.StreamWriter($rels.Open())
    $w.Write('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>')
    $w.Dispose()
  }

  $reader = New-Object IO.StreamReader($rels.Open())
  [xml]$relsXml = $reader.ReadToEnd()
  $reader.Dispose()

  $relType = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"

  $existing = $relsXml.Relationships.Relationship | Where-Object {
    $_.Type -eq $relType -and $_.Target -eq "customUI/customUI.xml"
  }

  if (-not $existing) {
    $nextId = 1
    foreach ($r in $relsXml.Relationships.Relationship) {
      if ($r.Id -match '^rId(\d+)$') {
        $n = [int]$Matches[1]; if ($n -ge $nextId) { $nextId = $n + 1 }
      }
    }
    $newRel = $relsXml.CreateElement("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")
    $newRel.SetAttribute("Id", "rId$nextId")
    $newRel.SetAttribute("Type", $relType)
    $newRel.SetAttribute("Target", "customUI/customUI.xml")

    [void]$relsXml.Relationships.AppendChild($newRel)

    # Replace rels entry
    $rels.Delete()
    $relsNew = $zip.CreateEntry($relsPath)
    $w3 = New-Object IO.StreamWriter($relsNew.Open())
    $w3.Write($relsXml.OuterXml)
    $w3.Dispose()
  }
}

$zip.Dispose(); $fs.Dispose()

# Swap temp back to output
Move-Item -Force $tmp $OutXlam

Write-Host "Built: $OutXlam"
if ($needsRibbonPatch) { Write-Host " - Injected Ribbon customUI and relationships." } else { Write-Host " - No customUI.xml provided; skipped Ribbon injection." }
