param(
    [Parameter(Mandatory=$true)][string]$SrcDir,        # Folder with .bas/.cls/.frm source files
    [Parameter(Mandatory=$true)][string]$TemplateXlam, # Minimal template .xlam (contains references, ThisWorkbook, etc.)
    [Parameter(Mandatory=$true)][string]$OutXlam,      # Output add-in path
    [string]$CustomUIXml = ""                          # Optional: path to customUI.xml
)

# --- Guardrails ---
if (!(Test-Path $SrcDir)) { throw "SrcDir not found: $SrcDir" }
if (!(Test-Path $TemplateXlam)) { throw "TemplateXlam not found: $TemplateXlam" }
$srcFiles = Get-ChildItem -Path $SrcDir -File | Where-Object { $_.Extension -in ".bas",".cls",".frm" }
if ($srcFiles.Count -eq 0) { throw "No .bas/.cls/.frm files in $SrcDir" }

# --- Start Excel (requires: Excel installed + Trust access to VBOM enabled) ---
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open((Resolve-Path $TemplateXlam).Path)

    # Remove existing code modules (keep ThisWorkbook + any sheet modules)
    $vbproj = $wb.VBProject
    $toRemove = @()
    foreach ($comp in $vbproj.VBComponents) {
        # vbext_ct_Document = 100, standard module=1, class=2, form=3
        if ($comp.Type -ne 100) {
            $toRemove += $comp
        }
    }
    foreach ($comp in $toRemove) {
        $vbproj.VBComponents.Remove($comp)
    }

    # Import new modules
    foreach ($f in $srcFiles) {
        [void]$vbproj.VBComponents.Import((Resolve-Path $f.FullName).Path)
    }

    # Save as .xlam (FileFormat 55)
    $wb.SaveAs((Resolve-Path (Split-Path -Parent $OutXlam) | Out-String).Trim() + "\" + (Split-Path -Leaf $OutXlam), 55)
    $wb.Close($true)
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# Optional: inject customUI.xml into the .xlam (zip) after save
if ($CustomUIXml -and (Test-Path $CustomUIXml)) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $outPath = (Resolve-Path $OutXlam).Path
    $tempPath = [System.IO.Path]::GetTempFileName()
    Remove-Item $tempPath
    Copy-Item $outPath $tempPath

    $fs = [System.IO.File]::Open($tempPath, 'Open', 'ReadWrite', 'None')
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update)

        # Ensure folder entry exists and replace customUI/customUI.xml
        $entry = $zip.GetEntry("customUI/customUI.xml")
        if ($entry) { $entry.Delete() }
        # Create directory structure if not present (ZipArchive does it implicitly by entry name)
        $newEntry = $zip.CreateEntry("customUI/customUI.xml")
        $writer = New-Object System.IO.StreamWriter($newEntry.Open())
        try {
            $xml = Get-Content -Raw -Path $CustomUIXml
            $writer.Write($xml)
        } finally {
            $writer.Dispose()
        }

        # Update [Content_Types].xml with correct Override if missing
        $ct = $zip.GetEntry("[Content_Types].xml")
        if ($ct) {
            $ctStream = $ct.Open()
            $sr = New-Object System.IO.StreamReader($ctStream)
            $content = $sr.ReadToEnd()
            $sr.Dispose()
            $ctStream.Dispose()

            [xml]$ctXml = $content
            $ns = @{ "ct" = "http://schemas.openxmlformats.org/package/2006/content-types" }
            $overrideExists = $ctXml.Types.Override | Where-Object { $_.PartName -eq "/customUI/customUI.xml" }

            if (-not $overrideExists) {
                $newOverride = $ctXml.CreateElement("Override", $ns.ct)
                $newOverride.SetAttribute("PartName", "/customUI/customUI.xml")
                $newOverride.SetAttribute("ContentType", "application/vnd.ms-office.customUI+xml")
                $ctXml.Types.AppendChild($newOverride) | Out-Null
            }

            # Save back
            $ctTemp = [System.IO.Path]::GetTempFileName()
            $ctXml.Save($ctTemp)

            # Replace entry
            $ct.Delete()
            $ctNew = $zip.CreateEntry("[Content_Types].xml")
            $w = New-Object System.IO.StreamWriter($ctNew.Open())
            try { $w.Write((Get-Content -Raw $ctTemp)) } finally { $w.Dispose() }
            Remove-Item $ctTemp -Force
        }
        $zip.Dispose()
    } finally {
        $fs.Dispose()
    }

    # Replace original with modified
    Move-Item -Force $tempPath $outPath
}

Write-Host "Built: $OutXlam"
