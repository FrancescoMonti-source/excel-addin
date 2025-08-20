---
editor_options: 
  markdown: 
    wrap: 80
---

# Excel Add-in Development Guide

## 📂 Repository Layout 

-   excel-addin/

    -   src/ \# VBA source modules (.bas, .cls, .frm)

    -   addin_template.xlam \# Master add-in template (contains Ribbon +
        references)

    -   customUI.xml \# Ribbon definition (for version control)

    -   build.ps1 \# Build script (rebuilds .xlam from sources)

    -   dist/ \# Compiled add-in(s)

**addin_template.xlam:** \
The “master” file, edited in Office RibbonX Editor. Always contains the real
Ribbon (customUI/customUI.xml inside the zip).

**customUI.xml:** \
A plain-text copy of the ribbon XML stored in Git. Keep this in sync with the
one inside the template.

Workflow: edit Ribbon in RibbonX → export/copy to customUI.xml → commit. → If
you edit customUI.xml directly, re-import into addin_template.xlam.

**src/:** \
All VBA source modules under version control.

**dist/:** \
Build output folder. Ignore contents in Git.

## ⚙️ Build Process

1.  Open Office ribbonX Editor –\> open addin_template.xlam

2.  Insert Office 2010 CustomUI part –\> paste the code inside customUI.xml –\>
    save

3.  Run in PowerShell: `/build.ps1`

Steps performed by the script:

-   Open addin_template.xlam in headless Excel.

-   Delete all non-document VBA components.

-   Import fresh .bas, .cls, .frm files from src/.

-   Save as dist\my\_addin.xlam.

## 🎨 Ribbon Workflow

Ribbon XML is not injected by script — it lives inside addin_template.xlam.

Always keep customUI.xml in Git as the textual source of truth.

To update:

-   Open addin_template.xlam in Office RibbonX Editor.

-   Import or paste updated XML from customUI.xml.

-   Save the template.

-   Re-run the build script.

## 🚀 Tips

-   **Close Excel before building.** It locks files and can block saves.

<!-- -->

-    **Unblock the add‑in (once):** Right‑click `dist\my_addin.xlam` →
    **Properties** → if you see **Unblock**, tick it → OK.

-   **Trusted Location:** Excel → File → Options → **Trust Center** → Trust
    Center Settings → **Trusted Locations** → **Add new location…** → point to
    your repo’s `dist\` folder (tick **Subfolders** if needed).

<!-- -->

-    **Macro settings:** Trust Center → **Macro Settings** → “**Disable VBA
    macros with notification**” (or enable if you sign).\
    Also tick **Trust access to the VBA project object model** (required for the
    build to import modules).

<!-- -->

-    **Show UI errors:** File → Options → **Advanced** → General → tick **Show
    add‑in user interface errors** (helps diagnose Ribbon XML issues).

<!-- -->

-    **Enable Developer tab:** File → Options → **Customize Ribbon** → tick
    **Developer** (for VBE access & debugging).

<!-- -->

-    **Load the right file:** File → Options → Add‑ins → **Manage: Excel Add‑ins
    → Go…** → **Browse…** to `dist\my_addin.xlam`, tick it. Remove any stale
    entries (especially ones under `%APPDATA%\Microsoft\AddIns`).

<!-- -->

-    **Compile check:** VBE (`Alt+F11`) → **Debug → Compile VBAProject**. If the
    menu greys out with no errors, you’re good.

-    **References check (if code uses external libs):** VBE → **Tools →
    References…** → look for **MISSING** entries and fix paths.

<!-- -->

-    **If Ribbon doesn’t appear:**

    -   Confirm `addin_template.xlam` actually contains a single **customUI**
        part (use Office RibbonX Editor).

    -   If it’s corrupted, delete the `customUI` branch in RibbonX Editor and
        re‑create it (paste your `customUI.xml`) → Save.

    -   Rebuild and re‑enable the add‑in.

    -    **Optional (auto‑load at startup):** You can copy the built `.xlam` to
        `%APPDATA%\Microsoft\Excel\XLSTART` to load automatically for your user
        profile (handy for daily use—keep Git copy in `dist\` too).
