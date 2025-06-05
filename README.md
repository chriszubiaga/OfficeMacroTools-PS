# Office VBA Management Scripts for PowerShell

This repository contains PowerShell scripts designed to help manage VBA (Visual Basic for Applications) projects within Microsoft Office files (Excel, Word, PowerPoint). You can inspect VBA macro code and remove specific VBA modules programmatically.

## Features

* **Inspect VBA Macros**: View the names, types, and code content of all VBA components within a specified Office file.
* **Remove VBA Modules**: Programmatically delete specific VBA modules (e.g., standard modules, class modules) from an Office file.
* **Cross-Application Support**: Works with:
    * Excel: `.xlsm`, `.xls`
    * Word: `.docm`, `.dotm`
    * PowerPoint: `.pptm`, `.ppsm`

## Scripts

1.  **`Get-OfficeMacroInfo.ps1`**: Reads and displays information about VBA components and their code from an Office file.
2.  **`Remove-OfficeVbaModule.ps1`**: Removes a specified VBA module from an Office file.

## Prerequisites

1.  **Windows PowerShell**: These scripts are designed for PowerShell (typically version 5.1 or later).
2.  **Microsoft Office Installed**: The relevant Microsoft Office application (Excel, Word, or PowerPoint) must be installed on the machine where the script is run. The scripts use COM automation to interact with the Office applications.
3.  **"Trust access to the VBA project object model"**: This setting **MUST BE ENABLED** in the Trust Center for each Office application you intend to use with these scripts.
    * To enable:
        1.  Open the Office application (e.g., Excel).
        2.  Go to `File > Options > Trust Center > Trust Center Settings...`.
        3.  Select `Macro Settings`.
        4.  Under "Developer Macro Settings", check the box for **"Trust access to the VBA project object model"**.
        5.  Click `OK` to close the dialogs.
        6.  Repeat for Word and PowerPoint if you plan to use the scripts with those file types.

## Usage

### 1. `Get-OfficeMacroInfo.ps1`

This script inspects an Office file and outputs details about its VBA project, including components and their code.

**Parameters:**

* `-FilePath <string>`: (Mandatory) The full path to the Office file you want to inspect.

**Example:**

```powershell
.\Get-OfficeMacroInfo.ps1 -FilePath "C:\Path\To\Your\Workbook.xlsm"
