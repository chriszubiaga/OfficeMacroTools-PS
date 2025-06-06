param (
    [Parameter(Mandatory=$true, HelpMessage="Path to the Office file (e.g., .xlsm, .docm, .pptm)")]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$FilePath,

    [Parameter(Mandatory=$true, HelpMessage="Name of the VBA module to remove (e.g., Module1)")]
    [string]$ModuleNameToRemove
)

# Determine Office Application and settings based on file extension
$fileExtension = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
$appComObjectString = $null
$documentsOrPresentationsCollection = $null

switch ($fileExtension) {
    ".xlsm" {
        $appComObjectString = "Excel.Application"
        $documentsOrPresentationsCollection = "Workbooks"
    }
    ".xls" {
        $appComObjectString = "Excel.Application"
        $documentsOrPresentationsCollection = "Workbooks"
    }
    ".docm" {
        $appComObjectString = "Word.Application"
        $documentsOrPresentationsCollection = "Documents"
    }
    ".dotm" {
        $appComObjectString = "Word.Application"
        $documentsOrPresentationsCollection = "Documents"
    }
    ".pptm" {
        $appComObjectString = "PowerPoint.Application"
        $documentsOrPresentationsCollection = "Presentations"
    }
    ".ppsm" {
        $appComObjectString = "PowerPoint.Application"
        $documentsOrPresentationsCollection = "Presentations"
    }
    default {
        Write-Host "ERROR: Unsupported file type: $fileExtension. This script supports .xlsm, .xls, .docm, .dotm, .pptm, .ppsm for module removal." -ForegroundColor Red
        exit 1
    }
}

# --- Important: Office Trust Center Setting & File State ---
Write-Host "[*] Ensure 'Trust access to the VBA project object model' is ENABLED in the Trust Center settings of $($appComObjectString.Split('.')[0])." -ForegroundColor Yellow
Write-Host "[*] Ensure the target file '$FilePath' is NOT open in the Office application before running this script." -ForegroundColor Yellow

# Attempt to check for an exclusive file lock before proceeding
try {
    $testStream = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $testStream.Close()
    $testStream.Dispose()
}
catch [System.IO.IOException] {
    Write-Host "[!] ERROR: File '$FilePath' is currently in use or locked." -ForegroundColor Red
    Write-Host "[*] Please close the file in the Office application and ensure no other process is locking it, then try running the script again." -ForegroundColor Yellow
    exit 1
}

# Create the specific Office Application COM object
$appObject = New-Object -ComObject $appComObjectString
$appObject.Visible = $false
if ($appComObjectString -eq "Word.Application") {
    $appObject.DisplayAlerts = 0 # wdAlertsNone for Word
} else {
    $appObject.DisplayAlerts = $false # Suppress prompts for other applications
}

# Variables for COM cleanup
$documentObject = $null
$vbaProject = $null
$vbComponent = $null
$moduleRemoved = $false # Flag to track if a removal and save are necessary

try {
    Write-Host "[*] Attempting to open '$FilePath' with '$($appComObjectString.Split('.')[0])' for editing..." -ForegroundColor Cyan
    if ($appComObjectString -eq "PowerPoint.Application") {
        # PowerPoint Open parameters: Open(FileName, ReadOnly, Untitled, WithWindow)
        # For editing: ReadOnly=0 (msoFalse), Untitled=0 (msoFalse), WithWindow=0 (msoFalse for no UI)
        $documentObject = $appObject.($documentsOrPresentationsCollection).Open($FilePath, 0, 0, 0)
    } elseif ($appComObjectString -eq "Word.Application") {
        # Word Open parameters: Open(FileName, ConfirmConversions, ReadOnly, ...)
        # For editing: ConfirmConversions=$false, ReadOnly=$false
        $documentObject = $appObject.($documentsOrPresentationsCollection).Open($FilePath, $false, $false)
    } else { # Excel
        # Excel Open parameters: Open(FileName, UpdateLinks, ReadOnly)
        # For editing: UpdateLinks=0 (don't update external links), ReadOnly=$false
        $documentObject = $appObject.($documentsOrPresentationsCollection).Open($FilePath, 0, $false)
    }

    if (-not $documentObject) {
        Write-Host "ERROR: Failed to open the file for editing: '$FilePath'" -ForegroundColor Red
        throw "FileOpenFailed"
    }
    Write-Host "[*] Successfully opened: '$($documentObject.Name)'" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------"

    $hasVBProject = $false
    if ($appComObjectString -eq "Excel.Application") {
        $hasVBProject = $documentObject.HasVBProject
    } else {
        # Word/PowerPoint don't have a direct 'HasVBProject' property.
        # Attempt to access VBProject and assume true if successful, otherwise catch will handle.
        try {
            if ($documentObject.VBProject) { $hasVBProject = $true }
        } catch {
             # $hasVBProject remains false if VBProject access fails
        }
    }

    if ($hasVBProject) {
        $vbaProject = $documentObject.VBProject
        Write-Host "[*] VBA Project Name: '$($vbaProject.Name)'" -ForegroundColor DarkCyan

        try {
            $vbComponent = $vbaProject.VBComponents.Item($ModuleNameToRemove)
        }
        catch { # Handles cases where the module name does not exist.
            $vbComponent = $null
        }

        if ($vbComponent) {
            $componentNameFound = $vbComponent.Name
            $componentTypeFound = $vbComponent.Type
            Write-Host "[*] Found module: '${componentNameFound}'" -ForegroundColor Magenta
            Write-Host "[*] Component Type: $($componentTypeFound.ToString()) (Raw Value: $componentTypeFound)" -ForegroundColor Blue

            # vbext_ct_Document (Type 100) components like 'ThisWorkbook', 'Sheet1', 'ThisDocument'
            # cannot be removed directly. Their code can be cleared instead if needed.
            if ($componentTypeFound -eq 100) { # vbext_ct_Document
                 Write-Host "[*] Module '${componentNameFound}' is a Document-type component and cannot be removed directly. You can clear its code using a different script/logic." -ForegroundColor Yellow
            } else {
                Write-Host "[*] Attempting to remove module '${componentNameFound}'..." -ForegroundColor DarkCyan
                $vbaProject.VBComponents.Remove($vbComponent)
                Write-Host "[=] Module '${componentNameFound}' successfully removed from the VBA project." -ForegroundColor Green
                $moduleRemoved = $true
            }
        } else {
            Write-Host "[*] Module '$ModuleNameToRemove' not found in the VBA project of '$($documentObject.Name)'." -ForegroundColor Yellow
        }
         Write-Host "--------------------------------------------------"
    } else {
        Write-Host "[*] No VBA project found or accessible in '$($documentObject.Name)'." -ForegroundColor Yellow
        Write-Host "--------------------------------------------------"
    }

    if ($moduleRemoved) {
        Write-Host "[*] Saving changes to '$($documentObject.Name)'..." -ForegroundColor Cyan
        $documentObject.Save()
        Write-Host "[=] File saved successfully." -ForegroundColor Green
    } elseif (-not $hasVBProject) {
        Write-Host "[*] No action taken as no VBA project was found or accessible." -ForegroundColor Cyan
    } elseif (-not $vbComponent) { # Module was not found, but project was accessible
        Write-Host "[*] No action taken as module '$ModuleNameToRemove' was not found." -ForegroundColor Cyan
    } else { # Module was found but was a document type
        Write-Host "[*] No removal action taken for document component '$($vbComponent.Name)'. No save initiated by this script for this reason." -ForegroundColor Cyan
    }
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    if ($_.Exception.Message -like "*file is open*" -or $_.Exception.Message -like "*0x800A03EC*") {
         Write-Host "[*] This error can occur if the file is already open in the Office application, is corrupt, or locked by another process." -ForegroundColor Yellow
    }
}
finally {
    # Release COM objects. Order matters: specific objects before general ones.
    if ($null -ne $vbComponent) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbComponent) | Out-Null
        Remove-Variable vbComponent -ErrorAction SilentlyContinue
    }
    if ($null -ne $vbaProject) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbaProject) | Out-Null
        Remove-Variable vbaProject -ErrorAction SilentlyContinue
    }
    if ($documentObject) {
        try {
            # Document was explicitly saved if a module was removed.
            # Close without prompting for save or saving again.
            if ($appComObjectString -eq "Word.Application") {
                $documentObject.Close([ref]0) # wdDoNotSaveChanges = 0
            } elseif ($appComObjectString -eq "PowerPoint.Application") {
                $documentObject.Close()
            } else { # Excel
                $documentObject.Close($false) # SaveChanges:=$false
            }
        } catch {
            Write-Host "[*] Could not gracefully close the document object. Error: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($documentObject) | Out-Null
        Remove-Variable documentObject -ErrorAction SilentlyContinue
    }

    if ($appObject) {
        try {
            $appObject.Quit()
        } catch {
            Write-Host "[*] Could not gracefully quit the Office application. Error: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($appObject) | Out-Null
        Remove-Variable appObject -ErrorAction SilentlyContinue
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}