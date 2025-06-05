# Define parameters at the top of the script
param (
    [Parameter(Mandatory=$true, HelpMessage="Path to the Office file (e.g., .xlsm, .docm, .pptm)")]
    [ValidateScript({Test-Path $_ -PathType Leaf})] # Ensures file exists
    [string]$FilePath,

    [Parameter(Mandatory=$true, HelpMessage="Name of the VBA module to remove (e.g., Module1)")]
    [string]$ModuleNameToRemove
)

# Determine Office Application and settings based on file extension
$fileExtension = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
$appComObjectString = $null
$openMethodName = $null
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
        Write-Error "Unsupported file type: $fileExtension. This script supports .xlsm, .xls, .docm, .dotm, .pptm, .ppsm for module removal."
        exit 1
    }
}

# --- Important: Office Trust Center Setting ---
Write-Warning "Ensure 'Trust access to the VBA project object model' is ENABLED in the Trust Center settings of $($appComObjectString.Split('.')[0])."
Write-Warning "Ensure the target file '$FilePath' is NOT open in the Office application."

# Create the specific Office Application COM object
$appObject = New-Object -ComObject $appComObjectString
$appObject.Visible = $false
$appObject.DisplayAlerts = $false # Suppress prompts; we'll handle saving.

# For Word, to be absolutely sure no dialogs appear during programmatic save/close
if ($appComObjectString -eq "Word.Application") {
    # WdAlertLevel constants: wdAlertsNone = 0, wdAlertsMessageBox = -2, wdAlertsAll = -1
    # $appObject.DisplayAlerts = 0 # wdAlertsNone (already set $appObject.DisplayAlerts generally)
}

# Variables for COM cleanup
$documentObject = $null
$vbaProject = $null
$vbComponent = $null
$moduleRemoved = $false

try {
    # Open the file for editing (not ReadOnly)
    Write-Host "Attempting to open $FilePath with $($appComObjectString.Split('.')[0]) for editing..."
    if ($appComObjectString -eq "PowerPoint.Application") {
        # PowerPoint Open method: Open(FileName, [ReadOnly As MsoTriState = msoFalse], [Untitled As MsoTriState = msoFalse], [WithWindow As MsoTriState = msoTrue])
        # msoTrue = -1, msoFalse = 0. For editing, ReadOnly must be msoFalse.
        $documentObject = $appObject.($documentsOrPresentationsCollection).Open($FilePath, 0, 0, 0) # ReadOnly=msoFalse, Untitled=msoFalse, WithWindow=msoFalse
    } elseif ($appComObjectString -eq "Word.Application") {
        # Word Open method: Open(FileName, [ConfirmConversions], [ReadOnly], [AddToRecentFiles], [PasswordDocument], [PasswordTemplate], [Revert], [WritePasswordDocument], [WritePasswordTemplate], [Format])
        $documentObject = $appObject.($documentsOrPresentationsCollection).Open($FilePath, $false, $false) # ConfirmConversions=$false, ReadOnly=$false
    } else { # Excel
        # Excel Open method: Open(FileName, [UpdateLinks], [ReadOnly])
        $documentObject = $appObject.($documentsOrPresentationsCollection).Open($FilePath, 0, $false) # UpdateLinks=0 (don't update), ReadOnly=$false
    }

    if (-not $documentObject) {
        Write-Error "Failed to open the file for editing: $FilePath"
        throw "FileOpenFailed"
    }
    Write-Host "Successfully opened: $($documentObject.Name)"

    # Access the VBA project
    $hasVBProject = $false
    if ($appComObjectString -eq "Excel.Application") {
        $hasVBProject = $documentObject.HasVBProject
    } else {
        # For Word/PowerPoint, we attempt to access and catch if not present or use a try-get approach.
        try {
            if ($documentObject.VBProject) { $hasVBProject = $true }
        } catch {}
    }

    if ($hasVBProject) {
        $vbaProject = $documentObject.VBProject
        Write-Host "Accessed VBA Project: $($vbaProject.Name)"

        try {
            $vbComponent = $vbaProject.VBComponents.Item($ModuleNameToRemove)
        }
        catch {
            $vbComponent = $null # Module not found
        }

        if ($vbComponent) {
            # Cannot remove Document-type components like ThisWorkbook, Sheet1, ThisDocument
            if ($vbComponent.Type -eq 100) { # vbext_ct_Document
                 Write-Warning "Module '$($ModuleNameToRemove)' is a Document-type component and cannot be removed directly. You can clear its code instead."
            } else {
                Write-Host "Found module: $($vbComponent.Name) (Type: $($vbComponent.Type)). Attempting to remove..."
                $vbaProject.VBComponents.Remove($vbComponent)
                Write-Host "Module '$ModuleNameToRemove' removed from the VBA project."
                $moduleRemoved = $true
            }
        } else {
            Write-Warning "Module '$ModuleNameToRemove' not found in the VBA project of '$($documentObject.Name)'."
        }
    } else {
        Write-Warning "No VBA project found or accessible in '$($documentObject.Name)'."
    }

    # Save the changes if a module was removed (or if other modifications were intended)
    if ($moduleRemoved) {
        Write-Host "Saving changes to '$($documentObject.Name)'..."
        $documentObject.Save()
        Write-Host "File saved successfully."
    } elseif (-not $hasVBProject -or -not $vbComponent) {
        Write-Host "No changes made as module was not found or no VBA project present."
    } else {
        Write-Host "Module '$ModuleNameToRemove' is a document component; no removal attempted. No save initiated by this script."
    }

}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    if ($_.Exception.Message -like "*file is open*" -or $_.Exception.Message -like "*0x800A03EC*") {
         Write-Warning "This error can occur if the file is already open in the Office application, is corrupt, or locked."
    }
}
finally {
    # Release the specific VBComponent COM object if it was retrieved
    if ($vbComponent -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbComponent) | Out-Null
        Remove-Variable vbComponent -ErrorAction SilentlyContinue
    }
    # Release the VBAProject COM object if it was retrieved
    if ($vbaProject -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($vbaProject) | Out-Null
        Remove-Variable vbaProject -ErrorAction SilentlyContinue
    }

    # Close the document/workbook/presentation
    if ($documentObject) {
        try {
            # We already saved if changes were made. Close without saving further changes.
            if ($appComObjectString -eq "Word.Application") {
                 $documentObject.Close([ref]0) # wdDoNotSaveChanges = 0
            } elseif ($appComObjectString -eq "PowerPoint.Application") {
                 # PowerPoint's Close doesn't have a SaveChanges parameter.
                 # If $documentObject.Saved is false and we want to discard, this is it.
                 # If we wanted to force save on close: $documentObject.Save(); $documentObject.Close()
                 $documentObject.Close()
            } else { # Excel
                 $documentObject.Close($false) # SaveChanges:=$false
            }
        } catch {
            Write-Warning "Could not gracefully close the document object. Error: $($_.Exception.Message)"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($documentObject) | Out-Null
        Remove-Variable documentObject -ErrorAction SilentlyContinue
    }

    # Quit the Office Application
    if ($appObject) {
        try {
            $appObject.Quit()
        } catch {
            Write-Warning "Could not gracefully quit the Office application. Error: $($_.Exception.Message)"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($appObject) | Out-Null
        Remove-Variable appObject -ErrorAction SilentlyContinue
    }

    # Suggest garbage collection
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}