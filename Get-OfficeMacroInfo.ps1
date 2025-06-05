param (
    [Parameter(Mandatory=$true, HelpMessage="Path to the Office file (e.g., .xlsm, .docm, .pptm)")]
    [string]$FilePath
)

# Determine Office Application and settings based on file extension
$fileExtension = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
$appComObjectString = $null
$openMethodName = $null
$documentsOrPresentationsCollection = $null

switch ($fileExtension) {
    ".xlsm" {
        $appComObjectString = "Excel.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Workbooks"
    }
    ".xls" {
        $appComObjectString = "Excel.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Workbooks"
    }
    ".docm" {
        $appComObjectString = "Word.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Documents"
    }
    ".dotm" {
        $appComObjectString = "Word.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Documents"
    }
    ".pptm" {
        $appComObjectString = "PowerPoint.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Presentations"
    }
    ".ppsm" {
        $appComObjectString = "PowerPoint.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Presentations"
    }
    default {
        Write-Error "Unsupported file type: $fileExtension. This script supports .xlsm, .xls, .docm, .dotm, .pptm, .ppsm for macro information retrieval."
        exit 1
    }
}

# --- Important: Office Trust Center Setting ---
Write-Warning "Ensure 'Trust access to the VBA project object model' is ENABLED in the Trust Center settings of the respective Office application ($($appComObjectString.Split('.')[0]))."

$appObject = New-Object -ComObject $appComObjectString
$appObject.Visible = $false
if ($appComObjectString -eq "Word.Application") {
    $appObject.DisplayAlerts = 0
} else {
    $appObject.DisplayAlerts = $false
}

$documentObject = $null

try {
    Write-Host "[*] Attempting to open '$FilePath' with '$($appComObjectString.Split('.')[0])'..." -ForegroundColor Cyan
    if ($appComObjectString -eq "PowerPoint.Application") {
        # PowerPoint Open parameters: Open(FileName, ReadOnly, Untitled, WithWindow)
        # -1 (msoTrue) for ReadOnly, 0 (msoFalse) for Untitled and WithWindow (no UI)
        $documentObject = $appObject.($documentsOrPresentationsCollection).$openMethodName($FilePath, -1, 0, 0)
    } else {
        # Excel Open parameters: Open(FileName, UpdateLinks, ReadOnly)
        # Word Open parameters: Open(FileName, ConfirmConversions, ReadOnly)
        # For both, $false for UpdateLinks/ConfirmConversions, $true for ReadOnly.
        $documentObject = $appObject.($documentsOrPresentationsCollection).$openMethodName($FilePath, $false, $true)
    }

    if (-not $documentObject) {
        Write-Error "Failed to open the file: $FilePath"
        throw "FileOpenFailed"
    }

    Write-Host "[*] Successfully opened: '$($documentObject.Name)'" -ForegroundColor Cyan
    Write-Host "[*] Reading macros from: '$($documentObject.Name)'" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------"

    $hasVBProjectProperty = $null
    if ($appComObjectString -eq "Excel.Application") {
        $hasVBProjectProperty = $documentObject.HasVBProject
    } else {
        # Word and PowerPoint don't have a direct 'HasVBProject' property.
        # Assume a project might exist and attempt to access VBProject, catching errors if it doesn't.
        $hasVBProjectProperty = $true
    }

    if ($hasVBProjectProperty) {
        $vbaProject = $null
        try {
            $vbaProject = $documentObject.VBProject
        } catch {
            Write-Warning "Could not access the VBA project. This can happen if:"
            Write-Warning "1. The file has no VBA macros."
            Write-Warning "2. 'Trust access to the VBA project object model' is NOT enabled in the respective Office application's Trust Center."
            Write-Warning "3. The file's VBA project is password protected."
        }

        if ($vbaProject) {
            Write-Host "[*] VBA Project Name: '$($vbaProject.Name)'" -ForegroundColor DarkCyan # Changed from Yellow
            Write-Host "[*] VBComponents found: $($vbaProject.VBComponents.Count)" -ForegroundColor DarkCyan # Changed from Yellow
            foreach ($component in $vbaProject.VBComponents) {
                $componentName = $component.Name
                $componentTypeString = try { $component.Type.ToString() } catch { "Unknown" }

                Write-Host "[*] Component Name: '${componentName}'" -ForegroundColor Magenta
                Write-Host "[*] Component Type: $componentTypeString (Raw Value: $($component.Type))" -ForegroundColor Blue # Changed from DarkYellow

                if ($component.CodeModule) {
                    $linesOfCode = $component.CodeModule.CountOfLines
                    if ($linesOfCode -gt 0) {
                        Write-Host "[=] Code in ${componentName}:" -ForegroundColor Red 
                        Write-Host "----- Start -----" -ForegroundColor Red
                        $macroCode = $component.CodeModule.Lines(1, $linesOfCode)
                        Write-Host $macroCode -ForegroundColor Red
                        Write-Host "------ End ------" -ForegroundColor Red
                        Write-Host "---------------------------" # This separator appears after the code block
                    } else {
                        Write-Host "[=] No code found in ${componentName}." -ForegroundColor Green
                        Write-Host "---------------------------"
                    }
                } else {
                    Write-Host "[*] No CodeModule for component ${componentName} (this might be expected for some component types)." -ForegroundColor DarkGray
                    Write-Host "---------------------------"
                }
            }
        } else {
            Write-Warning "No VBA project found or accessible in '$($documentObject.Name)'."
            Write-Warning "Ensure 'Trust access to the VBA project object model' is enabled in the $($appComObjectString.Split('.')[0]) Trust Center."
        }
    } else {
         Write-Warning "The file '$($documentObject.Name)' does not report having a VBA project (e.g., Excel's HasVBProject is false)."
    }
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    if ($_.Exception.Message -like "*0x800A03EC*") {
         Write-Warning "This error (often 0x800A03EC) can indicate that the file is corrupt, not a valid Office document for the specified application, or Excel/Word cannot access it."
    }
}
finally {
    if ($documentObject) {
        try {
            # Close the document without saving changes, as this script is read-only.
            if ($appComObjectString -eq "Word.Application") {
                # Word's Close method: Close([SaveChanges], [OriginalFormat], [RouteDocument])
                # [ref]0 corresponds to WdSaveOptions.wdDoNotSaveChanges
                $documentObject.Close([ref]0)
            } elseif ($appComObjectString -eq "PowerPoint.Application") {
                # PowerPoint's Close method does not take a SaveChanges parameter.
                # Relies on the .Saved property or explicit .Save() if changes were made (not in this script).
                $documentObject.Close()
            } else {
                # Excel's Close method: Close([SaveChanges])
                $documentObject.Close($false) # $false means do not save changes.
            }
        } catch {
            Write-Warning "Could not gracefully close the document object. Error: $($_.Exception.Message)"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($documentObject) | Out-Null
        Remove-Variable documentObject -ErrorAction SilentlyContinue
    }

    # Quit the Office application and release its COM object.
    if ($appObject) {
        try {
            $appObject.Quit()
        } catch {
            Write-Warning "Could not gracefully quit the Office application. Error: $($_.Exception.Message)"
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($appObject) | Out-Null
        Remove-Variable appObject -ErrorAction SilentlyContinue
    }
    Write-Host "--------------------------------------------------"
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}