# Define parameters at the top of the script
param (
    [Parameter(Mandatory=$true, HelpMessage="Path to the Office file (e.g., .xlsm, .docm, .pptm)")]
    [ValidateScript({Test-Path $_ -PathType Leaf})] # Ensures file exists
    [string]$FilePath
)

# Determine Office Application and settings based on file extension
$fileExtension = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
$appComObjectString = $null
$openMethodName = $null
$documentProperty = $null # e.g., ActiveWorkbook, ActiveDocument, ActivePresentation (less reliable than direct object)
$documentsOrPresentationsCollection = $null

switch ($fileExtension) {
    ".xlsm" {
        $appComObjectString = "Excel.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Workbooks"
    }
    ".xls" { # Older Excel format, can contain macros (XLS)
        $appComObjectString = "Excel.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Workbooks"
    }
    ".docm" {
        $appComObjectString = "Word.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Documents"
    }
    ".dotm" { # Word Macro-Enabled Template
        $appComObjectString = "Word.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Documents"
    }
    ".pptm" {
        $appComObjectString = "PowerPoint.Application"
        $openMethodName = "Open" # PowerPoint's Open method needs slightly different params
        $documentsOrPresentationsCollection = "Presentations"
    }
    ".ppsm" { # PowerPoint Macro-Enabled Show
        $appComObjectString = "PowerPoint.Application"
        $openMethodName = "Open"
        $documentsOrPresentationsCollection = "Presentations"
    }
    default {
        Write-Error "Unsupported file type: $fileExtension. This script supports .xlsm, .xls, .docm, .dotm, .pptm, .ppsm."
        exit 1
    }
}

# --- Important: Office Trust Center Setting ---
Write-Warning "Ensure 'Trust access to the VBA project object model' is ENABLED in the Trust Center settings of the respective Office application ($($appComObjectString.Split('.')[0]))."

# Create the specific Office Application COM object
$appObject = New-Object -ComObject $appComObjectString
$appObject.Visible = $false
$appObject.DisplayAlerts = $false # For Word, this might be $appObject.Application.DisplayAlerts = wdAlertsNone (0)

# Variables for COM cleanup
$documentObject = $null # This will hold the Workbook, Document, or Presentation

try {
    # Adjust DisplayAlerts for Word specifically if needed (more granular control)
    if ($appComObjectString -eq "Word.Application") {
        # wdAlertsNone = 0, wdAlertsMessageBox = -2, wdAlertsAll = -1
        # Setting DisplayAlerts to $false on the app object is often enough.
        # $appObject.DisplayAlerts = 0 # wdAlertsNone - Suppresses all alerts and messages
    }

    # For newer Office versions, consider AutomationSecurity if strictly needed,
    # but "Trust Access to VBA Project" is paramount for reading.
    # if ($appObject.Version -ge "12.0") { # Office 2007 and later
    #     $appObject.AutomationSecurity = 2 # msoAutomationSecurityByUI (usually safest for reading)
    # }

    # Open the file
    Write-Host "Attempting to open $FilePath with $($appComObjectString.Split('.')[0])..."
    if ($appComObjectString -eq "PowerPoint.Application") {
        # PowerPoint Open method: Open(FileName, [ReadOnly As MsoTriState = msoFalse], [Untitled As MsoTriState = msoFalse], [WithWindow As MsoTriState = msoTrue])
        # We want ReadOnly if possible, and no new window if running invisibly.
        # msoTrue = -1, msoFalse = 0
        $documentObject = $appObject.($documentsOrPresentationsCollection).$openMethodName($FilePath, -1, 0, 0) # ReadOnly, No Untitled, No Window
    } else {
        # Excel/Word Open method: Open(FileName, [UpdateLinks], [ReadOnly])
        $documentObject = $appObject.($documentsOrPresentationsCollection).$openMethodName($FilePath, $false, $true) # UpdateLinks=$false, ReadOnly=$true
    }

    if (-not $documentObject) {
        Write-Error "Failed to open the file: $FilePath"
        throw "FileOpenFailed"
    }

    Write-Host "Successfully opened: $($documentObject.Name)"
    Write-Host "Reading macros from: $($documentObject.Name)"
    Write-Host "--------------------------------------------------"

    # Access the VBA project (VBProject property is common)
    # Check if the document has a VBA project
    $hasVBProjectProperty = $null
    if ($appComObjectString -eq "Excel.Application") {
        $hasVBProjectProperty = $documentObject.HasVBProject
    } else {
        # Word and PowerPoint don't have a direct 'HasVBProject' like Excel.
        # We attempt to access VBProject and catch an error if it's not there or inaccessible.
        # Or, we can check if VBProject is $null after trying to access it.
        $hasVBProjectProperty = $true # Assume true, and let the try/catch handle it.
    }


    if ($hasVBProjectProperty) { # For Excel, this is an explicit check. For others, we proceed to try.
        $vbaProject = $null
        try {
            $vbaProject = $documentObject.VBProject
        } catch {
            Write-Warning "Could not access the VBA project. This can happen if:"
            Write-Warning "1. The file has no VBA macros."
            Write-Warning "2. 'Trust access to the VBA project object model' is NOT enabled in the respective Office application's Trust Center."
            Write-Warning "3. The file's VBA project is password protected."
            # No need to re-throw here if $vbaProject remains $null
        }

        if ($vbaProject) {
            Write-Host "VBA Project Name: $($vbaProject.Name)"
            Write-Host "Modules found: $($vbaProject.VBComponents.Count)"
            Write-Host "--------------------------------------------------"

            foreach ($component in $vbaProject.VBComponents) {
                $componentName = $component.Name
                # Determine component type as string more generically
                $componentTypeString = try { $component.Type.ToString() } catch { "Unknown" }

                Write-Host "Component Name: ${componentName}" # Fixed variable interpolation
                Write-Host "Component Type: $componentTypeString (Raw Value: $($component.Type))"

                if ($component.CodeModule) {
                    $linesOfCode = $component.CodeModule.CountOfLines
                    if ($linesOfCode -gt 0) {
                        Write-Host "Code in ${componentName}:" # Fixed variable interpolation
                        Write-Host "---------------------------"
                        $macroCode = $component.CodeModule.Lines(1, $linesOfCode)
                        Write-Output $macroCode
                        Write-Host "---------------------------"
                    } else {
                        Write-Host "No code found in ${componentName}."
                        Write-Host "---------------------------"
                    }
                } else {
                    Write-Host "No CodeModule for component ${componentName} (this might be expected for some component types)."
                    Write-Host "---------------------------"
                }
            }
        } else {
            # This else block is now more likely to be hit for Word/PowerPoint if no project, or if Excel's HasVBProject was false.
            Write-Warning "No VBA project found or accessible in '$($documentObject.Name)'."
            Write-Warning "Ensure 'Trust access to the VBA project object model' is enabled in the $($appComObjectString.Split('.')[0]) Trust Center."
        }
    } else { # This branch primarily for Excel if HasVBProject is explicitly false.
         Write-Warning "The file '$($documentObject.Name)' does not report having a VBA project (e.g., Excel's HasVBProject is false)."
    }
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
    if ($_.Exception.Message -like "*0x800A03EC*") {
         Write-Warning "This error (often 0x800A03EC) can indicate that the file is corrupt, not a valid Office document for the specified application, or Excel/Word cannot access it."
    }
    # Additional specific error checks can be added here.
}
finally {
    # Close the document/workbook/presentation
    if ($documentObject) {
        try {
            # For Word, the Close method doesn't take arguments in the same way for just closing without saving.
            # It has a SaveChanges parameter (WdSaveOptions: wdDoNotSaveChanges = 0, wdSaveChanges = -1, wdPromptToSaveChanges = -2)
            if ($appComObjectString -eq "Word.Application") {
                 $documentObject.Close([ref]0) # wdDoNotSaveChanges
            } elseif ($appComObjectString -eq "PowerPoint.Application") {
                 $documentObject.Close() # PowerPoint's Close has no direct save option, relies on Saved property
            } else { # Excel
                 $documentObject.Close($false) # $false means don't save changes
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