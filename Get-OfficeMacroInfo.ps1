param (
    [Parameter(Mandatory = $true, HelpMessage = "Path to the Office file (e.g., .xlsm, .docm, .pptm)")]
    [string]$FilePath,

    [Parameter(HelpMessage = "If specified, attempts to enable 'Trust access to the VBA project object model' via the registry for the current user. The script will attempt to revert this change upon completion. Requires Office applications to be restarted for changes to take full effect.")]
    [switch]$EnableTrustAccess
)

# Determine Office Application and settings based on file extension
$fileExtension = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
$appComObjectString = $null
$documentsOrPresentationsCollection = $null
$appNameForRegistry = $null

switch ($fileExtension) {
    ".xlsm" { $appComObjectString = "Excel.Application"; $appNameForRegistry = "Excel"; $documentsOrPresentationsCollection = "Workbooks" }
    ".xls" { $appComObjectString = "Excel.Application"; $appNameForRegistry = "Excel"; $documentsOrPresentationsCollection = "Workbooks" }
    ".docm" { $appComObjectString = "Word.Application"; $appNameForRegistry = "Word"; $documentsOrPresentationsCollection = "Documents" }
    ".dotm" { $appComObjectString = "Word.Application"; $appNameForRegistry = "Word"; $documentsOrPresentationsCollection = "Documents" }
    ".pptm" { $appComObjectString = "PowerPoint.Application"; $appNameForRegistry = "PowerPoint"; $documentsOrPresentationsCollection = "Presentations" }
    ".ppsm" { $appComObjectString = "PowerPoint.Application"; $appNameForRegistry = "PowerPoint"; $documentsOrPresentationsCollection = "Presentations" }
    default {
        Write-Host "[!] Unsupported file type: $fileExtension. This script supports .xlsm, .xls, .docm, .dotm, .pptm, .ppsm." -ForegroundColor Red
        exit 1
    }
}

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

$appObject = $null
$documentObject = $null
$initialAccessVBOMValue = 0
$keyExistedInitially = $false
$scriptModifiedAccessVBOM = $false
$registryPathForVBOM = $null
$officeVersion = $null
$trustAccessIsEnabled = $false
$exitDueToTrustSetting = $false

try {
    $appObject = New-Object -ComObject $appComObjectString
    $appObject.Visible = $false
    if ($appComObjectString -eq "Word.Application") {
        $appObject.DisplayAlerts = 0
    }
    else {
        $appObject.DisplayAlerts = $false
    }

    # --- Check and Optionally Enable "Trust access to the VBA project object model" setting ---
    try {
        $officeVersion = $appObject.Version
        $registryPathForVBOM = "HKCU:\Software\Microsoft\Office\$officeVersion\$appNameForRegistry\Security"
        
        $accessVBOMProperty = Get-ItemProperty -Path $registryPathForVBOM -Name "AccessVBOM" -ErrorAction SilentlyContinue
        if ($null -ne $accessVBOMProperty) {
            $keyExistedInitially = $true
            $initialAccessVBOMValue = $accessVBOMProperty.AccessVBOM
        }
        else {
            $keyExistedInitially = $false
            $initialAccessVBOMValue = 0    
        }

        if ($initialAccessVBOMValue -eq 1) {
            Write-Host "[*] 'Trust access to the VBA project object model' is currently ENABLED for $appNameForRegistry (Version $officeVersion)." -ForegroundColor Yellow
            $trustAccessIsEnabled = $true
            if ($EnableTrustAccess.IsPresent) {
                Write-Host "[*] (-EnableTrustAccess specified) No action needed as setting is already enabled." -ForegroundColor Yellow
            }
        }
        else {
            if ($EnableTrustAccess.IsPresent) {
                Write-Host "[*] 'Trust access to the VBA project object model' is currently DISABLED (Value: $initialAccessVBOMValue, Existed: $keyExistedInitially) for $appNameForRegistry (Version $officeVersion)." -ForegroundColor Yellow
                Write-Host "[*] Parameter -EnableTrustAccess specified. Attempting to enable this setting..." -ForegroundColor Yellow
                Write-Host "[*] Ensure $appNameForRegistry is closed. A restart of $appNameForRegistry will be required for this change to take full effect." -ForegroundColor Yellow

                try {
                    if (-not (Test-Path $registryPathForVBOM)) {
                        New-Item -Path $registryPathForVBOM -Force -ErrorAction Stop | Out-Null
                    }
                    Set-ItemProperty -Path $registryPathForVBOM -Name "AccessVBOM" -Value 1 -Type DWord -Force -ErrorAction Stop
                    Write-Host "[*] Successfully set AccessVBOM=1 in the registry for $appNameForRegistry." -ForegroundColor Green
                    $scriptModifiedAccessVBOM = $true
                    $trustAccessIsEnabled = $true
                }
                catch {
                    Write-Host "[!] Failed to enable 'Trust access to the VBA project object model' in the registry: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host "[!] This operation may require elevated permissions or $appNameForRegistry to be closed." -ForegroundColor Red
                    $exitDueToTrustSetting = $true
                    Write-Host "[!] Script will exit due to failure to enable required setting." -ForegroundColor Red
                }
            }
            else {
                Write-Host "[!] 'Trust access to the VBA project object model' is currently DISABLED for $appNameForRegistry (Version $officeVersion)." -ForegroundColor Red
                Write-Host "[!] This setting is REQUIRED for the script to access VBA macro information." -ForegroundColor Red
                Write-Host "[*] Please enable this setting manually or re-run with -EnableTrustAccess." -ForegroundColor DarkCyan
                $exitDueToTrustSetting = $true
            }
        }
    }
    catch {
        Write-Host "[*] Could not reliably determine or set 'Trust access to VBA project model': $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "[*] Please ensure the setting is enabled manually. Script will attempt to continue if not explicitly exiting, but VBA access may fail." -ForegroundColor Yellow
    }

    if ($exitDueToTrustSetting) {
        Write-Host "[*] Exiting script due to 'Trust access to VBA project object model' setting prerequisite not being met or failing to enable." -ForegroundColor DarkCyan
        throw "TrustAccessPrerequisiteFailed"
    }

    # --- Main Document Processing ---
    Write-Host "[*] Attempting to open '$FilePath' with '$($appComObjectString.Split('.')[0])'..." -ForegroundColor Cyan
    if ($appComObjectString -eq "PowerPoint.Application") {
        $documentObject = $appObject.($documentsOrPresentationsCollection).Open($FilePath, -1, 0, 0)
    }
    else {
        $documentObject = $appObject.($documentsOrPresentationsCollection).Open($FilePath, $false, $true)
    }

    if (-not $documentObject) {
        Write-Host "[!] Failed to open the file: $FilePath" -ForegroundColor Red
        throw "FileOpenFailed"
    }

    Write-Host "[*] Successfully opened: '$($documentObject.Name)'" -ForegroundColor Cyan
    Write-Host "[*] Reading macros from: '$($documentObject.Name)'" -ForegroundColor Cyan
    Write-Host "--------------------------------------------------"

    if (-not $trustAccessIsEnabled) {
        Write-Host "[*] 'Trust access to VBA project model' may not be effectively enabled (e.g., Office app restart needed). VBA project access might fail." -ForegroundColor Yellow
    }

    $hasVBProjectProperty = $null
    if ($appComObjectString -eq "Excel.Application") {
        $hasVBProjectProperty = $documentObject.HasVBProject
    }
    else {
        $hasVBProjectProperty = $true
    }

    if ($hasVBProjectProperty) {
        $vbaProject = $null
        try {
            $vbaProject = $documentObject.VBProject
        }
        catch {
            Write-Host "[*] Could not access the VBA project. This can happen if:" -ForegroundColor Yellow
            Write-Host "[*] 1. The file has no VBA macros." -ForegroundColor Yellow
            Write-Host "[*] 2. 'Trust access to the VBA project object model' is effectively NOT enabled (ensure $appNameForRegistry was restarted if changed)." -ForegroundColor Yellow
            Write-Host "[*] 3. The file's VBA project is password protected." -ForegroundColor Yellow
        }

        if ($vbaProject) {
            Write-Host "[*] VBA Project Name: '$($vbaProject.Name)'" -ForegroundColor DarkCyan
            Write-Host "[*] VBComponents found: $($vbaProject.VBComponents.Count)" -ForegroundColor DarkCyan
            Write-Host "--------------------------------------------------"

            foreach ($component in $vbaProject.VBComponents) {
                $componentName = $component.Name
                $componentTypeString = try { $component.Type.ToString() } catch { "Unknown" }

                Write-Host "[*] Component Name: '${componentName}'" -ForegroundColor Magenta
                Write-Host "[*] Component Type: $componentTypeString (Raw Value: $($component.Type))" -ForegroundColor Blue

                if ($component.CodeModule) {
                    $linesOfCode = $component.CodeModule.CountOfLines
                    if ($linesOfCode -gt 0) {
                        Write-Host "[=] Code in ${componentName}:" -ForegroundColor Red
                        Write-Host "----- Start -----" -ForegroundColor Red
                        $macroCode = $component.CodeModule.Lines(1, $linesOfCode)
                        Write-Host $macroCode -ForegroundColor Red
                        Write-Host "------ End ------" -ForegroundColor Red
                        Write-Host "---------------------------"
                    }
                    else {
                        Write-Host "[=] No code found in '${componentName}'." -ForegroundColor Green
                        Write-Host "---------------------------"
                    }
                }
                else {
                    Write-Host "[*] No CodeModule for component '${componentName}' (this might be expected for some component types)." -ForegroundColor DarkCyan
                    Write-Host "---------------------------"
                }
            }
        }
        else {
            Write-Host "[*] No VBA project was ultimately accessed in '$($documentObject.Name)'." -ForegroundColor DarkCyan
            Write-Host "[*] This could be normal if the file has no macros, or if Trust Access setting is not effective yet." -ForegroundColor DarkCyan
        }
    }
    else {
        Write-Host "[*] The file '$($documentObject.Name)' does not report having a VBA project (e.g., Excel's HasVBProject is false)." -ForegroundColor DarkCyan
    } # End if ($hasVBProjectProperty)
} # End main Try
catch {
    if ($_.TargetObject -ne "TrustAccessPrerequisiteFailed") {
        Write-Host "[!] An error occurred during script execution: $($_.Exception.Message)" -ForegroundColor Red
    }
    if ($_.TargetObject -eq "FileOpenFailed") {
    }
    elseif ($_.Exception.Message -like "*0x800A03EC*") {
        Write-Host "[*] An Office COM error (often 0x800A03EC) occurred. This can indicate the file is corrupt or inaccessible." -ForegroundColor Yellow
    }
}
finally {
    if ($appObject) {
        try {
            $appObject.Quit()
        }
        catch { Write-Host "[*] Could not gracefully quit Office app: $($_.Exception.Message)" -ForegroundColor Yellow }

        # Ensure COM object is released before attempting registry changes that Office might contend with during its own shutdown.
        Start-Sleep -Milliseconds 250

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($appObject) | Out-Null
        Remove-Variable appObject -ErrorAction SilentlyContinue

        # ---- ATTEMPT REGISTRY REVERSION AND PERFORM FINAL CHECK ----
        if ($scriptModifiedAccessVBOM -and $registryPathForVBOM) {
            Write-Host "[*] Script had modified AccessVBOM. Attempting to revert to initial state (Original Value: '$initialAccessVBOMValue', KeyExistedInitially: $keyExistedInitially) for $appNameForRegistry." -ForegroundColor Yellow
            try {
                if ($keyExistedInitially) {
                    Set-ItemProperty -Path $registryPathForVBOM -Name "AccessVBOM" -Value $initialAccessVBOMValue -Type DWord -Force -ErrorAction Stop
                    Write-Host "[*] Successfully attempted to set AccessVBOM back to '$initialAccessVBOMValue'." -ForegroundColor Yellow
                }
                else {
                    # Key did not exist initially, so we remove the value the script created
                    Remove-ItemProperty -Path $registryPathForVBOM -Name "AccessVBOM" -Force -ErrorAction Stop
                    Write-Host "[*] Successfully attempted to remove AccessVBOM value as it did not exist initially." -ForegroundColor Yellow
                }
                # This info message is about the UI reflecting the change we just attempted.
                Write-Host "[*] A restart of $appNameForRegistry may be required for this reversion to be fully reflected in the Trust Center UI." -ForegroundColor Yellow
            }
            catch {
                Write-Host "[!] Failed to revert AccessVBOM setting in the registry: $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "[*] The AccessVBOM setting might remain enabled. Please check $appNameForRegistry's Trust Center manually." -ForegroundColor Yellow
            }

            Start-Sleep -Seconds 1 # Give registry a moment after our reversion attempt before the final check.
        
            $finalAccessVBOMProperty = Get-ItemProperty -Path $registryPathForVBOM -Name "AccessVBOM" -ErrorAction SilentlyContinue

            if ($finalAccessVBOMProperty) {
                $finalValue = $finalAccessVBOMProperty.AccessVBOM

                if ($keyExistedInitially) {
                    if ($finalValue -eq $initialAccessVBOMValue) {
                        Write-Host "[*] AccessVBOM successfully reverted to '$initialAccessVBOMValue' and confirmed." -ForegroundColor Green
                    }
                    else {
                        Write-Host "[*] Script attempted to revert AccessVBOM to '$initialAccessVBOMValue', but the FINAL registry value is '$finalValue'." -ForegroundColor Yellow
                        Write-Host "[*] The Office application ($appNameForRegistry) may have modified this setting during or after closing." -ForegroundColor Yellow
                        Write-Host "[!] ACTION: Please manually verify and adjust 'Trust access to the VBA project object model' in $appNameForRegistry's Trust Center if needed." -ForegroundColor Yellow
                    }
                }
                else {
                    # Key was created by script, script attempted to remove it.
                    Write-Host "[*] Script attempted to remove the AccessVBOM value (as it was created by the script), but it still exists with value '$finalValue'." -ForegroundColor Yellow
                    Write-Host "[*] The Office application ($appNameForRegistry) may have re-created or modified this setting during or after closing." -ForegroundColor Yellow
                    Write-Host "[!] ACTION: Please manually verify and disable 'Trust access to the VBA project object model' in $appNameForRegistry's Trust Center." -ForegroundColor Yellow
                }
            }
            else {
                # AccessVBOM value does not exist in registry after reversion attempt
                if (-not $keyExistedInitially) {
                    # Script created it and successfully removed it.
                    Write-Host "[*] AccessVBOM successfully reverted (value removed as it was created by script) and confirmed." -ForegroundColor Green
                }
                else {
                    # Script attempted to set it to $initialAccessVBOMValue (e.g., 0), but found it deleted.
                    if ($initialAccessVBOMValue -eq 0) {
                        # If initial was 0, and now it's gone, it's still effectively disabled.
                        Write-Host "[*] AccessVBOM effectively reverted. Initial value was '$initialAccessVBOMValue', now the key is not found (effectively disabled)." -ForegroundColor Green
                    }
                    else {
                        # Initial value was something other than 0, and now it's gone.
                        Write-Host "[*] Script attempted to revert AccessVBOM to '$initialAccessVBOMValue', but the FINAL registry value is not found (it was deleted)." -ForegroundColor Yellow
                        Write-Host "[*] This means 'Trust access to the VBA project object model' is effectively disabled, but the key's absence differs from its initial state if it wasn't 0." -ForegroundColor Yellow
                        Write-Host "[!] ACTION: Please manually verify 'Trust access to the VBA project object model' in $appNameForRegistry's Trust Center." -ForegroundColor Yellow
                    }
                }
            }
        }
        elseif ($registryPathForVBOM) {
            # This means $scriptModifiedAccessVBOM was false.
            # Script did not modify AccessVBOM.
        }
    } # End of "if ($appObject)"
    Write-Host "--------------------------------------------------"
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}