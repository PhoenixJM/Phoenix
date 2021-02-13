## From https://gallery.technet.microsoft.com/Offline-Servicing-of-VHDs-df776bda
param (
    [Parameter(Mandatory=$true)] [string]$VhdPath,
    [Parameter(Mandatory=$true)] [string]$MountDir,
    [Parameter(Mandatory=$true)] [string]$WsusServerName,
    [Int32]$WsusServerPort = 8530,
    [string]$WsusTargetGroupName,
    [Parameter(Mandatory=$true)] [string]$WsusContentPath,
    [string]$LogFile = "$VhdPath-dism.log",
    [switch]$Confirm = $False,
    [switch]$WhatIf = $False
)

if ($Debug) {
    $global:DebugPreference = "Continue"
    Write-Debug "Debug messages enabled."
} else {
    $global:DebugPreference = "SilentlyContinue"
}
if ($Verbose) {
    $global:VerbosePreference = "Continue"
    Write-Debug "Verbose messages enabled."
} else {
    $global:VerbosePreference = "SilentlyContinue"
}

Set-StrictMode -Version 2

Write-Debug "Performing sanity checks ..."
if (-Not $(Test-Path "$VhdPath")) {
    Write-Error "VHD specified through parameter -VhdPath does not exist. Aborting."
    Return
}
if (-Not $(Test-Path "$MountDir")) {
    Write-Error "Mount directory specified through -MountDir does not exist. Aborting."
    Return
}
if (-Not $(Test-Path "$WsusContentPath")) {
    Write-Error "WSUS content specified through -WsusContentPath does not exist. Aborting."
    Return
}

Write-Debug "Resetting log file ..."
"" | Set-Content "$LogFile"

Write-Debug "Loading assemblies ..."
# Namespace: http://msdn.microsoft.com/en-us/library/microsoft.updateservices.administration%28v=VS.85%29.aspx
# DEPRECATED: [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | Out-Null
if (-Not $(Test-Path "$Env:ProgramFiles\Update Services\Api\Microsoft.UpdateServices.Administration.dll")) {
    Write-Error "Unable to load assembly for WSUS. Please install the WSUS snapin. Aborting."
    Return
}
Add-Type -Path "$Env:ProgramFiles\Update Services\Api\Microsoft.UpdateServices.Administration.dll"

Write-Debug "Instantiating WSUS object ..."
# Establish connection with WSUS server
# AdminProxy: http://msdn.microsoft.com/de-de/library/microsoft.updateservices.administration.adminproxy_members%28v=vs.85%29.aspx
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($WsusServerName, $False, $WsusServerPort)

Write-Verbose "Creating update scope ..."
# Build an update scope to specify which updates to process
# UpdateScope: http://msdn.microsoft.com/en-us/library/microsoft.updateservices.administration.updatescope_members%28v=vs.85%29.aspx
$UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
# Only approved updates
# ApprovedStates: http://msdn.microsoft.com/en-us/library/microsoft.updateservices.administration.approvedstates%28v=vs.85%29.aspx
$UpdateScope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::LatestRevisionApproved
# Only updates which are not installed
# UpdateInstallationStates: http://msdn.microsoft.com/en-us/library/microsoft.updateservices.administration.updateinstallationstates%28v=vs.85%29.aspx
#$UpdateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Administration.UpdateInstallationStates]::NotInstalled
# Select updates released since the first day of the current month
#$Now = Get-Date
#$UpdateScope.FromArrivalDate = $Now.AddDays(-1 * ($Now.Day - 1))
# Only consider updates approved for the specified computer target group
if ($WsusTargetGroupName) {
    $TargetGroup = $wsus.GetComputerTargetGroups() | where { $_.Name -eq $WsusTargetGroupName }
    if ($TargetGroup) {
        $UpdateScope.ApprovedComputerTargetGroups.Add($TargetGroup) | Out-Null
    } else {
        Write-Error "WSUS computer target group called $WsusTargetGroupName was not found. Aborting."
        Return
    }
}

Write-Verbose "Mouting VHD file ..."
# Mount VHD file using dism
if ($(dism /Get-MountedImageInfo | where { $_.toLower().Contains($MountDir.ToLower()) } | Measure-Object -Line).Lines -eq 1) {
    Write-Warning "VHD is already mounted. Continuing in 10 seconds ..."
    Start-Sleep 10

} else {
    dism /Mount-Image /ImageFile:"$VhdPath" /Index:1 /MountDir:"$MountDir" | Add-Content "$LogFile"
    if ($(dism /Get-MountedImageInfo | where { $_.toLower().Contains($MountDir.toLower()) } | Measure-Object -Line).Lines -eq 0) {
        Write-Error "Error mounting VHD. Aborting."
        Return
    }
}

Write-Verbose "Fetching selected updates ..."
# Collect updates and process them individually
# IUpdateServer: http://msdn.microsoft.com/de-de/library/microsoft.updateservices.administration.iupdateserver_members%28v=vs.85%29.aspx
# IUpdate: http://msdn.microsoft.com/de-de/library/microsoft.updateservices.administration.iupdate_members%28v=vs.85%29.aspx
$Installables = @{}
$wsus.GetUpdates($UpdateScope) | ForEach {
    Write-Debug "Processing hotfix $($_.Title) ..."

    # Get the files associated with an update and don't process PSF files
    $_.GetInstallableItems().Files | Where { $_.FileUri.LocalPath -match '.cab' } | ForEach {
        # Substitute the WSUS content path and replace slashes with backslashes
        $FileName = $_.FileUri.LocalPath.Replace('/Content', "$WsusContentPath").Replace('/', '\')
        Write-Debug "Processing installable file $FileName ..."

        # Make sure that the file really exists
        if ($(Test-Path "$FileName") -And -Not $Installables.ContainsKey($FileName)) {
            $Installables.Add($FileName, $_.Name)

        } else {
            Write-Debug "Installable file $FileName does not exist or has already been processed. Skipping."
        }
    }
}

$Confimed = $False
if ($Confirm) {
    Write-Debug "Need user confirmation to continue ..."
    $Input = Read-Host 'Apply updates? (y/n) [n]'
    if ($Input.ToLower().Equals("y")) {
        Write-Debug "User has confirmed to continue."
        $Confimed = $True
    }
}

if (-Not $Confirm -Or $Confimed) {
    Write-Verbose "Enumerating collected updates ..."
    $Installables.Keys | ForEach {
        $FileName = $_
        $Title = $Installables.Get_Item($FileName)

        Write-Host "Applying installable file $FileName ($Title) ..."
        # Add the update as an additional package to the mounted VHD file
        $PackageInfo = dism /Image:"$MountDir" /Get-PackageInfo /PackagePath:"$FileName"
        if ($($PackageInfo | where { $_ -eq "Applicable : Yes" } | Measure-Object -Line).Lines -eq 0) {
            Write-Debug "Package is not applicable to VHD. Skipping."

        } elseif ($($PackageInfo | where { $_ -eq "Install Time : " } | Measure-Object -Line).Lines -eq 0) {
            Write-Debug "Package is already installed. Skipping."

        } elseif ($($PackageInfo | where { $_ -eq "Completely offline capable : Yes" } | Measure-Object -Line).Lines -eq 0) {
            Write-Warning "Package does not support offline servicing. Skipping."

        } else {
            $AddPackage = dism /Image:"$MountDir" /Add-Package /PackagePath:"$FileName"
            if ($($AddPackage | where { $_ -eq "The operation completed successfully." } | Measure-Object -Line ).Lines -eq 1) {
                Write-Debug "Successfully applied."
            } else {
                Write-Warning "Failed to apply. See log file $LogFile for details."
            }
        }
    }
}

if (-Not $WhatIf) {
    Write-Host "Commiting changes to VHD file ..."
    # Unmount the VHD file
    dism /Unmount-Image /MountDir:"$MountDir" /Commit | Add-Content "$LogFile"
} else {
    Write-Verbose "Discarding changes to VHD file ..."
    # Unmount the VHD file
    dism /Unmount-Image /MountDir:"$MountDir" /Discard | Add-Content "$LogFile"
}

Write-Verbose "Finished."