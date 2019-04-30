<#
.SYNOPSIS
  An automation script to build the necessary WinSxS repair sources for Windows 10 1809+

.DESCRIPTION
  Beginning in Windows 10 1809 cumulative updates do not include the full binaries for
  updates to files since RTM. The new mechanism relies on reverse/forward/null differentials
  stored on a workstation to perform the update. However, in the case of unrecoverable
  corruption on a file the update instructs the computer to reach out to Windows Update to
  repair the file. In environments were Windows Update has been made unavailable whether
  through group policy or network filtering, you must define your own local Windows Repair
  Source (WRS). This is greatly complicated by the fact that all of the necessary binaries 
  are not stored long term - so defining a Windows Repair Source requires you to either 
  create a new WRS for every other update (N and N-1) or to store all the files contained in
  WinSxS over the course of many updates.

  This script will automatically create a folder containing all versions of WinSxS folders
  over the course of the monthly update cycle. It does so in the following steps:
    1. Mount a RTM WIM
    2. Update the WIM with a cumulative update
    3. Copy the WinSxS changes to a designated WRS folder per OS
    4. Repeat steps 2-3 in order until all updates (starting with the oldest)
       have been applied.

.NOTES
  Author: Nathan Ziehnert (@theznerd / nathan@z-nerd.com)
  Version: 1.0
  Release Notes:
    - 1.0: Initial Release

.LINK
  https://z-nerd.com/blog/2019/04/29-windows-repair-source/

.EXAMPLE
  .\Build-WRS.ps1 -FileLogging -LogFilePath c:\Windows\Temp\WRS.log -ImagePath D:\WRS\_OS -ImageIndex 3 -WRSPath D:\WRS -Verbose

.PARAMETER FileLogging
  A switch to enable logging to file.

.PARAMETER LogFilePath
  Path where you would like to write the log file to.

.PARAMETER ImagePath
  Path to the base folder for all your OS folders. These
  OS folders should contain the WIM file and all updates.

.PARAMETER ImageIndex
  An integer representing the index of the image you want
  to apply against. It would be best to run this against
  the index of the highest version you have licenses for.

.PARAMETER WRSPath
  Path where you want to store your WRS files. The script
  converts the OS to a version based on RTM. So for
  multiple OSes you would have multiple folders below this
  target (X:\WRS\17763.1, X:\WRS\17134.48, etc.)

.PARAMETER Debug
  Enables some debug logging.

.PARAMETER Verbose
  Enables verbose logging to the console.
#>
param
(
    # Enable Log File
    [Parameter()]
    [switch]
    $FileLogging,

    # Log File Path
    [Parameter()]
    [string]
    $LogFilePath="$ENV:TEMP\Build-WRS.log",
    
    # OS Image / Update Path
    [Parameter(Mandatory=$true)]
    [String]
    $ImagePath,

    # OS Index
    [Parameter(Mandatory=$true)]
    [Int]
    $ImageIndex,

    # WRS Path
    [Parameter(Mandatory=$true)]
    [String]
    $WRSPath
)

#region FUNCTIONS
function Write-Logs
{
    <#
    .SYNOPSIS
    Creates a log entry in all applicable logs (CMTrace compatible File and Verbose logging).
    #>
    [CmdletBinding()]
    Param
    (
        # Log File Enabled
        [Parameter()]
        [switch]
        $FileLogging,

        # Log File Path
        [Parameter()]
        [string]
        $LogFilePath,

        # Log Description
        [Parameter(mandatory=$true)]
        [string]
        $Description,

        # Log Source
        [Parameter(mandatory=$true)]
        [string]
        $Source,

        # Log Level
        [Parameter(mandatory=$false)]
        [ValidateRange(1,4)]
        [int]
        $Level,

        # Debugging Enabled
        [Parameter(mandatory=$false)]
        [switch]
        $Debugging
    )
    
    # Get Current Time (UTC)
    $dt = [DateTime]::UtcNow

    $lt = switch($Level)
    {
        1 { 'Informational' }
        2 { 'Warning' }
        3 { 'Error' }
        4 { 'Debug' }
    }

    if($FileLogging)
    {
        # Create Pretty CMTrace Log Entry
        if(($Level -lt 4) -or $Debugging)
        {
            if($Level -ne 1)
            {
                $cmtl  = "<![LOG[`($lt`) $Description]LOG]!>"
            }
            else
            {
                $cmtl  = "<![LOG[$Description]LOG]!>"
            }
            $cmtl += "<time=`"$($dt.ToString('HH:mm:ss.fff'))+000`" "
            $cmtl += "date=`"$($dt.ToString('M-d-yyyy'))`" "
            $cmtl += "component=`"$Source`" "
            $cmtl += "context=`"$($ENV:USERDOMAIN)\$($ENV:USERNAME)`" "
            $cmtl += "type=`"$Level`" "
            $cmtl += "thread=`"$($pid)`" "
            $cmtl += "file=`"`">"
    
            # Write a Pretty CMTrace Log Entry
            $cmtl | Out-File -Append -Encoding UTF8 -FilePath "$LogFilePath"
        }
    }

    if(($Level -lt 4) -or $Debugging)
    {
        if($VerbosePreference -ne 'SilentlyContinue')
        {
            Write-Verbose -Message "[$dt] ($lt) $Source`: $Description"
        }
    }
}

function Mount-WIM
{
    param
    (
        # OS Object
        [Parameter(Mandatory=$true)]
        [OS]
        $OSObj,

        # Unmount Switch
        [Parameter()]
        [Switch]
        $Unmount
    )
    
    if(!$Unmount)
    {
        try
        {
            Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Mount Image" -Description "Mounting $($OSObj.Name) version $($OSObj.Version) to $($OSObj.TempPath)" -Level 1 -Debugging:$Debug
            if(!(Test-Path $OSObj.TempPath)){ New-Item -Path $OSObj.TempPath -ItemType Directory -Force }
            Mount-WindowsImage -ImagePath $OSObj.Path -Index $OSObj.Index -Path $OSObj.TempPath
        }
        catch
        {
            Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Mount Image" -Description "Failed to mount image $WIMPath with the following exception: $($Error[0].Exception)" -Level 3 -Debugging:$Debug
            return $false
        }
        Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Mount Image" -Description "Successfully mounted image." -Level 1 -Debugging:$Debug
    }
    else
    {
        Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Mount Image" -Description "Unmounting $($OSObj.Name) version $($OSObj.Version) from $($OSObj.TempPath)" -Level 1 -Debugging:$Debug
        Dismount-WindowsImage -Path $OSObj.TempPath -Discard
        Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Mount Image" -Description "Successfully unmounted image." -Level 1 -Debugging:$Debug
    }
    return $true
}

function Install-CumulativeUpdateOffline
{
    param
    (
        # Cumulative Update Path
        [Parameter(Mandatory=$true)]
        [String]
        $UpdatePath,

        # Offline OS Path
        [Parameter(Mandatory=$true)]
        [String]
        $OfflineOSPath
    )
    # Get KB
    $KBArticle = (([regex]::match($UpdatePath, "-([^/])+-")).Groups[0].Value).Replace("-","")

    try
    {
        Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Install CU" -Description "Installing Cumulative Update $KBArticle" -Level 1 -Debugging:$Debug
        Add-WindowsPackage -Path "$OfflineOSPath" -PackagePath "$UpdatePath" -Verbose:$Verbose
    }
    catch
    {
        Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Install CU" -Description "Failed to install CU $KBArticle with the following exception: $($Error[0].Exception)" -Level 3 -Debugging:$Debug
        return $false
    }
    Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Mount Image" -Description "Successfully installed CU $KBArticle." -Level 1 -Debugging:$Debug
    return $true
}

function Merge-WinSxS
{
    param
    (
        # Cumulative Path for SxS Files
        [Parameter(Mandatory=$true)]
        [String]
        $DestinationPath,

        # Updated Path To Copy From
        [Parameter(Mandatory=$true)]
        [String]
        $SourcePath
    )
    Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Merge SxS Folders" -Description "Beginning merge of SxS files." -Level 1 -Debugging:$Debug
    $CopyProcess = Start-Process -FilePath "$ENV:windir\System32\Robocopy.exe" -WindowStyle Hidden -Verbose:$Verbose -ArgumentList "`"$SourcePath`" `"$DestinationPath`" /e /xo /fft /copy:DAT /r:5 /w:5 " -Wait -PassThru
    switch($CopyProcess.ExitCode)
    {
        0 { $result = $true; $warning = $false; $description = "No files were copied." }
        1 { $result = $true; $warning = $false; $description = "All files were copied successfully." }
        2 { $result = $true; $warning = $false; $description = "There are some additional files in the destination directory that are not present in the source directory. No files were copied." }
        3 { $result = $true; $warning = $false; $description = "Some files were copied. Additional files were present. No failure was encountered." }
        5 { $result = $true; $warning = $false; $description = "Some files were copied. Some files were mismatched. No failure was encountered." }
        6 { $result = $true; $warning = $true; $description = "Additional files and mismatched files exist. No files were copied and no failures were encountered. This means that the files already exist in the destination directory." }
        7 { $result = $true; $warning = $true; $description = "Files were copied, a file mismatch was present, and additional files were present." }
        8 { $result = $false; $warning = $true; $description = "Several files did not copy." }
        default { $result = $false; $warning = $true; $description = "Several files had failures during the copy operation." }
    }

    if($result)
    {
        if($warning)
        {
            Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Merge SxS Folders" -Description "Merge Complete. WARNING: $description" -Level 2 -Debugging:$Debug
            return $result
        }
        else
        {
            Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Merge SxS Folders" -Description "Merge Complete. RESULT: $description" -Level 1 -Debugging:$Debug
            return $result
        }
    }
    else
    {
        Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Merge SxS Folders" -Description "Merge Completed with Failures. RESULT: $description" -Level 3 -Debugging:$Debug
        return $result
    }
}
#endregion FUNCTIONS

#region CLASSES
class OS
{
    [string]$Name
    [string]$Version
    [string]$Path
    [int]$Index
    [xml]$History
    [string]$TempPath
}
#endregion CLASSES

#region Script
Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Initializing" -Description "===============================================================" -Level 1 -Debugging:$Debug
Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Initializing" -Description "Starting Build-WRS" -Level 1 -Debugging:$Debug
Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Initializing" -Description "===============================================================" -Level 1 -Debugging:$Debug

# Get WIM Files
Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Initializing" -Description "Looking for WIM files in $ImagePath" -Level 4 -Debugging:$Debug
$WimFiles = Get-ChildItem -Path "$ImagePath" -Recurse -Filter "*.wim"
Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Initializing" -Description "Found $($WimFiles.Count) WIM files." -Level 4 -Debugging:$Debug

$osa = @() # OS Array

Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Initializing" -Description "Reviewing Images" -Level 1 -Debugging:$Debug
# Create OS Objects
foreach($image in $WimFiles)
{
    Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Initializing" -Description "Reviewing $($image.Name)" -Level 1 -Debugging:$Debug
    $os = [OS]::new();
    $os.Path = $image.FullName
    if(Test-Path "$(Split-Path ($os.Path))\OSHistory.xml")
    {
        $os.History = [xml](Get-Content "$(Split-Path ($os.Path))\OSHistory.xml")
    }
    else
    {
        $os.History = [xml](Get-Content (New-Item -Path "$(Split-Path ($os.Path))\OSHistory.xml" -ItemType "File" -Value "<?xml version=`"1.0`" ?>`r`n<Updates>`r`n</Updates>" -Force))
    }
    
    # Get WIM Info
    $wimInfo = Get-WindowsImage -ImagePath "$($os.Path)" -Index $ImageIndex
    $os.Name = $wimInfo.ImageName
    $os.Version = $wimInfo.Version
    $os.Index = $ImageIndex
    $os.TempPath = "$(Split-Path ($os.Path))\Temp"
    $osa += $os
}
Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Initializing" -Description "Image Review Complete" -Level 1 -Debugging:$Debug

foreach($os in $osa)
{
    Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Operating System" -Description "Processing Image: $($os.Path)" -Level 1 -Debugging:$Debug
    $result = Mount-WIM -OSObj $os
    if($result)
    {
        $RTMComplete = $true
        # Base Updates
        if(($os.History.Updates.Update | Where-Object {$_.KB -eq "RTM"}) -eq $null)
        {
            # Append RTM Version
            $u = $os.History.CreateElement("Update")
            $u.SetAttribute("KB","RTM")
            $u.SetAttribute("Applied",$false)
            $u.SetAttribute("Version",$os.Version)
            $u.SetAttribute("Path","")

            # Copy files
            $result = Merge-WinSxS -SourcePath "$($os.TempPath)\Windows\WinSxS" -DestinationPath "$WRSPath\$($os.Version)"
            if($result)
            {
                $u.SetAttribute("Applied",$true)
            }
            else
            {
                $RTMComplete = $false
            }
            $os.History.DocumentElement.AppendChild($u)
            $void = $os.History.Save("$(Split-Path ($os.Path))\OSHistory.xml")
        }
        
        if($RTMComplete)
        {
            # Get Updates
            $updates = Get-ChildItem (Split-Path $os.Path) -Filter "*.msu"
            foreach($update in $updates)
            {
                Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Update" -Description "Processing Update: $($update.FullName)" -Level 1 -Debugging:$Debug
                $allUpdates = ($os.History.Updates.Update).KB
                $UpdateKB = ([regex]::Match($update.Name,"(?i)kb\d{7}")).Value

                # Extract the XML file from the update for version number
                Start-Process "$ENV:WINDIR\System32\extrac32.exe" -ArgumentList "/Y `"$($Update.FullName)`" *.xml /L `"$(Split-Path $os.Path)`"" -Wait
                $updateXMLLoc = Get-ChildItem "$(Split-Path $os.Path)" -Filter "*$UpdateKB*.xml"
                [xml]$updateXML = Get-Content $updateXMLLoc.FullName

                if($UpdateKB -notin $allUpdates)
                {
                    Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Update" -Description "Update not found in XML file" -Level 1 -Debugging:$Debug
                    # Add it to the tracking XML
                    $u = $os.History.CreateElement("Update")
                    $u.SetAttribute("KB",$UpdateKB)
                    $u.SetAttribute("Applied",$false)
                    $u.SetAttribute("Version",$updateXML.unattend.servicing.package.assemblyIdentity.version)
                    $u.SetAttribute("Path",$update.FullName)
                    $os.History.DocumentElement.AppendChild($u)
                }
            }

            :updateLoop foreach($update in ($os.History.Updates.Update | Sort-Object {[version]$_.Version}))
            {
                if($update.Applied -eq "False")
                {
                    $result = Install-CumulativeUpdateOffline -UpdatePath "$($update.Path)" -OfflineOSPath $os.TempPath
                    if($result)
                    {
                        if(!(test-path "$WRSPath\$($os.Version)"))
                        {
                            New-Item -Path "$WRSPath\$($os.Version)" -ItemType Directory -Force
                        }
                        $update.Applied = "True"
                        $result = Merge-WinSxS -SourcePath "$($os.TempPath)\Windows\WinSxS" -DestinationPath "$WRSPath\$($os.Version)"
                        $void = $os.History.Save("$(Split-Path ($os.Path))\OSHistory.xml")
                    }
                    else
                    {
                        Mount-WIM -OSObj $os -Unmount
                        $update.Applied = "False"
                        $void = $os.History.Save("$(Split-Path ($os.Path))\OSHistory.xml")
                        break updateLoop
                    }
                }
                else
                {
                    Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Update" -Description "Update already applied - skipping." -Level 1 -Debugging:$Debug
                }
            }
        }
    }
    else
    {
        Write-Logs -FileLogging:$FileLogging -LogFilePath "$LogFilePath" -Source "Operating System" -Description "Image did not mount properly." -Level 1 -Debugging:$Debug     
    }
    $void = Mount-WIM -OSObj $os -Unmount
}
#endregion Script
