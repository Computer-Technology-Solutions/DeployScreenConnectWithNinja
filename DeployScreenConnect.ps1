#Default Custom 1: Company Name
#Default Custom 2: Site Name
#Default Custom 3: Department Name
#Default Custom 4: Device Type
#Adjust for your own uses at the point where it's set up further down the script

#The following takes in the parameter for the hostname of the screenconnect URL and validates it as a hostname
[CmdletBinding()]
param (
    [Parameter(
        Mandatory,
        HelpMessage = "Enter the hostname of your ScreenConnect instance, such as 'example.screenconnect.com'",
        Position = 0
    )]
    [Alias("Host","URL")]
    [ValidateScript({
        #Regex pattern mostly checks but doesn't check for a label starting or ending with a dash
        $hostnameregex = "^(?:[a-zA-Z0-9\-]{0,63}\.)*(?:[a-zA-Z0-9\-]{0,63})+" 
        $h = $PSItem
        
        #if it doesn't match the regex pattern, it's not valid
        If ($h -notmatch $hostnameregex) {return $false}

        #Split the hostname out into labels and make sure they don't start or end with a dash
        $labels = $h.split(".")
        foreach ($label in $labels) {
            If (($label[0] -eq "-") -or ($label[-1] -eq "-")) {return $false}
        }
        return $true
    })]
    [string]
    $SCHost
)

#Better error handling for Ninja
#Normal errors are so ugly on their own in the activity output
#Also you can take conditional actions based on the exit code
function myErrorMessage {
    param (
        [Parameter(Mandatory)]$ErrorNumber,
        [Parameter(Mandatory)]$ErrorMessage,
        [Parameter()]$ErrorObject
    )
    Write-Output $ErrorMessage
    If ($ErrorObject) {Write-Output $ErrorObject}
    Exit $ErrorNumber
}

#Set the temp folder location
If ($env:TEMP) {$tempFolder = $env:TEMP}
Else {$tempFolder = "C:\Windows\Temp"}

Write-Output "Getting company info..."
If (-not $env:NINJA_ORGANIZATION_NAME) {myErrorMessage -ErrorNumber 31 -ErrorMessage "Unable to get organization name from environmental variable."}
If (-not $env:NINJA_LOCATION_NAME) {myErrorMessage -ErrorNumber 32 -ErrorMessage "Unable to get location name form environmental variable."}

$CompanyName = $env:NINJA_ORGANIZATION_NAME
$SiteName = $env:NINJA_LOCATION_NAME

Write-Output "Company is $CompanyName"
Write-Output "Site is $SiteName"

Write-Output "Getting computer type..."

$validInfo = $false

#We need to figure out if it's a client or server
#Prefer Get-ComputerInfo. Use that first.

If (Get-Command -Name "Get-ComputerInfo" -ErrorAction SilentlyContinue) {
    $ComputerInfo = Get-ComputerInfo
    $InstallationType = $ComputerInfo.WindowsInstallationType
    If ($InstallationType -in @("Client","Server")) {$validInfo = $true}
}

#If that doesn't return valid info, fall back to the WMI method.
#If that fails, error out.

If (-not $validInfo) {
    Try {$osInfo = Get-WmiObject -Class Win32_OperatingSystem} Catch {myErrorMessage -ErrorNumber 27 -ErrorMessage "Unable to query WMI for OS Info" -ErrorObject $PSItem}
    #Convert number to usable string
    switch ($osInfo.ProductType) {
        1 {$InstallationType = "Client"}
        {$PSItem -in @(2,3)} {$InstallationType = "Server"}
        Default {myErrorMessage -ErrorNumber 28 -ErrorMessage "Unable to determine device type. WMI query did not return valid value for OS Product Type."}
    }
}

#Make sure it's one of those two.
If ($InstallationType -notin @("Client","Server")) {myErrorMessage -ErrorNumber 22 -ErrorMessage "Unable to determine client/server from get-computerinfo"}


#If it's a server, we'll go straight to calling it a server. We don't care if it's physical or virtual.
If ($InstallationType -eq "Server") {$DeviceType = "Server"}

#Define form factor types
$desktopTypes = @(3,4,5,6,7,13,15,16,35)
$laptopTypes = @(8,9,10,11,12,14,30,31,32)
$tabletTypes = @(30)

#If it's a workstation, determine the chassis type and use that to determine the type.
#If it doesn't match one of the known types, just call it a "workstation"
If ($InstallationType -eq "Client") {
    Try {$chassisType = Get-CimInstance -ClassName Win32_SystemEnclosure | Select-Object -ExpandProperty ChassisTypes} Catch {myErrorMessage -ErrorNumber 23 -ErrorMessage "Unable to get WMI information about chassis type." -ErrorObject $PSItem}
    switch ($chassisType) {
        { $_ -in $desktopTypes } {
            $DeviceType =  "Desktop"
        }
        { $_ -in $laptopTypes } {
            $DeviceType = "Laptop"
        }
        { $_ -in $tabletTypes } {
            $DeviceType =  "Tablet"
        }
        17 {"Main System Chassis"}
        18 {"Expansion Chassis"}
        19 {"SubChassis"}
        20 {"Bus Expansion Chassis"}
        21 {"Peripheral Chassis"}
        22 {"RAID Chassis"}
        23 {"Rack Mount Chassis"}
        24 {"Sealed-case PC"}
        25 {"Multi-system chassis"}
        26 {"Compact PCI"}
        27 {"Advanced TCA"}
        28 {"Blade"}
        29 {"Blade Enclosure"}
        33 {"IoT Gateway"}
        34 {"Embedded PC"}
        36 {"Stick PC"}
        Default {
            $DeviceType = "Unknown"
        }
    }
}

<#
Other = 1
Unknown = 2
Desktop = 3
Low Profile Desktop = 4
Pizza Box = 5
Mini Tower = 6
Tower = 7
Portable = 8
Laptop = 9
Notebook = 10
Hand Held = 11
Docking Station = 12
All in One = 13
Sub Notebook = 14
Space-Saving = 15
Lunch Box = 16
Main System Chassis = 17
Expansion Chassis = 18
SubChassis = 19
Bus Expansion Chassis = 20
Peripheral Chassis = 21
RAID Chassis = 22
Rack Mount Chassis = 23
Sealed-case PC = 24
Multi-system chassis = 25
Compact PCI = 26
Advanced TCA = 27
Blade = 28
Blade Enclosure = 29
Tablet = 30
Convertible = 31
Detachable = 32
IoT Gateway = 33
Embedded PC = 34
Mini PC = 35
Stick PC = 36
#>

Write-Output "Device type is $DeviceType"

#Escape strings so the names end up right
$CompanyName = [System.Uri]::EscapeDataString($CompanyName)
$SiteName = [System.Uri]::EscapeDataString($SiteName)
$Hostname = [System.Uri]::EscapeDataString($((hostname).toUpper()))
$DeviceType = [System.Uri]::EscapeDataString($DeviceType)

#Translate our variable names to custom property to clarify URL
#If you grabbed this off github, just ignore Custom 5. That's just for me.
#Adjust any of these for your internal use if you'd like.

$Custom1 = $CompanyName
$Custom2 = $SiteName
$Custom3 = ""
$Custom4 = $DeviceType
$Custom5 = "Yes"
$Custom6 = ""
$Custom7 = ""
$Custom8 = ""

#Derive the download url from the information we had earlier
$DownloadURL = "https://$($SCHost)/Bin/ConnectWiseControl.ClientSetup.msi?e=Access&y=Guest&t=$($Hostname)&c=$($Custom1)&c=$($Custom2)&c=$($Custom3)&c=$($Custom4)&c=$($Custom5)&c=$($Custom6)&c=$($Custom7)&c=$($Custom8)"

#Set the download location
$Installer = "${tempFolder}\ConnectWiseControl.ClientSetup.msi"

#Path to log file with date/time
$logPath = "${tempFolder}\ScreenConnectInstall_$((Get-Date -Format `"o`").Replace(`":`",`"`")).log"

If (Test-Path -Path $Installer) {
    Write-Output "Removing old installer..."
    Try {Remove-Item -Force -Path $Installer}
    Catch {myErrorMessage -ErrorNumber 29 -ErrorMessage "Unable to remove previous installer" -ErrorObject $PSItem}
}

Write-Output "Starting download..."

$rangeRequestErrorMessage = "The server does not support the necessary HTTP protocol. Background Intelligent Transfer Service (BITS) requires that the server support the Range protocol header."
$switchedPrority = $false

#Download using BITS
Try {$TransferJob = Start-BitsTransfer -Source $DownloadURL -Destination $Installer -Asynchronous -Priority Low} Catch {myErrorMessage -ErrorNumber 24 -ErrorMessage "Unable to start BITS transfer download of installer" -ErrorObject $PSItem}

#Loop through, checking progress every 15 seconds. Do so while the installer isn't downloaded.
While (-not (Test-Path $Installer)) {
    #Get the progress
    $Progress = (Get-BitsTransfer -JobId $TransferJob.JobId)
    $JobState = $Progress.JobState

    #If there's an error, drill down into that
    If ($JobState -eq "Error") {

        #If the error is due to range requests being blocked, then this can usually be solved by changing the priority to the default.
        #If we haven't already tried switching it to the default, and this is the error we're receving, then do so.
        If ((-not $switchedPrority) -and ($Progress.ErrorDescription -like "*$rangeRequestErrorMessage*")) {
            Write-Output "Range requests blocked. Switching to unspecified priority."
            #Clean up after ourselves before starting a new job
            Try {Remove-BitsTransfer -BitsJob $TransferJob} Catch {myErrorMessage -ErrorNumber 31 -ErrorMessage "Unable to remove previous BITS transfer job" -ErrorObject $PSItem}
            Try {$TransferJob = Start-BitsTransfer -Source $DownloadURL -Destination $Installer -Asynchronous} Catch {myErrorMessage -ErrorNumber 30 -ErrorMessage "Unable to start BITS transfer download of installer" -ErrorObject $PSItem}
            $switchedPrority = $true
            $Progress = (Get-BitsTransfer -JobId $TransferJob.JobId)
            $JobState = $Progress.JobState
        }
        Else {
            myErrorMessage -ErrorNumber 31 -ErrorMessage "Unable to download.`nError: ${Progress.ErrorDescription}"
        }
    }
    If ($JobState -eq "Transferred") {
        #If the installer isn't there, but the state is transferred, then it needs to change it from a temp file to the actual file. Use resume to do this.
        Try {Resume-BitsTransfer -BitsJob $TransferJob} Catch {myErrorMessage -ErrorNumber 25 -ErrorMessage "Unable to resume BITS transfer" -ErrorObject $PSItem}
    }
    Else {
        #Otherwise, give a status update
        Write-Output "Downloading Installer. Job state is $JobState. Progress: $(($Progress.BytesTransferred/$Progress.BytesTotal)*100)%"
    }
    #Repeat every 15 seconds until finished
    Start-Sleep -Seconds 15
}

Write-Output "Starting installation..."
#Try the installer
Try {$installerProcess = (Start-Process -FilePath "msiexec.exe" -ArgumentList "/package $installer /qn /l `"$logPath`"" -NoNewWindow -PassThru)}
Catch {myErrorMessage -ErrorMessage "Unable to install with msiexec" -ErrorNumber 26 -ErrorObject $PSItem}

#Wait for it to finish
While (-not $installerProcess.HasExited) {Start-Sleep -Seconds 5}

#Write the install log to the console

$installLogContent = (Get-Content -Raw -Path $logPath)

Write-Output $installLogContent