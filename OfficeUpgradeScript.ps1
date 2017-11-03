# A few helpful functions

function GetIPSubnet () {
    $nic = gwmi -computer . -class "win32_networkadapterconfiguration" | Where-Object {$_.defaultIPGateway -ne $null}
    $IP = $nic.ipaddress | select-object -first 1
    $IPSubnet = (([ipaddress] $ip).GetAddressBytes()[0..2] -join ".")

    Write-Output -InputObject $IPSubnet
} # Simple function to get the ip subnet of a computer

function GetSite ($IPAddress) {
    switch ($IPAddress) {
        "10.5.5" {"site01"}
        "10.19.19" {"site02"}
        "default" {"AMG"}
    } # Switch
} # Simple function to get the location of the ONSServer based upon IP Subnet

function GetONSServer ($SiteID) {
switch ($SiteID) {
     "site01" {"site01-sv-1"}
     "site02" {"site02-sv-1"}
    default{"site00-sv-1"}
} # Simple function to get the location of the ONSServer based upon site ID

# Usage 
$ComputerSubnet = GetIPSubnet
$Site = GetSite($ComputerSubnet)
$ONSServer = GetONSServer($Site)

# 64 bit test
[boolean]$Is64Bit = [boolean]((Get-WmiObject -Class Win32_Processor | Where-Object { $_.DeviceID -eq 'CPU0' } | Select-Object -ExpandProperty AddressWidth) -eq '64')

# Determine Office Directory
if($Is64Bit){
$dirOffice = Join-Path "${env:ProgramFiles(X86)}" "Microsoft Office"
} Else {
$dirOffice = Join-Path "${env:ProgramFiles}" "Microsoft Office"
} # 64 or 32 bit Office Path


#Check to see if Office 365 ProPlus is already installed, exit if already installed
if (Get-InstalledApplication -Name "Microsoft Office 365 ProPlus - en-us"){
    Write-Log "Microsoft Office 365 ProPlus was detected. Program will now exit."
    Exit-Script -ExitCode 0
    }
else # Begin "Remove old office else"
{

# Show Welcome Message for Powershell v2 (There is a bug in the Toolkit main script that breaks the Date deadline without the -DeferTimes parameter.
Show-InstallationWelcome -CloseApps "excel,groove,onenote,infopath,onenote,outlook,mspub,powerpnt,winword,winproj,visio,communicator,lync" -CheckDiskSpace -RequiredDiskSpace 2048 -ForceCloseAppsCountdown 300 -TopMost $true 

# Display Pre-Install cleanup status
Show-InstallationProgress "Performing Pre-Install cleanup. This may take some time. Please wait... Please do not power off or restart your computer."

# This line can be edited or repeated for office13, office14, etc
if (Test-Path ("$dirOffice\Office12\excel.exe")){
    
    Write-Log "Microsoft Office 2007 was detected. Will be uninstalled."
    # Test path to ONS to remove old version, if can't reach server then quit
        if (Test-Path "\\$ONSServer\ons\office12\setup.exe"){
	        Execute-Process -FilePath "\\$ONSServer\ons\office12\setup.exe" -Arguments "/uninstall ENTERPRISE /config ""\\$ONSServer\ons\office12\ENTERPRISE.ww\silentuninstallconfig.xml""" -IgnoreExitCodes "3010"
        }
        Else {
            Write-Log "Could not connect to Office Network Share to remove Office 2007."
            Show-DialogBox -Title "Error contacting network share" -Text "Microsoft Office Setup was unable to contact the server for required files. Please contact the helpdesk for assistance" -Icon Stop -Buttons OK
            Exit-Script -ExitCode 99999
        }# Could not find ONS Error
} # If office found

# Remove Visio Viewer if installed
    If (Get-InstalledApplication 'visio viewer'){
            Write-Log -Message "Removing Visio Viewer"
            Remove-MSIApplications 'visio viewer'
    } # End Remove Visio Viewer 

} # end "remove old office" else

# Begin Install Office
Show-InstallationProgress "Installing Office 365 Professional Plus. This may take some time. Please wait...Please do not power off or restart your computer." -TopMost $true
Execute-Process -FilePath "Setup.exe" -Arguments "/configure `"$($site)Configuration.xml`"" -WindowStyle Hidden -IgnoreExitCodes "3010"