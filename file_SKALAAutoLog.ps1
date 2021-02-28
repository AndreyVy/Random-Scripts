<#
.SYNOPSIS
SKALA AutoLog script

.DESCRIPTION
Gather relevant System information and Log files, Registry and other info,
which may be suitable for troubleshooting.
    
.PARAMETER Silent
Optional parameter that will make script not to stop at the end,
or give User any prompt.

.PARAMETER UploadPath
Optional network/local folder to upload resulting file.
User should have Change/Modify permissions.

.EXAMPLE
PS > .\SKALA-AutoLog.ps1

Script collects all logs locally under %Public%\SKALA.
Pauses at the end. Opens folder in Explorer

.EXAMPLE
PS > .\SKALA-AutoLog.ps1 -UploadPath \\server\sharedfolder\Logs -Silent

Script collects all logs and uploads them to given network/local folder.
Verifies if given location is present and User have Write permissions.
Otherwise resulting file will be placed under %Public%\SKALA.
Does not stops at the end, does not have any prompts to User.

.NOTES
Authors:
Igor Zaharov

#>

# Get parameters from command-line
param(
        # [Parameter(Mandatory=$false, Position = 0)]
        [Alias('UP')]
        [String] $UploadPath,

        # [Parameter(Mandatory=$false, Position = 1)]
        [Switch] $Silent 
)
# *****************************************************************************

# *****************************************************************************
# Global Output file declaration
    $varOutputFile = ($env:TEMP + "\SKALA-AutoLog-Output.log")
# *****************************************************************************

# *****************************************************************************
# Function: Check if current User has WRITE access to given location
# creates test-file with random name in given location. File will be deleted after script is finished
Function Test-Write {
    [CmdletBinding()]
    param (
        [parameter()] [ValidateScript({[IO.Directory]::Exists($_.FullName)})]
        [IO.DirectoryInfo] $Path
    )
    try {
        $testPath = Join-Path $Path ([IO.Path]::GetRandomFileName())
        [IO.File]::Create($testPath, 1, 'DeleteOnClose') > $null
        # Or...
        <# New-Item -Path $testPath -ItemType File -ErrorAction Stop > $null #>
        return $true
    } catch {
        return $false
    } finally {
        Remove-Item $testPath -ErrorAction SilentlyContinue
    }
}
# *****************************************************************************

# *****************************************************************************
# Function: Pause to use in Legacy Powershell
Function Pause ($Message = "Press any key to continue . . . ") {
    if ((Test-Path variable:psISE) -and $psISE) {
        $Shell = New-Object -ComObject "WScript.Shell"
        $Button = $Shell.Popup("Click OK to continue.", 0, "Script Paused", 0)
    }
    else {     
        Write-Host -NoNewline $Message
        [void][System.Console]::ReadKey($true)
        Write-Host
    }
}
# *****************************************************************************

# *****************************************************************************
# Function: Check for running processes
Function GetProcList {
$PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($(New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$('ProcessName','UserName','CSName','Handle',"CommandLine"))))
$Processes = Get-wmiobject -Class Win32_Process
if ($Processes -ne $null) {
    foreach ($Process in $Processes) {
        $Process | 
            Add-Member -MemberType NoteProperty -Name 'UserName' -Value $($Process.getowner().user) -PassThru | 
            Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $PSStandardMembers -PassThru
        }
}
}
# *****************************************************************************

# *****************************************************************************
# Function: Custom Function to write output to both File and Console
function Write-Feedback($msg,$ForegroundColor)
{
    If ($ForegroundColor -eq ""){$ForegroundColor="Green"}
    Write-Host $msg -ForegroundColor $ForegroundColor 
    Add-Content $varOutputFile $msg
}
# *****************************************************************************

# *****************************************************************************
# Function: Check Uninstall registry
Function CheckUninstallRegKey($UninstallKey)
{
$computername=$env:computername
$array = @()

#Create an instance of the Registry Object and open the HKLM base key
$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey(‘LocalMachine’,$computername) 
#Drill down into the Uninstall key using the OpenSubKey Method
$regkey=$reg.OpenSubKey($UninstallKey) 
#Retrieve an array of string that contain all the subkey names
$subkeys=$regkey.GetSubKeyNames() 
#Open each Subkey and use GetValue Method to return the required values for each
    foreach($key in $subkeys){
        $thisKey=$UninstallKey+”\\”+$key 
        $thisSubKey=$reg.OpenSubKey($thisKey)

        if (-not $thisSubKey.getValue("DisplayName")) { continue }
        if ($thisSubKey.getValue("SystemComponent")) { continue }
        if ($thisSubKey.getValue("ParentDisplayName")) { continue }
        if (-not $thisSubKey.getValue("UninstallString") -and -not $thisSubKey.getValue("WindowsInstaller")) { continue }
         
        $obj = New-Object PSObject
        $obj | Add-Member -MemberType NoteProperty -Name “DisplayName” -Value $($thisSubKey.GetValue(“DisplayName”))
        $obj | Add-Member -MemberType NoteProperty -Name “Comments” -Value $($thisSubKey.GetValue(“Comments”))
        $obj | Add-Member -MemberType NoteProperty -Name “DisplayVersion” -Value $($thisSubKey.GetValue(“DisplayVersion”))
        $obj | Add-Member -MemberType NoteProperty -Name “Publisher” -Value $($thisSubKey.GetValue(“Publisher”))

        $array += $obj

    } 

$array | Where-Object { $_.DisplayName } | select DisplayName, DisplayVersion, Publisher, Comments
}
# *****************************************************************************

# *****************************************************************************
# Function: create XML, JSON and HTML with tree-structure/contents of folder and file attributes
# Version: 1.3
# Author: Maslukivskiy Oleksandr, EVRY
# Functions: BLookSize, BLookDate, Add-Tabstops, RecExportHTML

Function BLookDate
{
	PARAM ($date)
	$newDate = "{0:dd-MM-yyyy} {0:HH:mm}" -f $date
	return $newDate
}

Function BLookSize
{
	PARAM ($size)
	if($size -le 0.999Kb){$newSize = '{0} bytes' -f $size}
	elseif($size -le 0.999Mb){$newSize = '{0:N0} Kb' -f ($size / 1Kb)}
	elseif($size -le 0.999Gb){$newSize = '{0:N2} MB' -f ($size / 1Mb)}
	else { $newSize = '{0:N2} Gb' -f $size / 1Gb }
	return $newSize
}

Function Add-Tabstops
{
	PARAM($Count)
	$tabs = ""
	for($i=0; $i -lt $Count; $i++){$tabs += "  "}
	return $tabs
}

Function RecExportHTML
{
	PARAM (	[String]$targetPath,
			[String]$outputFile)
	
	$style = 	"<style>" + 
					"`nbody { font-family: Arial;}" + 
					"`nul.tree li { list-style-type: none; position: relative; color: black}" +  
					"`nul.tree li ul { display: none;}" +
					"`nul.tree li.open > ul { display: block;}" +
					"`nul.tree li a { color: orange; text-decoration: none;}" +
					"`nul.tree li a:before { height: 1em;  padding:0.1em;  font-size: .8em;  display: block;  position: absolute;  left: -1em;  top: .1em;}" +
					"`nul.tree li > a:not(:last-child):before {  content: '+';}" +
					"`nul.tree li.open > a:not(:last-child):before {  content: '-';}" +
					"`nul.p li {list-style-type: none; color: blue; font-size:small; }" +	#file property style
				"`n</style>"
	
	$script = 	"<script type=`"text/javascript`">" + 
					"`nvar tree = document.querySelectorAll('ul.tree a:not(:last-child)');" + 
				"`nfor(var i = 0; i < tree.length; i++){" + 
					"`ntree[i].addEventListener('click', function(e) {" + 
						"`nvar parent = e.target.parentElement;" + 
						"`nvar classList = parent.classList;" + 
						"`nif(classList.contains(`"open`")) {" + 
							"`nclassList.remove('open');" + 
							"`nvar opensubs = parent.querySelectorAll(':scope .open');" + 
							"`nfor(var i = 0; i < opensubs.length; i++){" + 
								"`nopensubs[i].classList.remove('open');" + 
							"`n}" +  
						"`n} else {classList.add('open');}" + 
					"`n});" + 
				"`n}" + 
				"`n</script>"
	
Function Rec-HtmlChildren
{	
		PARAM([string]$currentPath, $Level = 2)
		
		return $(Get-ChildItem -Path $currentPath | sort LastWriteTime -Descending | Where-Object{$_} | ForEach-Object{
			(Add-Tabstops $Level) +
		
			# File-------------------------------
			$(if(!$_.psiscontainer){"<li><a href=`"`#`" style=`"color:black;`">" +
			$($_.Name) + "</a>" + "`n" + (Add-Tabstops ($Level+1)) + "<ul class=`"p`">" +
			"`n" + (Add-Tabstops ($Level+2)) + "<li>Created: " + $(BLookDate($_.CreationTime)) + "</li>" + 
			"`n" + (Add-Tabstops ($Level+2)) + "<li>Modified: " + $(BLookDate($_.LastWriteTime)) + "</li>" +
			"`n" + (Add-Tabstops ($Level+2)) + "<li>Length: " + $(BLookSize($_.Length)) + "</li>" +
			"`n" + (Add-Tabstops ($Level+1)) + "</ul>`n" + (Add-Tabstops ($Level)) + "</li>" }) +
		
			# Folder-----------------------------
			$(if($_.psiscontainer){
			
				if(Test-Path "$($_.Fullname)\*") #check if folder is empty
				{
					"<li><a href=`"`#`">" +
					$($_.Name) + "</a>" + "`n" + (Add-Tabstops ($Level+1)) + "<ul>" +
					"`n" + (Rec-HtmlChildren -currentPath $_.FullName -Level ($Level+2)) + 
					"`n" + (Add-Tabstops ($Level+1)) + "</ul>`n" + (Add-Tabstops ($Level)) + "</li>"
				}
				else
				{
					"<li style=`"color:gray;`">" + $($_.Name) + "</li>"
				}
			})
		 

		}) -join "`n"
	}
	(
	"<html>" + "`n" +
	"<head>" + "`n" +
	$style + "`n" +
	"<title>$targetPath</title>" + "`n" +
	"</head>" + "`n" +
	"<body><h1>$targetPath</h1>" + "`n" +
	"<ul class=`"tree`">" + "`n" +
	(Rec-HtmlChildren -currentPath $targetPath) + "`n" +
	"</ul>" + "`n" +
	"</body>" + "`n" +
	$script + "`n" +
	"</html>"
	) | Set-Content -Path $outputFile -Encoding UTF8
}
# *****************************************************************************

# Clear console screen for ISE debugging
cls
# *****************************************************************************

# Get script start time
$script:StartTime = get-date
# *****************************************************************************

# Write header
Write-Feedback "Running script for SKALA SCCM troubleshooting" "Green"
Write-Feedback "Script version: 1.0.0.7" "Green"
Write-Feedback "" "Green"
Write-Feedback "" "Green"
Write-Feedback "" "Green"
Write-Feedback "" "Green"
Write-Feedback "" "Green"
Write-Feedback ("*     Script started at:         " + $script:StartTime) "Green"
    # Progress bar
    Write-Progress -Activity “Current activity:” -status “Collecting generic system information...” -percentComplete 1
# *****************************************************************************

# Increasing screen buffer for smaller Screen resolutions
$host.UI.RawUI.BufferSize = new-object System.Management.Automation.Host.Size(256,300)
# *****************************************************************************

# Check OS name
$OSCaption = (gwmi win32_operatingsystem).caption
Write-Feedback ("*          Current OS: " + $OSCaption) "Green"
# *****************************************************************************

# Check if OS is newer than Windows7 - otherwise use Legacy solutions in some cases
$NEWOS = [Environment]::OSVersion.Version -gt (new-object 'Version' 6,2)
    if ($NEWOS){
            Write-Feedback "*          OS is newer than Windows 7" "Green"
    }
    else {
            Write-Feedback "*          Warning: OS is Windows 7 or older - some Legacy functions will be used." "Gray"
    }
# *****************************************************************************

# Check OS architecture
$X64OS = (Get-WmiObject -Class Win32_ComputerSystem).SystemType -match ‘(x64)’
    if ($X64OS){
            Write-Feedback "*          OS bitness: x64" "Green"
            $OSBITNESS = "x64"
    }
    else {
            Write-Feedback "*          OS bitness: x86" "Green"
            $OSBITNESS = "x86"
    }
# *****************************************************************************

# Verify, which version of PowerShell is used - Legacy options will be used for <4.0
$PSVersion=(Get-Host).version
    if ($PSVersion -ge "4.0"){
           Write-Feedback ("*          PowerShell version: " + $PSVersion) "Green"
    }
    else {
           Write-Feedback ("*          Warning: PowerShell version: " + $PSVersion + " - some Legacy functions will be used.") "Gray"
    }
# *****************************************************************************

# Check OS locale
$OSLOCALE = (gwmi win32_operatingsystem).locale
# *****************************************************************************

# Verify if currently script running in elevated mode
# May be required to define whether some specific logs/contents will be possible to gather
function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}
 
if (!(Test-IsAdmin)){
    Write-Feedback "*          Warning: Running in NON-ELEVATED mode - some checks will be skipped!" "Gray"
}
else {
    Write-Feedback "*          Running in ELEVATED mode" "Green"
}
# *****************************************************************************


# Check Current USERNAME and give output
    Write-Feedback ("*          Currently running as user: " + $env:USERNAME) "Green"
# *****************************************************************************

# Create temporary folder for logs and data, delete if already present
    # Create variable for temporary folder
    $varTEMPFolder = ($env:TEMP + "\ClientLogs\" + $env:COMPUTERNAME)
    Write-Feedback ("*          Temporary folder for log files: " + $varTEMPFolder) "Green"

    # Verify path for Upload exists and not read-only
    if ($UploadPath -ne ""){
        $UploadPath = $UploadPath.Trimend('\')

        If(Test-path $UploadPath){
                If(Test-Write $UploadPath){
                    Write-Feedback ("*          Log storage: " + $UploadPath) "Green"
                }
                Else {
                    Write-Feedback ("*          Warning: Upload folder found, but Read/Only - switching to local storage") "Gray"
                    Write-Feedback ("*          Log storage: " + $env:PUBLIC + "\SKALA") "Green"
                }
        }
        Else {
            Write-Feedback ("*          Warning: Upload folder not found, switching to local storage") "Gray"
            Write-Feedback ("*          Log storage: " + $env:PUBLIC + "\SKALA") "Green"
        }
    }
    Else{
        Write-Feedback ("*          Log storage: " + $env:PUBLIC + "\SKALA") "Green"
    }
    
    Write-Feedback "*     ________________________________" "Green"
    Write-Feedback "" "Green"
    # *****************************************************************************
    
    # Remove temporary directory if present
        # Progress bar
        Write-Progress -Activity “Current activity:” -status “Creating temporary directory...” -percentComplete 2

    New-Item ($env:TEMP + "\ClientLogs") -type directory -force | Out-Null
    If(Test-path $varTEMPFolder) {Get-ChildItem -Path $varTEMPFolder -Recurse | Remove-Item -force -recurse}
    # Creating temporary directory
    Write-Feedback "*     Creating temporary directory..." "Green"
    New-Item $varTEMPFolder -type directory -force | Out-Null
    # Creating Public\SKALA directory
    $varPUBLICSKALAfolder = ($env:PUBLIC + "\SKALA")
    New-Item $varPUBLICSKALAfolder -type directory -force | Out-Null
# *****************************************************************************


# Collecting generic system information
    Write-Feedback "*     Collecting generic system information..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status “Collecting generic system information...” -percentComplete 3

    $varGenericInfofile = $varTEMPFolder + "\GenericSystemInformation.log" 
    $varSystemInfo = $varTEMPFolder + "\SystemInfo.log"
    
    # Free space
    Write-Output ("Free Space:") | out-file -filepath $varGenericInfofile -Append
    Try
    {
        Get-WmiObject win32_logicaldisk -ErrorAction Stop | ft -AutoSize deviceID, @{Label="Size(Gb)"; Expression={($_.size/1Gb) -as [int]}}, @{Label="Free Space (Gb)"; Expression={($_.freespace/1Gb) -as [int]}} | out-file -filepath $varGenericInfofile -Append
    }
    Catch
    {
        Write-Feedback "*          Warning: Failed to access logical drives" "Gray"
        Write-Output ("- n/a") | out-file -filepath $varGenericInfofile -Append
    }

    # Installed Printers
    Write-Output ("") | out-file -filepath $varGenericInfofile -Append
    Write-Output ("Installed Printers:") | out-file -filepath $varGenericInfofile -Append
    Try
    {
        Get-WMIObject -Class Win32_Printer -ErrorAction Stop | select Caption, Comment, Default, Local, Shared, DriverName, Status | fl | out-file -filepath $varGenericInfofile -Append
    }
    Catch
    {
        Write-Feedback "*          Warning: Failed to access installed printers" "Gray"
        Write-Output ("- n/a") | out-file -filepath $varGenericInfofile -Append
    }

    # .NETFramework 4 dump
    Write-Output ("") | out-file -filepath $varGenericInfofile -Append
    Write-Output (".NET Framework 4 registry dump:") | out-file -filepath $varGenericInfofile -Append
    Try
    {
        gci HKLM:SOFTWARE\Microsoft\.NETFramework\v4.0.30319 -recurse -ErrorAction Stop | select-object Name | out-file -filepath $varGenericInfofile -Append
    }
    Catch
    {
        Write-Feedback "*          Warning: Failed to access .NET Framework 4 registry" "Gray"
        Write-Output ("- n/a") | out-file -filepath $varGenericInfofile -Append
    }

    # InternetExplorer Registry dump
    Write-Output ("") | out-file -filepath $varGenericInfofile -Append
    Write-Output ("InternetExplorer Registry dump:") | out-file -filepath $varGenericInfofile -Append
    Try
    {
        If (Test-Path "HKLM:\SOFTWARE\Microsoft\Internet Explorer"){get-itemproperty -path HKLM:"SOFTWARE\Microsoft\Internet Explorer" -ErrorAction Stop | fl | out-file -filepath $varGenericInfofile -Append}
    }
    Catch
    {
        Write-Feedback "*          Warning: Failed to access InternetExplorer registry" "Gray"
        Write-Output ("- n/a") | out-file -filepath $varGenericInfofile -Append
    }

    # Compatibility Registry dump
    Write-Output ("") | out-file -filepath $varGenericInfofile -Append
    Write-Output ("Compatibility Registry dump:") | out-file -filepath $varGenericInfofile -Append
    Try
    {
        If (Test-Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"){get-itemproperty -path HKLM:"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" -ErrorAction Stop | ft -AutoSize | out-file -filepath $varGenericInfofile -Append}
        If (Test-Path "HKLM:\Software\Wow6432Node\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"){get-itemproperty -path HKLM:"Software\Wow6432Node\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers" -ErrorAction Stop | ft -AutoSize | out-file -filepath $varGenericInfofile -Append}
    }
    Catch
    {
        Write-Feedback "*          Warning: Failed to access Compatibility registry" "Gray"
        Write-Output ("- n/a") | out-file -filepath $varGenericInfofile -Append
    }

    # PreApproved ActiveX Registry dump
    Write-Output ("") | out-file -filepath $varGenericInfofile -Append
    Write-Output ("PreApproved ActiveX Registry dump:") | out-file -filepath $varGenericInfofile -Append
    Try
    {
        If (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Ext\PreApproved"){gci -path HKLM:"SOFTWARE\Microsoft\Windows\CurrentVersion\Ext\PreApproved" -ErrorAction Stop | ft -AutoSize | out-file -filepath $varGenericInfofile -Append}
        If (Test-Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Ext\PreApproved"){gci -path HKLM:"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Ext\PreApproved" -ErrorAction Stop | ft -AutoSize | out-file -filepath $varGenericInfofile -Append}
    }
    Catch
    {
        Write-Feedback "*          Warning: Failed to access PreApproved ActiveX registry" "Gray"
        Write-Output ("- n/a") | out-file -filepath $varGenericInfofile -Append
    }

        # SystemInfo
    systeminfo | out-file -filepath $varSystemInfo
# *****************************************************************************


# Collecting environment variables
    Write-Feedback "*     Collecting Environment Variables..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status “Collecting environment variables...” -percentComplete 4

    $varENVfile = $varTEMPFolder + "\EnvironmentVariables.log" 
    Get-ChildItem Env: | sort name | ft -Autosize -Wrap | out-file -filepath $varENVfile
# *****************************************************************************


# Collecting Services with their current state
    $varSERVICESfile = $varTEMPFolder + "\Services.log"
    Write-Feedback "*     Collecting Services with their current state..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status “Collecting Services with their current state...” -percentComplete 5
    Try
    {
        Get-WmiObject -Class Win32_Service -ErrorAction Stop | select displayname,name,state,startmode,pathname | sort displayname | ft -AutoSize -Wrap | out-file -filepath $varSERVICESfile
    }
    Catch
    {
        Write-Feedback "*          Warning: Failed to access services" "Gray"
    }
# *****************************************************************************


# Collecting running Processes...
    $varPROCfile = $varTEMPFolder + "\Processes.log"
    Write-Feedback "*     Collecting running Processes..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status “Collecting running Processes...” -percentComplete 6
    Try
    {
        GetProcList -ErrorAction Stop | sort ProcessName | ft -Autosize | out-file -filepath $varPROCfile
    }
    Catch
    {
        Write-Feedback "*          Warning: Failed to access processes" "Gray"
    }

        # Progress bar
        Write-Progress -Activity “Current activity:” -status “Collecting running Processes...” -percentComplete 10
# *****************************************************************************

    
# Collecting Execution Policies...
    Write-Feedback "*     Collecting Execution Policies..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status “Collecting Execution Policies...” -percentComplete 11

    $varEXECPOLICYfile = $varTEMPFolder + "\ExecutionPolicies.log" 
    Get-ExecutionPolicy -List | fl scope,executionpolicy | out-file -filepath $varEXECPOLICYfile
# *****************************************************************************

    
# Collecting GroupPolicies
    Write-Feedback "*     Collecting GroupPolicies..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting GroupPolicies...” -percentComplete 12

    $varGPRESULTfile = $varTEMPFolder + "\GPResult.html"
    gpresult /H $varGPRESULTfile | Out-Null
# *****************************************************************************


# Collecting last entries in EventViewer logs
    Write-Feedback "*     Collecting last entries in EventViewer logs:" "Green"
    $oldDate = $script:StartTime.adddays(-7)
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting EventViewer logs...” -percentComplete 15

    New-Item ($varTEMPFolder + "\EventViewer") -type directory -force | Out-Null
    Write-Feedback "*                    - Processing last 7 days of Application log..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Processing Application log...” -percentComplete 16
    $varEVENTWRAPPSfile = $varTEMPFolder + "\EventViewer\Application.csv"
    get-eventlog -logname "Application" | where{$_.timegenerated -gt $olddate} | select EventID, machinename, entrytype, message, source, timegenerated, username | export-csv -NoTypeInformation -Delimiter ';' $varEVENTWRAPPSfile
        
    Write-Feedback "*                    - Processing last 7 days of System log..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Processing System log...” -percentComplete 18
    $varEVENTWRSYSfile = $varTEMPFolder + "\EventViewer\System.csv"
    get-eventlog -logname "System" | where{$_.timegenerated -gt $olddate} | select EventID, machinename, entrytype, message, source, timegenerated, username | export-csv -NoTypeInformation -Delimiter ';' $varEVENTWRSYSfile

    $varEVENTWRAPPVADMfile = $varTEMPFolder + "\EventViewer\AppV-Admin.csv"
    $varEVENTWRAPPVOPRfile = $varTEMPFolder + "\EventViewer\AppV-Operational.csv"
    Try
    {
        get-WinEvent -logname "Microsoft-AppV-Client/Admin" -ErrorAction Stop | select-object -First 1000 | select ID, machinename, LevelDisplayName, Message, TimeCreated | export-csv -NoTypeInformation -Delimiter ';' $varEVENTWRAPPVADMfile
        Write-Feedback "*                    - Processing last 1000 entries in Admin AppV log..." "Green"
           # Progress bar
           Write-Progress -Activity “Current activity:” -status "Processing Admin AppV log...” -percentComplete 20
    }
    Catch
    {
        # Delete empty file
        If(Test-path $varEVENTWRAPPVADMfile) {Remove-item $varEVENTWRAPPVADMfile}
    }
        
    Try
    {
        get-WinEvent -logname "Microsoft-AppV-Client/Operational" -ErrorAction Stop | select-object -First 1000 | select ID, machinename, LevelDisplayName, Message, TimeCreated | export-csv -NoTypeInformation -Delimiter ';' $varEVENTWRAPPVOPRfile
        Write-Feedback "*                    - Processing last 1000 entries in Operational AppV log..." "Green"
           # Progress bar
           Write-Progress -Activity “Current activity:” -status "Processing Operational AppV log...” -percentComplete 20
    }
    Catch
    {
        # Delete empty file
        If(Test-path $varEVENTWRAPPVOPRfile) {Remove-item $varEVENTWRAPPVOPRfile}
    }    

# *****************************************************************************


# Collecting SCCM cache file details
    if (!(Test-IsAdmin)){
        Write-Feedback "*     Collecting SCCM cache file details..." "Green"
        Write-Feedback "*          Warning: Skipped, User is non-elevated Administrator" "Gray"
    }
    else {
        Write-Feedback "*     Collecting SCCM cache file details..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting SCCM cache file details...” -percentComplete 22

        $varGPRESULTfile = $varTEMPFolder + "\ccmcache.html"
        $varSCCMcachefolder = ($env:WINDIR + "\ccmcache")

        If(Test-path $varSCCMcachefolder){
        RecExportHTML $varSCCMcachefolder $varGPRESULTfile
        }
        Else {
            Write-Feedback "*          Warning: Failed to find SCCM cache folder - check skipped." "Gray"
        }
    }
# *****************************************************************************


# Collecting Public\SKALA log files
    Write-Feedback "*     Collecting Public\SKALA log files..." "Green"
    $varPUBLICSKALAtempfolder = $varTEMPFolder + "\SKALALogs"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting Public\SKALA log files...” -percentComplete 23

    New-Item $varPUBLICSKALAtempfolder -type directory -force | Out-Null
    $varPUBLICSKALAfolderLOG = ($env:PUBLIC + "\SKALA\*.log")
    $varPUBLICSKALAfolderTXT = ($env:PUBLIC + "\SKALA\*.txt")
   If(Test-path $varPUBLICSKALAfolderLOG){Copy-Item $varPUBLICSKALAfolderLOG $varPUBLICSKALAtempfolder -Recurse -Force | Out-Null}
   If(Test-path $varPUBLICSKALAfolderTXT){Copy-Item $varPUBLICSKALAfolderTXT $varPUBLICSKALAtempfolder -Recurse -Force | Out-Null}
# *****************************************************************************


# Collecting Oracle log files
    Write-Feedback "*     Collecting Oracle log files..." "Green"
    $varOracletempfolder = $varTEMPFolder + "\Oracle"
    $varOracleInventoryfolderX86 = ($env:systemdrive + "\Program Files (x86)\Oracle\Inventory")
    $varOracleInventoryfolder = ($env:programfiles + "\Oracle\Inventory")

        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting Oracle log files...” -percentComplete 24

    If(Test-path $varOracleInventoryfolderX86){
        New-Item $varOracletempfolder -type directory -force | Out-Null
        New-Item ($varOracletempfolder + "\Inventory(x86)") -type directory -force | Out-Null
    }

    If(Test-path $varOracleInventoryfolder){
        New-Item $varOracletempfolder -type directory -force | Out-Null
        New-Item ($varOracletempfolder + "\Inventory") -type directory -force | Out-Null
    }
    
   If(Test-path $varOracleInventoryfolderX86){Copy-Item ($varOracleInventoryfolderX86 + "\*") ($varOracletempfolder + "\Inventory(x86)") -Recurse -Force | Out-Null}
   If(Test-path $varOracleInventoryfolder){Copy-Item ($varOracleInventoryfolder + "\*") ($varOracletempfolder + "\Inventory") -Recurse -Force | Out-Null}
# *****************************************************************************


# Collecting Flash Player log files
    Write-Feedback "*     Collecting Flash Player log files..." "Green"
    $varFlashtempfolder = $varTEMPFolder + "\Macromed"
    $varFlashInventoryfolderX86 = ($env:windir + "\System32\Macromed\Flash")
    $varFlashInventoryfolder = ($env:windir + "\SysWow64\Macromed\Flash")
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting Flash Player log files...” -percentComplete 24

    If(Test-path $varFlashInventoryfolderX86){
        New-Item $varFlashtempfolder -type directory -force | Out-Null
        New-Item ($varFlashtempfolder + "\System32") -type directory -force | Out-Null
    }

    If(Test-path $varFlashInventoryfolder){
        New-Item $varFlashtempfolder -type directory -force | Out-Null
        New-Item ($varFlashtempfolder + "\SysWow64") -type directory -force | Out-Null
    }
    
   If(Test-path $varFlashInventoryfolderX86){Copy-Item ($varFlashInventoryfolderX86 + "\*.log") ($varFlashtempfolder + "\System32") -Recurse -Force | Out-Null}
   If(Test-path $varFlashInventoryfolder){Copy-Item ($varFlashInventoryfolder + "\*.log") ($varFlashtempfolder + "\SysWow64") -Recurse -Force | Out-Null}
# *****************************************************************************


# Collecting SAS log files
    Write-Feedback "*     Collecting SAS log files..." "Green"
    $varSAStempfolder = $varTEMPFolder + "\SAS"
    $varSASInventoryfolderX86 = ($env:windir + "\System32\config\systemprofile\AppData\Local\SAS")
    $varSASInventoryfolder = ($env:windir + "\SysWow64\config\systemprofile\AppData\Local\SAS")

        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting SAS log files...” -percentComplete 24

    If(Test-path $varSASInventoryfolderX86){
        New-Item $varSAStempfolder -type directory -force | Out-Null
        New-Item ($varSAStempfolder + "\System32-SAS") -type directory -force | Out-Null
    }

    If(Test-path $varSASInventoryfolder){
        New-Item $varSAStempfolder -type directory -force | Out-Null
        New-Item ($varSAStempfolder + "\SysWOW64-SAS") -type directory -force | Out-Null
    }
    
   If(Test-path $varSASInventoryfolderX86){Copy-Item ($varSASInventoryfolderX86 + "\*") ($varSAStempfolder + "\System32-SAS") -Recurse -Force | Out-Null}
   If(Test-path $varSASInventoryfolder){Copy-Item ($varSASInventoryfolder + "\*") ($varSAStempfolder + "\SysWOW64-SAS") -Recurse -Force | Out-Null}
# *****************************************************************************


# Collecting Windir\TEMP log files
    Write-Feedback "*     Collecting Windir\TEMP log files..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting Windir\TEMP log files...” -percentComplete 25

    $varWINTEMPfolder = ($env:WINDIR + "\TEMP")
    $varWINTEMPtempfolder = $varTEMPFolder + "\TEMP"

    If(Test-path $varWINTEMPfolder)
    {

        New-Item $varWINTEMPtempfolder -type directory -force | Out-Null
        $varWINTEMPfolderLOG = ($env:WINDIR + "\TEMP\*.log")
        $varWINTEMPfolderTXT = ($env:WINDIR + "\TEMP\*.txt")
        $varWINTEMPfolderINI = ($env:WINDIR + "\TEMP\*.ini")
        $varWINTEMPfolderCFG = ($env:WINDIR + "\TEMP\*.cfg")
        $varWINTEMPfolderHTM = ($env:WINDIR + "\TEMP\*.htm")
        $varWINTEMPfolderXML = ($env:WINDIR + "\TEMP\*.xml")
        $varWINTEMPfolderHTML = ($env:WINDIR + "\TEMP\*.html")
        If(Test-path $varWINTEMPfolder){Copy-Item $varWINTEMPfolderLOG $varWINTEMPtempfolder -Recurse -Force | Out-Null}
        If(Test-path $varWINTEMPfolder){Copy-Item $varWINTEMPfolderTXT $varWINTEMPtempfolder -Recurse -Force | Out-Null}
        If(Test-path $varWINTEMPfolder){Copy-Item $varWINTEMPfolderINI $varWINTEMPtempfolder -Recurse -Force | Out-Null}
        If(Test-path $varWINTEMPfolder){Copy-Item $varWINTEMPfolderCFG $varWINTEMPtempfolder -Recurse -Force | Out-Null}
        If(Test-path $varWINTEMPfolder){Copy-Item $varWINTEMPfolderHTM $varWINTEMPtempfolder -Recurse -Force | Out-Null}
        If(Test-path $varWINTEMPfolder){Copy-Item $varWINTEMPfolderXML $varWINTEMPtempfolder -Recurse -Force | Out-Null}
        If(Test-path $varWINTEMPfolder){Copy-Item $varWINTEMPfolderHTML $varWINTEMPtempfolder -Recurse -Force | Out-Null}
    }
    Else
    {
        Write-Feedback "*          Warning: Failed to access %TEMP%" "Gray"
    }

# *****************************************************************************


# Collecting SCCM log files
    Write-Feedback "*     Collecting SCCM log files..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting SCCM log files...” -percentComplete 27

    $varCCMtempfolder = $varTEMPFolder + "\SCCMLogs"
    New-Item $varCCMtempfolder -type directory -force | Out-Null
    $varCCMfolderLOG = ($env:WINDIR + "\CCM\Logs\*.log")

    If(Test-path $varCCMfolderLOG)
    {
        Copy-Item $varCCMfolderLOG $varCCMtempfolder -Recurse -Force | Out-Null
    }
    Else
    {
        Write-Feedback "*          Warning: Failed to access %WINDIR%\CCM\Logs" "Gray"
    }
# *****************************************************************************


# Collecting Registry branches
    Write-Feedback "*     Collecting Registry branches..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting Registry...” -percentComplete 30

    $varRegistrytempfolder = $varTEMPFolder + "\Registry"
    New-Item $varRegistrytempfolder -type directory -force | Out-Null
    
    if ($X64OS){
        $varRegistryUninstall = $varRegistrytempfolder + "\Uninstallx86.reg"
        If (Test-Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"){REG EXPORT HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall $varRegistryUninstall /y | Out-Null}
        $varRegistryActiveSetupX86 = $varRegistrytempfolder + "\ActiveSetupx86.reg"
        If (Test-Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Active Setup"){REG EXPORT "HKLM\SOFTWARE\Wow6432Node\Microsoft\Active Setup" $varRegistryActiveSetupX86 /y | Out-Null}
        $varRegistryRunX86 = $varRegistrytempfolder + "\Runx86.reg"
        If (Test-Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run"){REG EXPORT HKLM\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Run $varRegistryRunX86 /y | Out-Null}
        $varRegistryODBCX86 = $varRegistrytempfolder + "\ODBCx86.reg"
        If (Test-Path "HKLM:\SOFTWARE\Wow6432Node\ODBC"){REG EXPORT HKLM\SOFTWARE\Wow6432Node\ODBC $varRegistryODBCX86 /y | Out-Null}

        $varRegistryJAVAX86 = $varRegistrytempfolder + "\JavaX86.reg"
        If (Test-Path "HKLM:\SOFTWARE\Wow6432Node\JavaSoft"){REG EXPORT HKLM\SOFTWARE\Wow6432Node\JavaSoft $varRegistryJAVAX86 /y | Out-Null}
        $varRegistryORACLEX86 = $varRegistrytempfolder + "\OracleX86.reg"
        If (Test-Path "HKLM:\SOFTWARE\Wow6432Node\ORACLE"){REG EXPORT HKLM\SOFTWARE\Wow6432Node\ORACLE $varRegistryORACLEX86 /y | Out-Null}

        $varRegistryCITRIXX86 = $varRegistrytempfolder + "\CitrixX86.reg"
        If (Test-Path "HKLM:\SOFTWARE\Wow6432Node\Citrix"){REG EXPORT HKLM\SOFTWARE\Wow6432Node\Citrix $varRegistryCITRIXX86 /y | Out-Null}
    }

    $varRegistryUninstall = $varRegistrytempfolder + "\Uninstall.reg"
    If (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"){REG EXPORT HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall $varRegistryUninstall /y | Out-Null}
    $varRegistryActiveSetupX64 = $varRegistrytempfolder + "\ActiveSetup.reg"
    If (Test-Path "HKLM:\SOFTWARE\Microsoft\Active Setup"){REG EXPORT "HKLM\SOFTWARE\Microsoft\Active Setup" $varRegistryActiveSetupX64 /y | Out-Null}
    $varRegistryRunX64 = $varRegistrytempfolder + "\Run.reg"
    If (Test-Path "HKLM:\SOFTWARE\ODBC"){REG EXPORT HKLM\SOFTWARE\ODBC $varRegistryRunX64 /y | Out-Null}
    $varRegistryODBC = $varRegistrytempfolder + "\ODBC.reg"
    If (Test-Path "HKLM:\SOFTWARE\ODBC"){REG EXPORT HKLM\SOFTWARE\ODBC $varRegistryODBC /y | Out-Null}
    $varRegistryEVRY = $varRegistrytempfolder + "\EVRY.reg"
    If (Test-Path "HKLM:\SOFTWARE\EVRY"){REG EXPORT HKLM\SOFTWARE\EVRY $varRegistryEVRY /y | Out-Null}

    $varRegistryJAVA = $varRegistrytempfolder + "\Java.reg"
    If (Test-Path "HKLM:\SOFTWARE\JavaSoft"){REG EXPORT HKLM\SOFTWARE\JavaSoft $varRegistryJAVA /y | Out-Null}
    $varRegistryORACLE = $varRegistrytempfolder + "\Oracle.reg"
    If (Test-Path "HKLM:\SOFTWARE\ORACLE"){REG EXPORT HKLM\SOFTWARE\ORACLE $varRegistryORACLE /y | Out-Null}

    $varRegistryCITRIX = $varRegistrytempfolder + "\Citrix.reg"
    If (Test-Path "HKLM:\SOFTWARE\Citrix"){REG EXPORT HKLM\SOFTWARE\Citrix $varRegistryCITRIX /y | Out-Null}

    $varRegistryHKCU = $varRegistrytempfolder + "\HKCU.reg"
    If (Test-Path "HKCU:\SOFTWARE"){REG EXPORT HKCU\SOFTWARE $varRegistryHKCU /y | Out-Null}
# *****************************************************************************


# Collecting installed software
    $varInstalledSF = $varTEMPFolder + "\ProgramsAndFeatures.log"

    Write-Feedback "*     Collecting installed software..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting installed software...” -percentComplete 32

    $arrayMerged=CheckUninstallRegKey(”SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall”)
    if ($X64OS){
    $arrayMerged+=CheckUninstallRegKey(”SOFTWARE\\wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall”)
    }

    $arrayMerged | sort DisplayName | ft -auto | out-file -filepath $varInstalledSF
# *****************************************************************************


# Collecting enabled Firewall openings
    $varFWtempfolder = $varTEMPFolder + "\Firewalls"
    New-Item $varFWtempfolder -type directory -force | Out-Null
    $varFWfolderIN = ($varFWtempfolder + "\InboundRules.log")
    $varFWfolderOUT = ($varFWtempfolder + "\OutboundRules.log")
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting enabled Firewall openings...” -percentComplete 34

    if ($NEWOS){
            Write-Feedback "*     Collecting enabled Firewall openings..." "Green"
            Show-NetFirewallRule | where {$_.enabled -eq ‘true’ -AND $_.direction -eq ‘inbound’} | select displayname,profile | sort displayname | ft -AutoSize -Wrap | out-file -filepath $varFWfolderIN
            Show-NetFirewallRule | where {$_.enabled -eq ‘true’ -AND $_.direction -eq ‘outbound’} | select displayname,profile | sort displayname | ft -AutoSize -Wrap | out-file -filepath $varFWfolderOUT
            # Progress bar
            Write-Progress -Activity “Current activity:” -status "Collecting enabled Firewall openings...” -percentComplete 36
    }
    else {
            Write-Feedback "*     Collecting enabled Firewall openings (Legacy)..." "Green"
            netsh advfirewall firewall show rule name=all dir=in | out-file -filepath $varFWfolderIN
            netsh advfirewall firewall show rule name=all dir=out | out-file -filepath $varFWfolderOUT
            # Progress bar
            Write-Progress -Activity “Current activity:” -status "Collecting enabled Firewall openings...” -percentComplete 36
    }
# *****************************************************************************


# Collecting Scheduled Tasks
    if (!(Test-IsAdmin)){
        Write-Feedback "*     Collecting Scheduled Tasks..." "Green"
        Write-Feedback "*          Warning: Skipped, User is non-elevated Administrator" "Gray"
    }
    else {
        Write-Feedback "*     Collecting Scheduled Tasks..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting Scheduled Tasks...” -percentComplete 37

        $varSCHEDTASKStempfolder = $varTEMPFolder + "\ScheduledTasks"
        New-Item $varSCHEDTASKStempfolder -type directory -force | Out-Null

        $sch = New-Object -ComObject("Schedule.Service")
        $sch.Connect("localhost")
        $tasks = $sch.GetFolder("\").GetTasks(0)

        $outfile_temp = $varSCHEDTASKStempfolder + "\{0}.xml"

        $tasks | %{
          $xml = $_.Xml
          $task_name = $_.Name
          $outfile = $outfile_temp -f $task_name
          $xml | Out-File $outfile
        }
    }
# *****************************************************************************


# Collecting installed Hotfixes
    $varUpdatesfile = $varTEMPFolder + "\InstalledUpdates.csv" 
    Write-Feedback "*     Collecting installed Hotfixes..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting installed Hotfixes...” -percentComplete 39

    # Gives a list of all Microsoft Updates sorted by KB number/HotfixID
    $wu = new-object -com "Microsoft.Update.Searcher"
    $totalupdates = $wu.GetTotalHistoryCount()

    If (!($totalupdates -eq "0")){

        $all = $wu.QueryHistory(0,$totalupdates)
        # Define a new array to gather output
         $OutputCollection=  @()
        Foreach ($update in $all)
            {
            $string = $update.title
            $Regex = "KB\d*"
            $KB = $string | Select-String -Pattern $regex | Select-Object { $_.Matches }
             $output = New-Object -TypeName PSobject
             $output | add-member NoteProperty "HotFixID" -value $KB.' $_.Matches '.Value
             $output | add-member NoteProperty "Title" -value $string
             $OutputCollection += $output
            }

        Write-Feedback ("*                    " + $($OutputCollection.Count) + " Updates Found") "Green"
        # Output the collection sorted and formatted:
        $OutputCollection | select HotFixID, Title | Sort-Object HotFixID | export-csv -NoTypeInformation -Delimiter ';' $varUpdatesfile

    }
    Else {
        Write-Feedback "*          Warning: Failed to retrieve list, using secondary mechanism (Windows Updates only)..." "Gray"

            if ($PSVersion -ge "4.0"){
                $Hotfixes = Get-HotFix | select CSName,InstalledOn,Description,HotfixID,InstalledBy
                $HotfixCount = ($Hotfixes | Measure-Object).count
                Write-Feedback ("*                    " + $HotfixCount + " Updates Found") "Gray"
            }
            else {
                Write-Feedback ("*          Utilizing Legacy approach - WMIC...") "Gray"
                $Hotfixes = wmic qfe list
            }

        # Output the collection sorted and formatted:
        $Hotfixes | export-csv -NoTypeInformation -Delimiter ';' $varUpdatesfile
    }
# *****************************************************************************


# Collecting enabled Windows components
    Write-Feedback "*     Collecting enabled Windows components..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Collecting enabled Windows components...” -percentComplete 42

    $varFeaturesfile = $varTEMPFolder + "\InstalledWindowsComponents.log" 
    $(foreach ($feature in Get-WmiObject -Class Win32_OptionalFeature -Namespace root\CIMV2 -Filter "InstallState = 1") {$feature.Caption}) | sort | out-file -filepath $varFeaturesfile
# *****************************************************************************


# Collecting etc\hosts and services files
    Write-Feedback "*     Collecting etc\hosts and services files..." "Green"
    # Progress bar
    Write-Progress -Activity “Current activity:” -status "Collecting etc\hosts...” -percentComplete 43

    $varETCfolder = ($env:WINDIR + "\System32\Drivers\etc")
    $varETCtempfolder = $varTEMPFolder + "\ETC"
    New-Item $varETCtempfolder -type directory -force | Out-Null
    $varETCtempfolderHOSTS = ($env:WINDIR + "\System32\Drivers\etc\hosts")
    $varETCtempfolderSERVICES = ($env:WINDIR + "\System32\Drivers\etc\services")
    Copy-Item $varETCtempfolderHOSTS $varETCtempfolder -Recurse -Force | Out-Null
    Copy-Item $varETCtempfolderSERVICES $varETCtempfolder -Recurse -Force | Out-Null        
# *****************************************************************************


# Compressing result files
    Write-Feedback "*     Compressing result files..." "Green"
        # Progress bar
        Write-Progress -Activity “Current activity:” -status "Compressing result files...” -percentComplete 44

    $FolderUncompressed = $varTEMPFolder
    $FolderCompressed = ($env:TEMP + "\" + "_" + $env:COMPUTERNAME + "_" + "SKALA_TroubleshootingLogs.zip")
    $FolderSKALA = ($env:Public + "\SKALA")
    $CompressedOutput = "0"

    # Add timestamp to output log file to define end of active checks
    $ActiveEndTime = get-date
    Add-Content $varOutputFile ("All checks finished at:" + $ActiveEndTime)

    # Copy output file
    Copy-Item $varOutputFile $varTEMPFolder -Recurse -Force | Out-Null

   Try
   {
       $FolderCompressed = ($env:TEMP + "\" + "_" + $env:COMPUTERNAME + "_" + "SKALA_TroubleshootingLogs.zip")
       If(Test-path $FolderCompressed) {Remove-item $FolderCompressed}
           Add-Type -assembly "system.io.compression.filesystem" -ErrorAction Stop
       [io.compression.zipfile]::CreateFromDirectory($FolderUncompressed, $FolderCompressed)
       If(Test-path $FolderCompressed) {$CompressedOutput = "1"}
       Write-Host "*          Compression completed successfully." -ForegroundColor Green
   }
   Catch
   {
       $FolderCompressed = ($env:TEMP + "\" + "_" + $env:COMPUTERNAME + "_" + "SKALA_TroubleshootingLogs")
       If(Test-path $FolderCompressed){
            $CompressedOutput = "0"
            Write-Host "*          Warning: Failed to load compression module, using uncompressed output" -ForegroundColor Gray
       }
       Else {
            $CompressedOutput = "0"
            Write-Host "*          Warning: Failed to load compression module, uncompressed output failed to copy!" -ForegroundColor Gray
       }
   }   
# *****************************************************************************


# Upload to Shared location if available and accessible
if ($UploadPath -eq ""){
    # Use default location to upload
    $FolderSKALA = ($env:Public + "\SKALA")
    $UploadPath = $FolderSKALA
}
Else
{
    # Use default location to upload
}

if ($UploadPath -eq "C:\Users\Public\SKALA"){
    # Use default location to upload
    $FolderSKALA = ($env:Public + "\SKALA")
    $UploadPath = $FolderSKALA
}
Else
{
    # Use default location to upload
}

# Upload path parameter was given, need to upload to other location
    Write-Host "*     Uploading result files..." -ForegroundColor Green
    # Progress bar
    Write-Progress -Activity "Current activity:" -status "Uploading result files..." -percentComplete 50

        If(Test-path $UploadPath) {
            # Define uploaded file
            If ($CompressedOutput -eq "1")
            {
                $Uploadedfile = ($UploadPath + "\_" + $env:COMPUTERNAME + "_" + "SKALA_TroubleshootingLogs.zip")
                # Delete if already present
                If(Test-Write $UploadPath){
                    If(Test-path $Uploadedfile) {Remove-item $Uploadedfile}
                }
            }
            Else
            {
                $Uploadedfile = ($UploadPath + "\_" + $env:COMPUTERNAME + "_" + "SKALA_TroubleshootingLogs")
                # Delete if already present
                If(Test-Write $UploadPath){
                    If(Test-path $Uploadedfile) {Remove-item $Uploadedfile -force -recurse}
                }
            }

            # Verify path is not Read Only
            If(Test-Write $UploadPath){
                    # Copy File to new location
                    If(Test-path $FolderCompressed) {Copy-Item $FolderCompressed $UploadPath -Force | Out-Null }

                    # Delete local copy of file
                    If (Test-Path $FolderCompressed) {Get-ChildItem -Path $FolderCompressed -Recurse | Remove-Item -force -recurse}
            }
            Else {
                    Write-Host "*          Warning: Upload path not accessible, uploaded to %TEMP%" -ForegroundColor Gray
                    $Uploadedfile = ($env:TEMP + "\" + "_" + $env:COMPUTERNAME + "_" + "SKALA_TroubleshootingLogs")
            }
        }
# *****************************************************************************


# Cleaning Temporary files
    Write-Host "*     Cleaning Temporary files..." -ForegroundColor Green
        # Progress bar
        Write-Progress -Activity "Current activity:" -status "Cleaning Temporary files..." -percentComplete 75

    $varCleanupFolder = ($env:TEMP + "\ClientLogs")
    If (Test-Path $varCleanupFolder) {Get-ChildItem -Path $varCleanupFolder -Recurse | Remove-Item -force -recurse}

    # Delete output file in Temporary location
    If(Test-path $varOutputFile) {Remove-item $varOutputFile -force -Recurse}
# *****************************************************************************


# Show result output and exit
        # Progress bar
        Write-Progress -Activity "Current activity:" -status "Finished..." -percentComplete 100

    # Footer
    Write-Host ""
    Write-Host "*     ________________________________" -ForegroundColor Green
    Write-Host ""

    # Measure time elapsed by script
    $currentEND = get-date
    Write-Host ("*     Script finished at:         " + $currentEND) -ForegroundColor Green
    $elapsedTime = ($(get-date) - $script:StartTime)
    $retStr = [string]::format("{0} sec(s)", [int]$elapsedTime.TotalSeconds)
    Write-Host ("*     Elapsed time: " + $retStr) -ForegroundColor Green
    
    # Footer
    If ($CompressedOutput -eq "0")
    {
        Write-Host ("*     Files copied to: " + $Uploadedfile) -ForegroundColor Green
    }
    Else
    {
        Write-Host ("*     Files compressed to: " + $Uploadedfile) -ForegroundColor Green
    }

    # Open containing folder if running in UI mode, skip if in Silent mode
    if (!$Silent){
        Write-Host "*     Opening explorer on resulting dir." -ForegroundColor Green
        Invoke-Expression "explorer '/select,$Uploadedfile'"
        pause
        exit
    }
    Else {
        cls
        exit
    }
# *****************************************************************************

# SIG # Begin signature block
# MIIdlwYJKoZIhvcNAQcCoIIdiDCCHYQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUoJIoIuHG8N0AxYFjTjjTLvBT
# CxygghecMIIFJTCCBA2gAwIBAgIQCs84AtqWS9bip3NdqvcSzzANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTE3MDIwOTAwMDAwMFoXDTIwMDIx
# MTEyMDAwMFowYjELMAkGA1UEBhMCTk8xETAPBgNVBAgTCEFrZXJzaHVzMRAwDgYD
# VQQHEwdGb3JuZWJ1MRYwFAYDVQQKEw1FVlJZIE5vcmdlIEFTMRYwFAYDVQQDEw1F
# VlJZIE5vcmdlIEFTMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAwHsy
# YqP/P/4J49fOP5Y9Cf08elxJzvWTjK+Pfm8op7HolDym0vK/cJNNVoduUkNu1MA3
# mgldYYAJFJBqk0lvFjzGWPTmmYkOY+MwkBrcZfsFN4zc2b8QiH+DdT5/fYnP4PU1
# rcfrecoUBpD/MAr7blMD0EqDggiqgACMg36yc10cPsKicQoguwXrq2NK5+nAPl+o
# lZoJg42frRU0159PKouu+YhTWKxEqny1ZdM0VCywd35YSxKv8MhfTtF8uSNrX9KZ
# +xmgC7Y44El3dVRxxrlhDPsQTf/c9M4EK8oUOy197pmZP6kxHxwAqC48N+dpWNWs
# xa21mow18NNgsCXL9wIDAQABo4IBxTCCAcEwHwYDVR0jBBgwFoAUWsS5eyoKo6Xq
# cQPAYPkt9mV1DlgwHQYDVR0OBBYEFGPLbFDc2Wbbz1uCXwoEWWuNmjEYMA4GA1Ud
# DwEB/wQEAwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzB3BgNVHR8EcDBuMDWgM6Ax
# hi9odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vc2hhMi1hc3N1cmVkLWNzLWcxLmNy
# bDA1oDOgMYYvaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL3NoYTItYXNzdXJlZC1j
# cy1nMS5jcmwwTAYDVR0gBEUwQzA3BglghkgBhv1sAwEwKjAoBggrBgEFBQcCARYc
# aHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzAIBgZngQwBBAEwgYQGCCsGAQUF
# BwEBBHgwdjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tME4G
# CCsGAQUFBzAChkJodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRT
# SEEyQXNzdXJlZElEQ29kZVNpZ25pbmdDQS5jcnQwDAYDVR0TAQH/BAIwADANBgkq
# hkiG9w0BAQsFAAOCAQEAdaBYuUrsZKg5+ZPf7Z0Ejo5tRlEKfGQuJqgXudCsaXFt
# L7jEZJIrqa+j9BVpqQ/csltu4NL5SVS36kFJGwtpV/NCTxK8HBSt0/R9+z72Eg5B
# Q8qZ9aUi/v+46L28pZ4UUOPbiRCEaWQmFFIfkcqZRYniOXMA3U3AolKcy9MeyAwJ
# BwHY4ge2tuy+ChO9G2jijhp4/fkjj+y02ou7+ouBGQBaWsatuMiyTqIqubGEwOnv
# e+Bnqd497xjwqgTyVsuGd1a88w7gnyNgVrSyIjZNXfQQYiyFEAK01Z/062ETxfp0
# JEzZQwkC+wm/uF1H4x7voJg9NCdzOs62M3uvYgi+CjCCBTAwggQYoAMCAQICEAQJ
# GBtf1btmdVNDtW+VUAgwDQYJKoZIhvcNAQELBQAwZTELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEk
# MCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4XDTEzMTAyMjEy
# MDAwMFoXDTI4MTAyMjEyMDAwMFowcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERp
# Z2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMo
# RGlnaUNlcnQgU0hBMiBBc3N1cmVkIElEIENvZGUgU2lnbmluZyBDQTCCASIwDQYJ
# KoZIhvcNAQEBBQADggEPADCCAQoCggEBAPjTsxx/DhGvZ3cH0wsxSRnP0PtFmbE6
# 20T1f+Wondsy13Hqdp0FLreP+pJDwKX5idQ3Gde2qvCchqXYJawOeSg6funRZ9PG
# +yknx9N7I5TkkSOWkHeC+aGEI2YSVDNQdLEoJrskacLCUvIUZ4qJRdQtoaPpiCwg
# la4cSocI3wz14k1gGL6qxLKucDFmM3E+rHCiq85/6XzLkqHlOzEcz+ryCuRXu0q1
# 6XTmK/5sy350OTYNkO/ktU6kqepqCquE86xnTrXE94zRICUj6whkPlKWwfIPEvTF
# jg/BougsUfdzvL2FsWKDc0GCB+Q4i2pzINAPZHM8np+mM6n9Gd8lk9ECAwEAAaOC
# Ac0wggHJMBIGA1UdEwEB/wQIMAYBAf8CAQAwDgYDVR0PAQH/BAQDAgGGMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0
# cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNV
# HR8EejB4MDqgOKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRB
# c3N1cmVkSURSb290Q0EuY3JsMDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsME8GA1UdIARIMEYwOAYKYIZI
# AYb9bAACBDAqMCgGCCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20v
# Q1BTMAoGCGCGSAGG/WwDMB0GA1UdDgQWBBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAf
# BgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQsFAAOC
# AQEAPuwNWiSz8yLRFcgsfCUpdqgdXRwtOhrE7zBh134LYP3DPQ/Er4v97yrfIFU3
# sOH20ZJ1D1G0bqWOWuJeJIFOEKTuP3GOYw4TS63XX0R58zYUBor3nEZOXP+QsRsH
# DpEV+7qvtVHCjSSuJMbHJyqhKSgaOnEoAjwukaPAJRHinBRHoXpoaK+bp1wgXNlx
# sQyPu6j4xRJon89Ay0BEpRPw5mQMJQhCMrI2iiQC/i9yfhzXSUWW6Fkd6fp0ZGuy
# 62ZD2rOwjNXpDd32ASDOmTFjPQgaGLOBm0/GkxAG/AeB+ova+YJJ92JuoVP6EpQY
# hS6SkepobEQysmah5xikmmRR7zCCBmowggVSoAMCAQICEAMBmgI6/1ixa9bV6uYX
# 8GYwDQYJKoZIhvcNAQEFBQAwYjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lD
# ZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8GA1UEAxMYRGln
# aUNlcnQgQXNzdXJlZCBJRCBDQS0xMB4XDTE0MTAyMjAwMDAwMFoXDTI0MTAyMjAw
# MDAwMFowRzELMAkGA1UEBhMCVVMxETAPBgNVBAoTCERpZ2lDZXJ0MSUwIwYDVQQD
# ExxEaWdpQ2VydCBUaW1lc3RhbXAgUmVzcG9uZGVyMIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEAo2Rd/Hyz4II14OD2xirmSXU7zG7gU6mfH2RZ5nxrf2uM
# nVX4kuOe1VpjWwJJUNmDzm9m7t3LhelfpfnUh3SIRDsZyeX1kZ/GFDmsJOqoSyyR
# icxeKPRktlC39RKzc5YKZ6O+YZ+u8/0SeHUOplsU/UUjjoZEVX0YhgWMVYd5SEb3
# yg6Np95OX+Koti1ZAmGIYXIYaLm4fO7m5zQvMXeBMB+7NgGN7yfj95rwTDFkjePr
# +hmHqH7P7IwMNlt6wXq4eMfJBi5GEMiN6ARg27xzdPpO2P6qQPGyznBGg+naQKFZ
# OtkVCVeZVjCT88lhzNAIzGvsYkKRrALA76TwiRGPdwIDAQABo4IDNTCCAzEwDgYD
# VR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwggG/BgNVHSAEggG2MIIBsjCCAaEGCWCGSAGG/WwHATCCAZIwKAYIKwYBBQUH
# AgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwggFkBggrBgEFBQcCAjCC
# AVYeggFSAEEAbgB5ACAAdQBzAGUAIABvAGYAIAB0AGgAaQBzACAAQwBlAHIAdABp
# AGYAaQBjAGEAdABlACAAYwBvAG4AcwB0AGkAdAB1AHQAZQBzACAAYQBjAGMAZQBw
# AHQAYQBuAGMAZQAgAG8AZgAgAHQAaABlACAARABpAGcAaQBDAGUAcgB0ACAAQwBQ
# AC8AQwBQAFMAIABhAG4AZAAgAHQAaABlACAAUgBlAGwAeQBpAG4AZwAgAFAAYQBy
# AHQAeQAgAEEAZwByAGUAZQBtAGUAbgB0ACAAdwBoAGkAYwBoACAAbABpAG0AaQB0
# ACAAbABpAGEAYgBpAGwAaQB0AHkAIABhAG4AZAAgAGEAcgBlACAAaQBuAGMAbwBy
# AHAAbwByAGEAdABlAGQAIABoAGUAcgBlAGkAbgAgAGIAeQAgAHIAZQBmAGUAcgBl
# AG4AYwBlAC4wCwYJYIZIAYb9bAMVMB8GA1UdIwQYMBaAFBUAEisTmLKZB+0e36K+
# Vw0rZwLNMB0GA1UdDgQWBBRhWk0ktkkynUoqeRqDS/QeicHKfTB9BgNVHR8EdjB0
# MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVk
# SURDQS0xLmNybDA4oDagNIYyaHR0cDovL2NybDQuZGlnaWNlcnQuY29tL0RpZ2lD
# ZXJ0QXNzdXJlZElEQ0EtMS5jcmwwdwYIKwYBBQUHAQEEazBpMCQGCCsGAQUFBzAB
# hhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUHMAKGNWh0dHA6Ly9j
# YWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRENBLTEuY3J0MA0G
# CSqGSIb3DQEBBQUAA4IBAQCdJX4bM02yJoFcm4bOIyAPgIfliP//sdRqLDHtOhcZ
# cRfNqRu8WhY5AJ3jbITkWkD73gYBjDf6m7GdJH7+IKRXrVu3mrBgJuppVyFdNC8f
# cbCDlBkFazWQEKB7l8f2P+fiEUGmvWLZ8Cc9OB0obzpSCfDscGLTYkuw4HOmksDT
# jjHYL+NtFxMG7uQDthSr849Dp3GdId0UyhVdkkHa+Q+B0Zl0DSbEDn8btfWg8cZ3
# BigV6diT5VUW8LsKqxzbXEgnZsijiwoc5ZXarsQuWaBh3drzbaJh6YoLbewSGL33
# VVRAA5Ira8JRwgpIr7DUbuD0FAo6G+OPPcqvao173NhEMIIGzTCCBbWgAwIBAgIQ
# Bv35A5YDreoACus/J7u6GzANBgkqhkiG9w0BAQUFADBlMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29t
# MSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMDYxMTEw
# MDAwMDAwWhcNMjExMTEwMDAwMDAwWjBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMM
# RGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQD
# ExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwggEiMA0GCSqGSIb3DQEBAQUAA4IB
# DwAwggEKAoIBAQDogi2Z+crCQpWlgHNAcNKeVlRcqcTSQQaPyTP8TUWRXIGf7Syc
# +BZZ3561JBXCmLm0d0ncicQK2q/LXmvtrbBxMevPOkAMRk2T7It6NggDqww0/hhJ
# gv7HxzFIgHweog+SDlDJxofrNj/YMMP/pvf7os1vcyP+rFYFkPAyIRaJxnCI+QWX
# faPHQ90C6Ds97bFBo+0/vtuVSMTuHrPyvAwrmdDGXRJCgeGDboJzPyZLFJCuWWYK
# xI2+0s4Grq2Eb0iEm09AufFM8q+Y+/bOQF1c9qjxL6/siSLyaxhlscFzrdfx2M8e
# CnRcQrhofrfVdwonVnwPYqQ/MhRglf0HBKIJAgMBAAGjggN6MIIDdjAOBgNVHQ8B
# Af8EBAMCAYYwOwYDVR0lBDQwMgYIKwYBBQUHAwEGCCsGAQUFBwMCBggrBgEFBQcD
# AwYIKwYBBQUHAwQGCCsGAQUFBwMIMIIB0gYDVR0gBIIByTCCAcUwggG0BgpghkgB
# hv1sAAEEMIIBpDA6BggrBgEFBQcCARYuaHR0cDovL3d3dy5kaWdpY2VydC5jb20v
# c3NsLWNwcy1yZXBvc2l0b3J5Lmh0bTCCAWQGCCsGAQUFBwICMIIBVh6CAVIAQQBu
# AHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMAIABDAGUAcgB0AGkAZgBpAGMAYQB0
# AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMAIABhAGMAYwBlAHAAdABhAG4AYwBl
# ACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMAZQByAHQAIABDAFAALwBDAFAAUwAg
# AGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkAbgBnACAAUABhAHIAdAB5ACAAQQBn
# AHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgAIABsAGkAbQBpAHQAIABsAGkAYQBi
# AGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUAIABpAG4AYwBvAHIAcABvAHIAYQB0
# AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAAcgBlAGYAZQByAGUAbgBjAGUALjAL
# BglghkgBhv1sAxUwEgYDVR0TAQH/BAgwBgEB/wIBADB5BggrBgEFBQcBAQRtMGsw
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcw
# AoY3aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElE
# Um9vdENBLmNydDCBgQYDVR0fBHoweDA6oDigNoY0aHR0cDovL2NybDMuZGlnaWNl
# cnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDA6oDigNoY0aHR0cDov
# L2NybDQuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNybDAd
# BgNVHQ4EFgQUFQASKxOYspkH7R7for5XDStnAs0wHwYDVR0jBBgwFoAUReuir/SS
# y4IxLVGLp6chnfNtyA8wDQYJKoZIhvcNAQEFBQADggEBAEZQPsm3KCSnOB22Wymv
# Us9S6TFHq1Zce9UNC0Gz7+x1H3Q48rJcYaKclcNQ5IK5I9G6OoZyrTh4rHVdFxc0
# ckeFlFbR67s2hHfMJKXzBBlVqefj56tizfuLLZDCwNK1lL1eT7EF0g49GqkUW6aG
# MWKoqDPkmzmnxPXOHXh2lCVz5Cqrz5x2S+1fwksW5EtwTACJHvzFebxMElf+X+Ee
# vAJdqP77BzhPDcZdkbkPZ0XN1oPt55INjbFpjE/7WeAjD9KqrgB87pxCDs+R1ye3
# Fu4Pw718CqDuLAhVhSK46xgaTfwqIa1JMYNHlXdx3LEbS0scEJx3FMGdTy9alQgp
# ECYxggVlMIIFYQIBATCBhjByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNl
# cnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdp
# Q2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5nIENBAhAKzzgC2pZL1uKn
# c12q9xLPMAkGBSsOAwIaBQCgggGgMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEE
# MBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQZ
# 0hETHZnto9Vk9ZwK9TFdbzTN6DCCAT4GCisGAQQBgjcCAQwxggEuMIIBKqCCASaA
# ggEiAFUAUgBMAD0AaAB0AHQAcAA6AC8ALwBzAGMAYQBmAHMAMAAwADEALgBzAGMA
# YQAuAG8AcABlAHIALgBuAG8ALwBCAG8AbgBvAGIAbwAuAEcAaQB0AC4AUwBlAHIA
# dgBlAHIALwBTAEsAQQBMAEEALQBBAHUAdABvAEwAbwBnAC4AZwBpAHQAOwBIAEEA
# UwBIAD0ANQBlADkAMQAzAGQAZQA2AGQANAAzADEAOQA1ADcAOQAxAGQAMQBhAGEA
# NwA2ADgANgBiAGYAMgAzADMANABjAGQAYgAxADcANgA3ADAAYQA7AEcASQBUAF8A
# QwBPAE0ATQBJAFQAVABFAFIAXwBOAEEATQBFAD0ASQBnAG8AcgAgAFoAYQBoAGEA
# cgBvAHYwDQYJKoZIhvcNAQEBBQAEggEAmFP1ZvARdnMlqvpwHn6BbQHEF+p/WEvm
# rye+FBY5DIsv5iLsPRQhgUUH0z3AYw+RfRGwEO2Lf+MddvAFTOLWfpG8w4KQau2v
# cXdOaeruVsPdBIHkNsbmo62X4gHrODggZtvc880yAv/KnkkguLbH6MaJNgtalFsL
# tdhyPcCwGqWUggfm7LQwujreskwLEd7ecAcIxRF2am3OD9XF5asDaMQvdpgHIoPA
# LA9O4Kfzm4wOthwcKB2+RJy/IHrFmjIVyAhbmbYhbBiQxKQXzvm2ok2G81hagvB8
# tXEc1VtMT6KQMUqY4b+aw+1XSPdo5lfpvOt1U2WP9JKFs4mEpyTBBaGCAg8wggIL
# BgkqhkiG9w0BCQYxggH8MIIB+AIBATB2MGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNV
# BAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMQIQAwGaAjr/WLFr1tXq5hfwZjAJ
# BgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0B
# CQUxDxcNMTcwMzMwMTMxMTE1WjAjBgkqhkiG9w0BCQQxFgQUZYjdlLIaJCW5YUJq
# 2YsSowbwotEwDQYJKoZIhvcNAQEBBQAEggEAbJJqclnLKRdJBN3gSr1OlgV7MKRB
# k+eC7gZ+9KW/uixzWy4JWTtVvC+xKESdtMeK7wVDncHnrq3pehx/MnIGDGX+g+8e
# cowqr4fv8SW2I1YfM5YKA3BfYnlP3EHVj38bLYi5/nbZJgKgxbfFBc74go/GInxm
# e3ufI2s7Q8RyrZwukZcHatXtG3p2YykWDg5zdFgq+/ee/xg7CKsx7r7T+8GmHTeh
# kskOhnj3k0VegkpKVT/WQyohLoCKRKQWWju/TNsnzM/eAB/17PtdpgFvwi0j2rn7
# FJbZ1pmhdGkQ7pA+eL865Z6NBoK6BdlQsS7HvAmL++wd9zcb6S0m3nXw2g==
# SIG # End signature block
