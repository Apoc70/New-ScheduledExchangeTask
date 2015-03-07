<# 
    .SYNOPSIS 
    Add a new scheduled task for Exchange Server 2013 scripts

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.0, 2015-03-05

    Thanks to Michel de Rooij (michel@eightwone.com) for some PowerShell inspiration
    Thanks to Ed Wilson (Scripting Guy) for some more PowerShell inspiration

    Some code for handling scheduled tasks has been taken from
    http://blogs.technet.com/b/heyscriptingguy/archive/2015/01/16/use-powershell-to-create-scheduled-task-in-new-folder.aspx 

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 
    More information can be found at http://www.granikos.eu/en/scripts
    
    Learn more about *-ScheduledTask cmdlets at 

    .DESCRIPTION 
    This script adds a new scheduled task for an Exchange Server 2013 environment in a 
    new task scheduler group "Exchange".

    Providing a username and password the scheduled task will be configured to "Run whether user is logged on or not"
    
    When username and password are provided the Register-ScheduledTask cmdlet verfies the logon credentials 
    and will fails, if the credentials provided (username/password) are not valid.

    The cmdlet Register-ScheduledTask consumes the user password in clear text.
     
    .NOTES 
    Requirements 
    - Windows Server 2012 R2  
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 

    .PARAMETER TaskName
    Name of the scheduled task. 

    .PARAMETER ScriptName  
    Script filename to be executed by task scheduler without filepath

    .PARAMETER ScriptPath
    Filepath to the PowerShell script to be executed

    .PARAMETER GroupName
    Groupname for task scheduler grouping. Default 'Exchange'   

    .PARAMETER Description
    The description of the scheduled task. If empty description defaults to "Execute script SCRIPTNAME"

    .PARAMETER TaskUser
    Username to be set as task user. Format either DOMAIN\USER or USER@DOMAIN   

    .PARAMETER Password
    Password for TaskUser. 
    If not provided, the task will be automatically be created as "Run only when user is logged on"
    If provided, the task will automaticall be created as "Run whether the user is logged on or not"
    
    .EXAMPLE 
    .\New-ScheduledExchangeTask.ps1 -TaskName "My Task" -ScriptName TaskScript1.ps1 -ScriptPath D:\Automation -TaskUser DOMAIN\ServiceAccount -Password P@ssw0rd

    .EXAMPLE
    .\New-ScheduledExchangeTask.ps1 -TaskName "My Task" -ScriptName TaskScript1.ps1 -ScriptPath D:\Automation 
#>

param(
	[parameter(Mandatory=$true, HelpMessage='Task name for the scheduled task')]
		[string] $TaskName,
	[parameter(Mandatory=$true, HelpMessage='Task scheduler script name to execute (i.e. Run-ExchangeReport.ps1)')]
		[string] $ScriptName,
	[parameter(Mandatory=$true, HelpMessage='File path for the script to execute (i.e. D:\ScriptAutomation)')]
		[string] $ScriptPath,
    [parameter(Mandatory=$true, HelpMessage='Task scheduler group name. Default: Exchange')]
		[string] $GroupName = "Exchange",
    [parameter(Mandatory=$false, HelpMessage='Task description')]
		[string] $Description,
    [parameter(Mandatory=$false, HelpMessage='Domain user to be set as task user (DOMAIN\USER)')]
		[string] $TaskUser,
    [parameter(Mandatory=$false, HelpMessage='Password for task user')]
		[string] $Password
)

Set-StrictMode -Version Latest 

## Some Variables #########################
$ERR_OK = 0
$ERR_OSNOTSUPPPORTED = 1002
$ERR_EXCHANGESCRIPTNOTPRESENT = 1101
$ERR_EXCHANGEENVIRONMENTVARIABLENOTPRESENT = 1102

$exchangeRemoteScript = "RemoteExchange.ps1"

function CheckOSVersion {
    If( ($MajorOSVersion -ne "6.3") ) { 
        Write-Error "Windows Server 2012 or Windows Server 2012 R2 is required but not detected" 
        Exit $ERR_OSNOTSUPPPORTED
    }
    else {
        return $true
    }
}

function CheckExchange {
    try
    {
        if($env:ExchangeInstallPath -ne "") {
            if(-Not (Test-Path (Join-Path (Join-Path $env:ExchangeInstallPath "bin") $exchangeRemoteScript ))) {
                Exit $ERR_EXCHANGESCRIPTNOTPRESENT
            }
        }
    }
    catch
    {
        Write-Error "Exchange Server environment variable not present. Check your Exchange Server setup."
        Exit $ERR_EXCHANGEENVIRONMENTVARIABLENOTPRESENT      
    }

    return $true
}

# Create new scheduled task folder
function New-ScheduledTaskFolder {
    Param ($ScheduledTaskPath)
    Write-Verbose "Checking scheduled task folder (group)"
    $ErrorActionPreference = "stop"
    $scheduleObject = New-Object -ComObject schedule.service
    $scheduleObject.connect()
    $rootFolder = $scheduleObject.GetFolder("\")
    Try {$null = $scheduleObject.GetFolder($ScheduledTaskPath)}
    Catch { $null = $rootFolder.CreateFolder($ScheduledTaskPath) }
    Finally { $ErrorActionPreference = "continue" } 
}

# Create and register new scheduled task
function Create-AndRegisterExchangeTask {
    Param ($ExchangeTaskName, $ExchangeScriptPath, $ExchangeScheduledTaskPath, $ExchangeTaskDescription, $ExchangeTaskScript)

    $exchangeScriptPath = (Join-Path (Join-Path $env:ExchangeInstallPath "bin") $exchangeRemoteScript )

    # Build task argument, Run PowerShell window in hidden mode, load RemoteExchange.ps1 script, connect to Exchange and execute script
    $taskArgument = "-version 3.0 -NonInteractive -NoProfile -WindowsStyle Hidden -command "". '$($ExchangeScriptPath)'; Connect-ExchangeServer -auto; $($ExchangeTaskScript)"

    # Create task action
    $taskAction = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument $taskArgument

    # To do: get trigger config exposed
    $taskTrigger =  New-ScheduledTaskTrigger -Weekly -At 6am -DaysOfWeek Monday

    if($ExchangeTaskDescription -eq "") {
        $ExchangeTaskDescription = "Execute script $($ExchangeTaskScript)"
        Write-Verbose "No description provided, setting ExchangeTaskDescription to: $($ExchangeTaskDescription)"
    }

    Write-Verbose "Registering task"

    Register-ScheduledTask -Action $taskAction -Trigger $taskTrigger -TaskName $ExchangeTaskName -Description $ExchangeTaskDescription -TaskPath $ExchangeScheduledTaskPath -RunLevel Highest
}

# Add scheduled task configuration
function Add-ExchangeTaskSettings {
    Param ($ExchangeTaskName, $ExchangeScheduledTaskPath)
    
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Hidden 
    
    Set-ScheduledTask -TaskName $ExchangeTaskName -Settings $settings -TaskPath $ExchangeScheduledTaskPath
}

# Configure task user
function Add-ExchangeTaskUser {
    Param ($ExchangeTaskName, $ExchangeScheduledTaskPath, $ExchangeTaskUser, $ExchangeTaskPassword)
    Write-Verbose "Configuring task user"

    Set-ScheduledTask -TaskName $ExchangeTaskName -TaskPath $ExchangeScheduledTaskPath -User $ExchangeTaskUser -Password $ExchangeTaskPassword
}


## Main ###################################

$MajorOSVersion= [string](Get-WmiObject Win32_OperatingSystem | Select Version | Select @{n="Major";e={($_.Version.Split(".")[0]+"."+$_.Version.Split(".")[1])}}).Major
$MinorOSVersion= [string](Get-WmiObject Win32_OperatingSystem | Select Version | Select @{n="Minor";e={($_.Version.Split(".")[2])}}).Minor

if (CheckOSVersion -and CheckExchange) {

    if(Get-ScheduledTask -TaskName $TaskName -EA 0) {
        Write-Output "Task $($TaskName) exists. Task will be unregistered now."
        Unregister-ScheduledTask -TaskName $taskname -Confirm:$false
    }

    Write-Output "Creating new Exchange Scheduled Task: $($TaskName)"

    # Create a new scheduled task path (Task Scheduler UI calls it groups)
    New-ScheduledTaskFolder -ScheduledTaskPath $GroupName

    # Build script file path
    $taskScriptPath = Join-Path $ScriptPath $ScriptName

    # Build task scheduler group name
    $taskPath = "\" + $GroupName + "\"

    # Path to PowerShell Executable
    # $taExecute = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

    # Build 
    $exchangeScriptPath = (Join-Path (Join-Path $env:ExchangeInstallPath "bin") $exchangeRemoteScript )

    # Create and register Exchange Task
    Create-AndRegisterExchangeTask -ExchangeTaskName $TaskName -ExchangeScriptPath $exchangeScriptPath -ExchangeScheduledTaskPath $taskPath -ExchangeTaskDescription $Description -ExchangeTaskScript $taskScriptPath | Out-Null

    # Set Exchange Task settings
    Add-ExchangeTaskSettings -ExchangeTaskName $TaskName -ExchangeScheduledTaskPath $taskPath | Out-Null

    if(($TaskUser -ne "") -and ($Password -ne "")) {
        # Set task user and password to run task whether the user is logged on or not
        Add-ExchangeTaskUser -ExchangeTaskName $TaskName -ExchangeScheduledTaskPath $taskPath -ExchangeTaskUser $TaskUser -ExchangeTaskPassword $Password | Out-Null
    }

    Write-Output "Task $($TaskName) created!"
}