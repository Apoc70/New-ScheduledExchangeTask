# New-ScheduledExchangeTask.ps1
Add a new scheduled task for Exchange Server 2013 scripts

## Description
This script adds a new scheduled task for an Exchange Server 2013 environment in a new task scheduler group "Exchange".

Providing a username and password the scheduled task will be configured to "Run whether user is logged on or not"

When username and password are provided the Register-ScheduledTask cmdlet verfies the logon credentials and will fails, if the credentials provided (username/password) are not valid.

The cmdlet Register-ScheduledTask consumes the user password in clear text

## Parameters
### TaskName
Name of the scheduled task. 

### ScriptName  
Script filename to be executed by task scheduler without filepath

### ScriptPath
Filepath to the PowerShell script to be executed

### GroupName
Groupname for task scheduler grouping. Default 'Exchange'   

### Description
The description of the scheduled task. If empty description defaults to "Execute script SCRIPTNAME"

### TaskUser
Username to be set as task user. Format either DOMAIN\USER or USER@DOMAIN   

### Password
Password for TaskUser. 
If not provided, the task will be automatically be created as "Run only when user is logged on"
If provided, the task will automatically be created as "Run whether the user is logged on or not"

## Examples
```
.\New-ScheduledExchangeTask.ps1 -TaskName "My Task" -ScriptName TaskScript1.ps1 -ScriptPath D:\Automation -TaskUser DOMAIN\ServiceAccount -Password P@ssw0rd
```

```
.\New-ScheduledExchangeTask.ps1 -TaskName "My Task" -ScriptName TaskScript1.ps1 -ScriptPath D:\Automation 
```
## Note
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## TechNet Gallery
Find the script at TechNet Gallery
* https://gallery.technet.microsoft.com/New-ScheduledExchangeTask-449cb182

## Credits
Written by: Thomas Stensitzki

## Social

* My Blog: http://justcantgetenough.granikos.eu
* Archived Blog: http://www.sf-tools.net/
* Twitter: https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github: https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog: http://blog.granikos.eu/
* Website: https://www.granikos.eu/en/
* Twitter: https://twitter.com/granikos_de

## Additional Credits:
* Thanks to Michel de Rooij (https://eightwone.com) for some PowerShell inspiration
* Thanks to Ed Wilson (https://blogs.technet.microsoft.com/heyscriptingguy) for some more PowerShell inspiration