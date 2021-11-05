#region: Clear old sessions from PS
Write-host -ForegroundColor Yellow  "Clearing previous session connections to cloud servcies"
$logins = get-azcontext
foreach ($account in $logins)
{
Disconnect-AzAccount -Username $account.account
}
#endregion

#region:Functions
#Configuring logging... its a Function == When needing a log entry simply use: Log-Entry "Log text here"

$LogPath = "C:\LogDump\"+[string](Get-Date -format yyyyMdd)+" New365Grp.txt"
Write-host -ForegroundColor Yellow  "Logs to be found at c:\logdump"

#log entry with timestamp
Function Log-Entry {
                $Log = [string](Get-Date)+": "+$args
                Add-Content -Path $LogPath -Value $Log
}

#log entry with without imestamp
Function Log-EntrySimple {
                Add-Content -Path $LogPath -Value $args
}

#endregion

Log-EntrySimple "______________________________________"
Log-EntrySimple "_________Starting Code Pass___________"
Log-EntrySimple "______________________________________"

#region:script setup

##connect to EoL - no credentials cached to comply with Modern Auth
Write-host -ForegroundColor Yellow  "Connecting to EoL"
log-entry "Connecting to EoL"
Connect-ExchangeOnline

#pop up button to add manual input and notify of the next action required
   
$Shell          = New-Object -ComObject "WScript.Shell"
$Button         = $Shell.Popup("Click OK to continue and selct the Excel workbook to import.", 0, "New-UnifiedGroup", 0)

#pop up window to navigate to the excel document required, filtered to Excel only

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser    = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter      = 'SpreadSheet (*.xlsx)|*.xlsx'
}
$null           = $FileBrowser.ShowDialog()

#import the worksheet selected in the action above

$GroupName      = Import-Excel -path $filebrowser.FileName -worksheetname New-UnifiedGroup
Write-host -ForegroundColor Yellow  "Excel worksheet imported"
Log-Entry "Excel worksheet imported"
#endregion

#region:add 365Group
#add 365Group

#loop testing - $Group = $GroupName[0] - create an array of a single item
foreach ($Group in $GroupName) {
    #variable to build the 365 groupname and trim whitespace, if any
    $Str            = ($group.GroupName).Trim()
    #variable to build check group test
    Write-host -ForegroundColor Yellow  "Checking if Group already exists"
    log-entry "Checking if Group already exists"
    $CheckGroup     = get-UnifiedGroup $Str -ErrorAction silentlycontinue
    
    #if something came back then value is true and skip. with no value create 365 group
    If($CheckGroup) {
       Write-host -ForegroundColor Yellow  "a 365 group with the distinguish name was already found in 365" $Group.GroupName
       Log-Entry "a 365 group with the distinguish name was already found in 365" $Group.GroupName
    } 
    Else { 
       Write-host -ForegroundColor Green "creating group" $Group.GroupName
       Log-Entry "creating group" $Group.GroupName
       New-UnifiedGroup  -DisplayName $Str 
       while (!$CheckGroup) 
       #while the Group check is not true sleep 2 seconds and then try again
       {
        Write-Host "Check 365 Group Not found - starting sleep"
        log-entry "Check 365 Group Not found - starting sleep"
        start-sleep -seconds 2
        $CheckGroup     = Get-UnifiedGroup $Str  
       } 
       #stop email notifications about group membership - this can be resversed aftr the migration/project
       Write-host -ForegroundColor Yellow "email notifications are supressed in  " $Str
       Log-Entry  "email notifications are supressed in  " $Str
       Set-UnifiedGroup $Str -UnifiedGroupWelcomeMessageEnable:$false
    }
}
#endregion

Log-EntrySimple "______________________________________"
Log-EntrySimple "_________End of Code Pass_____________"
Log-EntrySimple "______________________________________"

