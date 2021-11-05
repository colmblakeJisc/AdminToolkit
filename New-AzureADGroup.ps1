#region:Functions

#Configuring logging... its a Function == When needing a log entry simply use: Log-Entry "Log text here"

$LogPath = "C:\LogDump\"+[string](Get-Date -Format yyyyMdd)+" NewCSCallQ.txt"
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

#region:script setup
Log-EntrySimple "______________________________________"
Log-EntrySimple "_________Starting Code Pass___________"
Log-EntrySimple "______________________________________"

#connect to AAD - no credentials cached to comply with Modern Auth
$ADContext = ""
$ErrorActionPreference= 'silentlycontinue' #seems to be a bug with the command below currently https://github.com/Azure/azure-docs-powershell-azuread/issues/591
$ADContext = Get-AzureADCurrentSessionInfo
$ErrorActionPreference= 'continue' #returning back to default 
if ($ADContext)
{
    $confirmation = Read-Host "Continue Deployment to : $($ADContext.TenantDomain) Are you Sure You Want To Proceed (y/n)"
    if ($confirmation -eq 'n') 
    {
        Write-host -ForegroundColor Yellow  "Connecting to AzureAD"
        Disconnect-AzureAD
        log-entry "Connecting to AzureAD"
        Connect-AzureAD
    }
}
Else
{
        Write-host -ForegroundColor Yellow  "Connecting to AzureAD"
        log-entry "Connecting to AzureAD"
        Connect-AzureAD
}
Write-host -ForegroundColor Yellow  "MicrosoftTeams"
log-entry "MicrosoftTeams"
Connect-MicrosoftTeams
##connect to EoL - no credentials cached to comply with Modern Auth
Write-host -ForegroundColor Yellow  "Connecting to EoL"
log-entry "Connecting to EoL"
Connect-ExchangeOnline

#pop up button to add manual input and notify of the next action required
   
$Shell          = New-Object -ComObject "WScript.Shell"
$Button         = $Shell.Popup("Click OK to continue and selct the Excel workbook to import.", 0, "New-AzureAdGroup", 0)

#pop up window to navigate to the excel document required, filtered to Excel only

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser    = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter      = 'SpreadSheet (*.xlsx)|*.xlsx'
}
$null           = $FileBrowser.ShowDialog()

#import the worksheet selected in the action above

$GroupName      = Import-Excel -path $filebrowser.FileName -worksheetname New-AzureADGroup
Write-host -ForegroundColor Yellow  "Excel worksheet imported"
Log-Entry "Excel worksheet imported"
#endregion

#region:add AADGroup
#add AADGroup
#loop testing - $Group = $GroupName[0] - create an array of a single item
foreach ($Group in $GroupName) {
    $Str            = $group.GroupName
    $CheckGroup     = Get-AzureADGroup -SearchString $Str
    
    #if something came back then value is true and skip. with no value create group
    If($CheckGroup) {
       Write-host -ForegroundColor Yellow  "a group with the distinguish name was already found in AAD" $Group.GroupName
       Log-Entry "a group with the distinguish name was already found in AAD" $Group.GroupName
    } 
    Else { 
       Write-host -ForegroundColor Yellow "creating group" $Group.GroupName
       Log-Entry "creating group" $Group.GroupName
       New-AzureADGroup -DisplayName $Group.GroupName -MailEnabled $False -SecurityEnabled $True -MailNickName "NotSet"    
    }
}
#endregion

Log-EntrySimple "______________________________________"
Log-EntrySimple "_________End of Code Pass___________"
Log-EntrySimple "______________________________________"


#end logging 