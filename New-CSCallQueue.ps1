#region: Clear old sessions from PS
Write-host -ForegroundColor Yellow  "Clearing previous session connections to cloud servcies"
Log-Entry "Clearing previous session connections to cloud servcies"
$logins = get-azcontext
foreach ($account in $logins)
{
Disconnect-AzAccount -Username $account.account
}
#endregion

#region:Functions
#Configuring logging... its a Function == When needing a log entry simply use: Log-Entry "Log text here"

$LogPath = "C:\LogDump\"+[string](Get-Date -format yyyyMdd)+" NewCSCallQ.txt"
Write-host -ForegroundColor Yellow  "Logs to be found at c:\logdump"
Log-Entry "Logs to be found at c:\logdump"
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

#starting code run divider
Log-EntrySimple "______________________________________"
Log-EntrySimple "_________Starting Code Pass___________"
Log-EntrySimple "______________________________________"

#connect to AAD - no credentials cached to comply with Modern Auth
Write-host -ForegroundColor Yellow  "Connecting to AzureAD"
log-entry "Connecting to AzureAD"
Connect-AzureAD
Write-host -ForegroundColor Yellow  "MicrosoftTeams"
log-entry "MicrosoftTeams"
Connect-MicrosoftTeams
##connect to EoL - no credentials cached to comply with Modern Auth
Write-host -ForegroundColor Yellow  "Connecting to EoL"
log-entry "Connecting to EoL"
Connect-ExchangeOnline

#pop up button to add manual input and notify of the next action required
   
$Shell          = New-Object -ComObject "WScript.Shell"
$Button         = $Shell.Popup("Click OK to continue and selct the Excel workbook to import.", 0, "New-CSCallQueue", 0)

#pop up window to navigate to the excel document required, filtered to Excel only

Add-Type -AssemblyName System.Windows.Forms
$FileBrowser    = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter      = 'SpreadSheet (*.xlsx)|*.xlsx'
}
$null           = $FileBrowser.ShowDialog()

#import the worksheet selected in the action above
$QueueName      = Import-Excel -path $filebrowser.FileName -worksheetname new-CsCallQueue
#get the onMicrosoft.com domain name required fro the resource account creation
$onMSDomainname = (Get-AzureAdDomain | Where-Object {($_.name -Like "*onmicrosoft.com") -and ($_.name -notLike "*mail.onmicrosoft.com")} | select-object -property name).name
#AppID for call queue
$AppID          =  "11cd3e2e-fccb-42ad-ad00-878b93575e07"
#endregion

#region:add queues

#variables for building the objects
#loop testing - $Queue = $QueueName[0] - create an array of a single item
foreach ($Queue in $QueueName) 
{
    #variable to build CQ name string
    $StrCQ           =  "CQ-"+$Queue.Name
    #variable to build check queue test
    $CheckQueue     = Get-CsCallQueue -NameFilter $StrCQ
    #variable to build resource account name
    $StrRA           =  "RA-"+$Queue.Name
    #variable to build UPN for the check process - takes name, replaces spaces then rebuilds as UPN 
    $UPN            = $($StrRA+"@"+$onMSDomainname).Replace(" ","")
    #variable to build check resource account for the while function
    $CheckRA         = Get-CsOnlineApplicationInstance -identity $UPN


    #if something came back then value is true,then skip. Else, with no value create call queue
    If($CheckQueue) 
    {
       Write-host -ForegroundColor Yellow  "a queue with the name was already found in AAD "$StrCQ
       log-entry "a queue with the name was already found in AAD "$StrCQ
    } 
    Else 
    { 
        Write-host -ForegroundColor Yellow "creating queue " $StrCQ
        Log-Entry  "creating queue " $StrCQ
         #Check Resource Account Exisit, if not create the Resource Account
        if (!$CheckRA) 
        {
            Write-Host -ForegroundColor Green "creating resouse account "  $StrRA 
            Log-Entry "creating resouse account "  $StrRA
            New-CsOnlineApplicationInstance -UserPrincipalName $UPN -ApplicationId $AppID  -DisplayName $StrRA
            $CheckRA         = Get-CsOnlineApplicationInstance -identity $UPN
        }
        #while the Resource Account check is not true sleep 2 seconds and then try again
        while (!$CheckRA)
        {
            Write-Host -ForegroundColor Yellow "Check RA Not found - starting sleep"
            log-entry "Check RA Not found - starting sleep"
            start-sleep -seconds 2
            $CheckRA         = Get-CsOnlineApplicationInstance -Identity $UPN
        }
        #Create the Call Queue
        if ($queuename.OverflowAction -eq "SharedVoicemail") 
        {
            $identity   = $(get-UnifiedGroup $Queue.Name).ExternalDirectoryObjectId
        }
        New-CsCallQueue `
        -Name $StrCQ `
        -RoutingMethod $queue.routingmethod `
        -AllowOptOut $queue.AllowOptOut `
        -ConferenceMode $queue.ConferenceMode `
        -PresenceBasedRouting $queue.PresenceBasedRouting `
        -AgentAlertTime $queue.AgentAlertTime `
        -LanguageId $queue.LanguageId `
        -OverflowThreshold $queue.OverflowThreshold `
        -OverflowAction $queue.OverflowAction `
        -OverflowActionTarget $identity `
        -OverflowSharedVoicemailTextToSpeechPrompt $queue.OverflowSharedVoicemailTextToSpeechPrompt `
        -EnableOverflowSharedVoicemailTranscription $queue.EnableOverflowSharedVoicemailTranscription `
        -TimeoutThreshold $queue.TimeoutThreshold `
        -TimeoutAction $queue.TimeoutAction `
        -TimeoutActionTarget $identity `
        -TimeoutSharedVoicemailTextToSpeechPrompt $queue.TimeoutSharedVoicemailTextToSpeechPrompt `
        -EnableTimeoutSharedVoicemailTranscription $queue.EnableTimeoutSharedVoicemailTranscription `
        -UseDefaultMusicOnHold $queue.UseDefaultMusicOnHold

        #check the queue has been created
        $CheckQueue     = Get-CsCallQueue -NameFilter $StrCQ           
        Write-host -ForegroundColor green "Successfully created queue CQ-"$StrCQ
        Log-entry -ForegroundColor green "Successfully created queue CQ-"$StrCQ
    }  
}

#endregion
#ending code run divider
Log-EntrySimple "______________________________________"
Log-EntrySimple "_________End of Code Pass___________"
Log-EntrySimple "______________________________________"


#end logging 