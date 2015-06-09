filter Mixin-TervisKanbanizeCardProperties {
    $_ | Add-Member -MemberType ScriptProperty -Name TrackITID -Value { [int]$($this.customfields | Where name -eq "trackitid" | select -ExpandProperty value) }
    $_ | Add-Member -MemberType ScriptProperty -Name ScheduledDate -Value { $($this.customfields | Where name -eq "Scheduled Date" | select -ExpandProperty value) }
    $_ | Add-Member -MemberType ScriptProperty -Name PositionInt -Value { [int]$this.position }
    $_ | Add-Member -MemberType ScriptProperty -Name PriorityInt -Value { 
        switch($this.color) {
            "#cc1a33" {1} #Red for priority 1
            "#f37325" {2} #Orange for priority 2
            "#77569b" {3} #Purple for priority 3
            "#067db7" {4} #Blue for priority 4
        }
    }
}

function Get-KanbanizeTervisHelpDeskCards {
    param(
        [switch]$HelpDeskProcess,
        [switch]$HelpDeskTechnicianProcess,
        [switch]$HelpDeskTriageProcess
    )
    
    $KanbanizeBoards = Get-KanbanizeProjectsAndBoards

    $BoardIDs = $KanbanizeBoards.projects.boards | 
    where {
        ($_.name -eq "Help Desk Technician Process" -and $HelpDeskTechnicianProcess) -or
        ($_.name -eq "Help Desk Process" -and $HelpDeskProcess) -or
        ($_.name -eq "Help Desk Triage Process" -and $HelpDeskTriageProcess)
    } | 
    select -ExpandProperty ID
    
    $Cards = @()

    Foreach ($BoardID in $BoardIDs) {
        $CardsFromBoard = Get-KanbanizeAllTasks -BoardID $BoardID
        $CardsFromBoard | Add-Member -MemberType NoteProperty -Name BoardID -Value $BoardID
        $Cards += $CardsFromBoard
    }

    $Cards | Mixin-TervisKanbanizeCardProperties
    $Cards
}

function Move-CompletedCardsThatHaveAllInformationToArchive {
    $OpenTrackITWorkOrders = get-TrackITWorkOrders
    
    $CardsThatCanBeArchived = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess | 
    where columnpath -Match "Done" |
    where type -ne "None" |
    where assignee -NE "None" |
    where color -in ("#cc1a33","#f37325","#77569b","#067db7") |
    where TrackITID |
    where TrackITID -NotIn $($OpenTrackITWorkOrders.woid)

    foreach ($Card in $CardsThatCanBeArchived) {
        Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.TaskID -Column "Archive"
    }
}

function Move-CardsInScheduledDateThatDontHaveScheduledDateSet {
    $CardsInScheduledDateThatDontHaveScheduledDateSet = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess -HelpDeskTriageProcess | 
    where columnpath -Match "Waiting for Scheduled date" | 
    where {$_.scheduleddate -eq $null -or $_.scheduleddate -eq "" }
    
    foreach ($Card in $CardsInScheduledDateThatDontHaveScheduledDateSet) {
        Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.TaskID -Column "In Progress.Waiting to be worked on"
    }
}

function Move-CardsInDoneListThatHaveStillHaveSomethingIncomplete {
    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess
    
    $CardsInDoneList = $Cards |
    where columnpath -Match "Done"
    
    $OpenTrackITWorkOrders = get-TrackITWorkOrders

    $CardsThatAreOpenInTrackITButDoneInKanbanize = Compare-Object -ReferenceObject $OpenTrackITWorkOrders.woid -DifferenceObject $Cardsindonelist.trackitid -PassThru -IncludeEqual |
    where sideindicator -EQ "=="

    foreach ($Card in $CardsThatCanBeArchived){
        Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.TaskID -Column "Archive"
    }
}

function Import-UnassignedTrackItsToKanbanize {
    Import-Module TrackITWebAPIPowerShell -Force

    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess -HelpDeskTriageProcess

$QueryToGetUnassignedWorkOrders = @"
Select Wo_num, task, request_fullname, request_email
  from [TRACKIT9_DATA].[dbo].[vTASKS_BROWSE]
  Where RESPONS IS Null AND
  WorkOrderStatusName != 'Closed'
"@

    Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot
    $TriageProcessBoardID = 29
    $TriageProcessStartingColumn = "Requested"

    $UnassignedWorkOrders = Invoke-SQL -dataSource sql -database TRACKIT9_DATA -sqlCommand $QueryToGetUnassignedWorkOrders

    foreach ($UnassignedWorkOrder in $UnassignedWorkOrders ) {
        $CardName = "" + $UnassignedWorkOrder.Wo_Num + " -  " + $UnassignedWorkOrder.Task    
        try {
            if($UnassignedWorkOrder.Wo_Num -in $($Cards.TrackITID)) {throw "There is already a card for this Track IT"}

            $Response = New-KanbanizeTask -BoardID $TriageProcessBoardID -Title $CardName -CustomFields @{"trackitid"=$UnassignedWorkOrder.Wo_Num;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w$($UnassignedWorkOrder.Wo_Num)"} -Column $TriageProcessStartingColumn -Lane "Planned Work"
            Edit-TrackITWorkOrder -WorkOrderNumber $UnassignedWorkOrder.Wo_Num -AssignedTechnician "Backlog" | Out-Null
        } catch {            
            $ErrorMessage = "Error running Import-UnassignedTrackItsToKanbanize: " + $CardName
            Send-MailMessage -From HelpDeskBot@tervis.com -to HelpDeskDispatch@tervis.com -subject $ErrorMessage -SmtpServer cudaspam.tervis.com -Body $_.Exception|format-list -force
        }
    }
}

function Import-TrackItToKanbanize {
    param (
        $TrackITID
    )
    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess -HelpDeskTriageProcess
    
$QueryToGetUnassignedWorkOrders = @"
Select Wo_num, task
  from [TRACKIT9_DATA].[dbo].[vTASKS_BROWSE]
  Where RESPONS IS Null AND
  WorkOrderStatusName != 'Closed' AND
  Wo_num = $TrackITID
"@
    Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot

    $TriageProcessBoardID = 29
    $TriageProcessStartingColumn = "Requested"

    $UnassignedWorkOrders = Invoke-SQL -dataSource sql -database TRACKIT9_DATA -sqlCommand $QueryToGetUnassignedWorkOrders

    foreach ($UnassignedWorkOrder in $UnassignedWorkOrders ) {
        $CardName = "" + $UnassignedWorkOrder.Wo_Num + " -  " + $UnassignedWorkOrder.Task    
        try {
            if($UnassignedWorkOrder.Wo_Num -in $($Cards.TrackITID)) {throw "There is already a card for this Track IT"}

            $Response = New-KanbanizeTask -BoardID $TriageProcessBoardID -Title $CardName -CustomFields @{"trackitid"=$UnassignedWorkOrder.Wo_Num;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w$($UnassignedWorkOrder.Wo_Num)"} -Column $TriageProcessStartingColumn -Lane "Planned Work"
            Edit-TrackITWorkOrder -WorkOrderNumber $UnassignedWorkOrder.Wo_Num -AssignedTechnician "Backlog" | Out-Null
        } catch {            
            $ErrorMessage = "Error running Import-TrackItToKanbanize: " + $CardName
            Send-MailMessage -From HelpDeskBot@tervis.com -to HelpDeskDispatch@tervis.com -subject $ErrorMessage -SmtpServer cudaspam.tervis.com -Body $_.Exception|format-list -force
        }
    }
}

$ApprovedWorkInstructionsInEvernote = @"
Printer toner swap
Printer waste toner box swap
Termination
Whitelist email address
Internet explorer browser settings reset
Distribution group member add
Distribution group member remove
Distribution group create
Monitor swap or add
Layer 1 equipment get
Personal phone work email install and TervisWifi install
Software chrome install
Software paint.net install
Uninstall software
iPhone Swap
iPhone get initialize and install
iPhone work email access grant
Active directory user photo update
EBS user responsibilities update
Software oracle sql developer install
Mailbox access grant
Termination IT
EBS rapid planning user responsibilities update
Active directory user password reset
Active directory user phone number update
CRM password reset
Tradeshow iPad initialize
Computer rename
Computer windows add to domain
Mailbox New
Remote application navision install
Employee or temp hire new
"@ -split "`r`n"

function compare-WorkInstructionTypesInEvernoteWithTypesInKanbanize {
    Compare-Object -ReferenceObject $(get-TervisKanbanizeTypes) -DifferenceObject $ApprovedWorkInstructionsInEvernote -IncludeEqual
}

function Find-CardsOnTechnicianBoardWithWorkInstructions {
    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskTechnicianProcess
    $Cards |
    #where columnpath -EQ "Requested.Ready to be worked on" |
    where type -In $ApprovedWorkInstructionsInEvernote
}

function Find-MostImportantWorkInstructionsToCreate {
    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskTechnicianProcess |
    where columnpath -EQ "Requested.Ready to be worked on" |
    where type -NotIn $ApprovedWorkInstructionsInEvernote
    
    $Cards|group type| sort count -Descending | select count, name
}

function Move-CardsWithWorkInstructionsToHelpDeskProcessBoard {
    $KanbanizeBoards = Get-KanbanizeProjectsAndBoards
    $HelpDeskProcessBoardID = $KanbanizeBoards.projects.boards | where name -EQ "Help Desk Process" | select -ExpandProperty ID
    $CardsToMove = Find-CardsOnTechnicianBoardWithWorkInstructions
    
    foreach ($Card in $CardsToMove) {
        Move-KanbanizeTask -BoardID $HelpDeskProcessBoardID -TaskID $Card.TaskID -Column "Requested.Ready to be worked on" -Lane "Planned Work"
    }
}

function get-TervisKanbanizeTypes {
    $KanbanizeBoards = Get-KanbanizeProjectsAndBoards
    $HelpDeskProcessBoardID = $KanbanizeBoards.projects.boards | 
    where name -EQ "Help Desk Process" | 
    select -ExpandProperty ID

    $Types = Get-KanbanizeFullBoardSettings -BoardID $HelpDeskProcessBoardID | select -ExpandProperty types
    $Types
}

function Get-TervisWorkOrderDetails {
    param(
        $Card
    )

    $WorkOrder = Get-TrackITWorkOrder -WorkOrderNumber $Card.TrackITID
    $Task = Get-KanbanizeTaskDetails -BoardID $Card.BoardID -TaskID $Card.TaskID -History yes -Event comment

    $Card | Select TaskID, Title, Type, deadline, PriorityInt| FL
    
    $WorkOrder.data | Add-Member -MemberType ScriptProperty -Name AllNotes -Value {
        $This.notes | GM | where membertype -EQ noteproperty | % { $This.notes.$($_.name) }
    }
    $WorkOrder.data.AllNotes | Add-Member -MemberType ScriptProperty -Name createddateDate -Value { get-date $this.createddate }
    
    $WorkOrder.data.AllNotes | 
    sort createddateDate -Descending |
    select createddateDate, CreatedBy, FullText |
    FL

    $Task.HistoryDetails |
    where historyevent -ne "Comment deleted" |
    Select EntryDate, Author, Details |
    FL
}

function Invoke-PrioritizeConfirmTypeAndMoveCard {
    [CmdletBinding()]
    param()

    $VerbosePreference = "continue"

    Import-Module KanbanizePowerShell -Force
    Import-module TrackItWebAPIPowerShell -Force

    Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot

    $KanbanizeBoards = Get-KanbanizeProjectsAndBoards

    $HelpDeskProcessBoardID = $KanbanizeBoards.projects.boards | where name -EQ "Help Desk Process" | select -ExpandProperty ID
    $HelpDeskTechnicianProcessBoardID = $KanbanizeBoards.projects.boards | where name -EQ "Help Desk Technician Process" | select -ExpandProperty ID

    $Types = get-TervisKanbanizeTypes

    $WaitingToBePrioritized = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess |
    where columnpath -Match "Waiting to be prioritized" |
    sort positionint

    $global:CardsThatNeedToBeCreatedTypes = @()
    $global:ToBeCreatedTypes = @()

    foreach ($Card in $WaitingToBePrioritized) {
        $WorkOrder = Get-TrackITWorkOrder -WorkOrderNumber $Card.TrackITID
        $Task = Get-KanbanizeTaskDetails -BoardID $Card.BoardID -TaskID $Card.TaskID -History yes -Event comment

        $Card | Select TaskID, Title, Type, deadline, PriorityInt| FL
    
        $WorkOrder.data | Add-Member -MemberType ScriptProperty -Name AllNotes -Value {
            $This.notes | GM | where membertype -EQ noteproperty | % { $This.notes.$($_.name) }
        }
        $WorkOrder.data.AllNotes | Add-Member -MemberType ScriptProperty -Name createddateDate -Value { get-date $this.createddate }
    
        $WorkOrder.data.AllNotes | 
        sort createddateDate -Descending |
        select createddateDate, CreatedBy, FullText |
        FL

        $Task.HistoryDetails |
        where historyevent -ne "Comment deleted" |
        Select EntryDate, Author, Details |
        FL

        read-host "Hit enter once you have reviewed the details about this request"

        if($Card.Type -ne "None") {
            $TypeCorrect = get-MultipleChoiceQuestionAnswered -Question "Type correct?" -Choices "Yes","No" | ConvertTo-Boolean               
        }
        
        if( !$TypeCorrect -or ($Card.Type -eq "None") )
        {        
            $SelectedType = $Types | Out-GridView -PassThru
    
            if ($SelectedType -ne $null) {
                Edit-KanbanizeTask -TaskID $Card.taskid -BoardID $Card.BoardID -Type $SelectedType        
            } else {
                $ToBeCreatedSelectedType = $global:ToBeCreatedTypes | Out-GridView -PassThru
                if($ToBeCreatedSelectedType -ne $null) {
                    $global:CardsThatNeedToBeCreatedTypes += [pscustomobject]@{taskid=$Card.taskid; type=$ToBeCreatedSelectedType;BoardID=$Card.BoardID}
                } else {
                    $ToBeCreatedSelectedType = read-host "Enter the new type you want to use for this card"
                    $global:CardsThatNeedToBeCreatedTypes += [pscustomobject]@{taskid=$Card.taskid; type=$ToBeCreatedSelectedType;BoardID=$Card.BoardID}
                    $global:ToBeCreatedTypes += $ToBeCreatedSelectedType
                }
            }
        }

        if($card.color -notin ("#cc1a33","#f37325","#77569b","#067db7")) {
            $Priority = get-MultipleChoiceQuestionAnswered -Question "What priority level should this request have?" -Choices 1,2,3,4 -DefaultChoice 3
            $color = switch($Priority) {
                1 { "cc1a33" } #Red for priority 1
                2 { "f37325" } #Orange for priority 2
                3 { "77569b" } #Yello for priority 3
                4 { "067db7" } #Blue for priority 4
            }
            Write-Verbose "Color: $color"
            Edit-KanbanizeTask -BoardID $Card.BoardID -TaskID $Task.taskid -Color $color
        }

        $WorkInstructionsForThisRequest = $card.Type -in $ApprovedWorkInstructionsInEvernote
        
        if($WorkInstructionsForThisRequest) {
            $DestinationBoardID = $HelpDeskProcessBoardID
        } else {
            $DestinationBoardID = $HelpDeskTechnicianProcessBoardID
        }
        Write-Verbose "Destination column: $DestinationBoardID"

        $NeedToBeEscalated = get-MultipleChoiceQuestionAnswered -Question "Does this need to be escalated?" -Choices "Yes","No" | 
        ConvertTo-Boolean
        
        if($NeedToBeEscalated) {
            $DestinationLane = "Unplanned Work"
            Move-KanbanizeTask -BoardID $DestinationBoardID -TaskID $Task.taskid -Lane $DestinationLane -Column "Requested.Ready to be worked on"
        } else { 
            $DestinationLane = "Planned Work"
            Move-KanbanizeTask -BoardID $DestinationBoardID -TaskID $Task.taskid -Lane $DestinationLane -Column "Requested.Ready to be worked on"

            <#
            $CardsThatNeedToBeSorted = $Cards | 
            where {$_.columnpath -eq $DestinationColumn -and $_.lanename -eq "Planned Work"} |
            sort positionint

            $SortedCards = $CardsThatNeedToBeSorted |
            sort priorityint, trackitid
            $PositionOfTheLastCardInTheSamePriortiyLevel = $SortedCards |
                where priorityint -EQ $(if($Card.PriorityInt){$Card.PriorityInt}else{$Priority}) |
                select -Last 1 -ExpandProperty PositionInt
            
            $RightPosition = if($PositionOfTheLastCardInTheSamePriortiyLevel) {
                $PositionOfTheLastCardInTheSamePriortiyLevel + 1
            } else { 0 }
            Write-Verbose "Rightposition in column: $RightPosition"
            
            Move-KanbanizeTask -BoardID $DestinationBoardID -TaskID $Task.taskid -Lane $DestinationLane -Column $DestinationColumn -Position $RightPosition
            #>
        }

        Write-Verbose "DestinationLane: $DestinationLane"
    }

    $global:ToBeCreatedTypes
    Read-Host "Create types in Kanbanize for all the types listed above and then hit enter"

    $global:CardsThatNeedToBeCreatedTypes
    $global:CardsThatNeedToBeCreatedTypes | % {
        Edit-KanbanizeTask -TaskID $_.taskid -BoardID $_.BoardID -Type $_.type
    }

}

function ConvertTo-Boolean {
    param(
        [Parameter(Mandatory=$false,ValueFromPipeline=$true)][string] $value
    )
    switch ($value) {
        "y" { return $true; }
        "yes" { return $true; }
        "true" { return $true; }
        "t" { return $true; }
        1 { return $true; }
        "n" { return $false; }
        "no" { return $false; }
        "false" { return $false; }
        "f" { return $false; } 
        0 { return $false; }
    }
}

function Get-RequestorMailtoLinkForCard {
    param(
        [Parameter(Mandatory=$true)]$Card
    )
    Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot

    $RequestorEmailAddress = Get-TrackITWorkOrderDetails -WorkOrderNumber $Card.TrackITID | 
    select -ExpandProperty Request_Email

    $MailToURI = New-MailToURI -To "$RequestorEmailAddress,tervis_notifications@tervis.com" -Subject $Card.Title + "{$Card.BoardID}{$Card.TaskID}"

    <# means to open the mailto without opening a new tab
        Kanbanize does not allow javascript: URIs in their custom fields so this does not currently work
    $JavaScriptFunctionForMailto = 'window.location.href = "$MailToURI"'

    $FinalURL = "javascript:(function()%7B$([Uri]::EscapeDataString($JavaScriptFunctionForMailto)) %7D)()"
    #>

    $MailToURI
}

function Close-TrackITWhenDoneInKanbanize {

$Message = @"
{Requestor},

Please do not reply to this email.

The work order referenced in the subject of this email has been closed out.

You may have an email from Tervis_Notifications@kanbnaize.com with more details on the resolution of your work order.

If you think this was closed out in error or this issue is not fully resolved please call extension 2248 or 941-441-3168.

Thanks,

Help Desk Team

"@

}
