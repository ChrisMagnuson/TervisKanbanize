﻿#Requires -Modules KanbanizePowerShell, TrackITUnOfficial, TrackITWebAPIPowerShell, get-MultipleChoiceQuestionAnswered, TervisTrackITWebAPIPowerShell
#Requires -Version 4

filter Mixin-TervisKanbanizeCardProperties {
    $_ | Add-Member -MemberType ScriptProperty -Name TrackITID -Value { [int]$($this.customfields | Where name -eq "trackitid" | select -ExpandProperty value) }
    $_ | Add-Member -MemberType ScriptProperty -Name TrackITIDFromTitle -Value { if ($this.Title -match " - ") { [int]$($this.Title -split " - ")[0] } }
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
    $_ | Add-Member -MemberType ScriptProperty -Name BoardID -Value { $this.BoardParent }
}

filter Mixin-TervisKanbanizeArchiveCardProperties {
    $_ | Add-Member -MemberType ScriptProperty -Name ArchivedDate -Value { get-dtate $this.createdorarchived }
}

function Get-KanbanizeTervisHelpDeskCards {
    param(
        [switch]$HelpDeskProcess,
        [switch]$HelpDeskTechnicianProcess,
        [switch]$HelpDeskTriageProcess,
        [Parameter(ParameterSetName='NotContainer')][switch]$ExcludeDoneAndArchive,
        [ValidateSet("archive")][Parameter(Mandatory=$true,ParameterSetName='Container')]$Container,
        [Parameter(Mandatory=$true,ParameterSetName="Container")]$FromDate,
        [Parameter(Mandatory=$true,ParameterSetName="Container")]$ToDate
    )
    $BoardIDs = Get-TervisKanbanizeHelpDeskBoardIDs -HelpDeskProcess:$HelpDeskProcess -HelpDeskTechnicianProcess:$HelpDeskTechnicianProcess -HelpDeskTriageProcess:$HelpDeskTriageProcess
    
    $Cards = @()

    Foreach ($BoardID in $BoardIDs) {
        if ($Container) {
            $CardsFromBoard = Get-TervisKanbanizeAllTasksFromArchive -BoardID $BoardID -FromDate $FromDate -ToDate $ToDate
            $CardsFromBoard | Mixin-TervisKanbanizeArchiveCardProperties
        } else {
            $CardsFromBoard = Get-KanbanizeAllTasks -BoardID $BoardID
        }

        $CardsFromBoard | Add-Member -MemberType NoteProperty -Name BoardID -Value $BoardID
        $Cards += $CardsFromBoard
    }

    if($ExcludeDoneAndArchive) {
        $Cards = $Cards |
        where columnpath -NotIn "Done","Archive"
    }

    $Cards | Mixin-TervisKanbanizeCardProperties
    $Cards
}

function Get-TervisKanbanizeAllTasksFromArchive {
    param(
        $BoardID,
        $FromDate,
        $ToDate
    )
    $progressPreference = 'silentlyContinue'
    
    $Cards = @()

    $ArchiveTaskResults = Get-KanbanizeAllTasks -BoardID $BoardID -Container archive -FromDate $FromDate -ToDate $ToDate
    if($ArchiveTaskResults) {
        $Cards += $ArchiveTaskResults |
        select -ExpandProperty Task

        $TotalNumberOfPages = $ArchiveTaskResults.numberoftasks/$ArchiveTaskResults.tasksperpage
        $TotalNumberOfPagesRoundedUp = [int][Math]::Ceiling($TotalNumberOfPages)

        if ($TotalNumberOfPagesRoundedUp -gt 1) {
            foreach ($PageNumber in 2..$TotalNumberOfPagesRoundedUp) {
               $Cards += Get-KanbanizeAllTasks -BoardID $BoardID -Container archive -Page $PageNumber -FromDate $FromDate -ToDate $ToDate |
               select -ExpandProperty Task
            }
        }
    
        $progressPreference = 'Continue' 
    }
    $Cards
}

function Get-TervisKanbanizeAllTaskDetails {
    param(
        $Cards
    )

    $AllCardDetails = @()
    foreach ($Card in $Cards) {       
        $AllCardDetails += Get-KanbanizeTaskDetails -BoardID $Card.Boardid -TaskID $Card.taskid -History yes
    }

    $AllCardDetails | Mixin-TervisKanbanizeCardDetailsProperties
}

filter Mixin-TervisKanbanizeCardDetailsProperties {
    $_ | Add-Member -MemberType ScriptProperty -Name CreatedDate -Value { 
        $This.HistoryDetails | 
        where historyevent -eq "Task created" |
        Select -ExpandProperty entrydate |
        Get-Date
    }

    $_ | Add-Member -MemberType ScriptProperty -Name CompletedDate -Value { 
        #Get the very last date this card was moved to Done, might have happened twice if it was brought back out of the archive because it was not finished
        if ($This.columnname -in "Archive","Done") {
            $This.HistoryDetails | 
            where historyevent -eq "Task moved" |
            where details -Match "to 'Done'" |
            Select -ExpandProperty entrydate |
            Get-Date |
            sort |
            select -Last 1
        }
    }

    $_ | Add-Member -MemberType ScriptProperty -Name CompletedYearAndWeek -Value {
        if($this.completedDate) {
            Get-YearAndWeekFromDate $This.CompletedDate
        }
    }

    $_ | Add-Member -MemberType ScriptProperty -Name CycleTimeTimeSpan -Value {
        if ($This.CompletedDate) {
            $this.CompletedDate - $this.CreatedDate
        } else {
            $(get-date) - $this.CreatedDate
        }
    }
}


function Get-TervisKanbanizeHelpDeskBoardIDs {
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
    $BoardIDs
}

function Move-CompletedCardsThatHaveAllInformationToArchive {
    $OpenTrackITWorkOrders = get-TervisTrackITWorkOrders
    
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

#Unfinished
function Move-CardsInDoneListThatHaveStillHaveSomethingIncomplete {
    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess
    
    $CardsInDoneList = $Cards |
    where columnpath -Match "Done"
    
    $OpenTrackITWorkOrders = get-TervisTrackITWorkOrders

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
        $CardName = "" + $UnassignedWorkOrder.Wo_Num + " - " + $UnassignedWorkOrder.Task    
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

function Get-ApprovedWorkInstructionsInEvernote {
    $ProjectsAndBoards = Get-KanbanizeProjectsAndBoards
    $Project = $projectsAndBoards.projects | where name -eq "Technical Services"
    $Board = $Project.boards | where name -EQ "Help Desk Standard Requests"
    $BoardSettings = Get-KanbanizeFullBoardSettings -BoardID $Board.id
    $Tasks = Get-KanbanizeAllTasks -BoardID $Board.id

    $Tasks | 
    where { $_.customfields | where name -EQ "Work Instruction" | select -ExpandProperty value } |
    Select -ExpandProperty Type
}

function compare-WorkInstructionTypesInEvernoteWithTypesInKanbanize {
    Compare-Object -ReferenceObject $(get-TervisKanbanizeTypes) -DifferenceObject $ApprovedWorkInstructionsInEvernote -IncludeEqual
}

function Find-CardsOnTechnicianBoardWithWorkInstructions {
    param(
        [switch]$ExcludeDoneAndArchive
    )
    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskTechnicianProcess -ExcludeDoneAndArchive:$ExcludeDoneAndArchive
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
    $HelpDeskProcessBoardID = Get-TervisKanbanizeHelpDeskBoardIDs -HelpDeskProcess
    $CardsToMove = Find-CardsOnTechnicianBoardWithWorkInstructions -ExcludeDoneAndArchive
    
    foreach ($Card in $CardsToMove) {
        Move-KanbanizeTask -BoardID $HelpDeskProcessBoardID -TaskID $Card.TaskID -Column $Card.columnpath -Lane "Planned Work"
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
    $WorkOrder = Get-TervisTrackITWorkOrder -WorkOrderNumber $Card.TrackITID
    $Task = Get-KanbanizeTaskDetails -BoardID $Card.BoardID -TaskID $Card.TaskID -History yes -Event comment

    $Card | Select TaskID, Title, Type, deadline, PriorityInt| FL
    
    $WorkOrder.AllNotes | 
    sort createddateDate -Descending |
    select createddateDate, CreatedBy, FullText |
    FL

    $Task.HistoryDetails |
    where historyevent -ne "Comment deleted" |
    Select EntryDate, Author, Details |
    FL
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

Function Find-CardsClosedInTrackITButOpenInKanbanize {
    $KanbanizeProjedctsAndBoards = Get-KanbanizeProjectsAndBoards
    $BoardIDs = $KanbanizeProjedctsAndBoards.projects.boards.ID

    $Cards = $null
    $BoardIDs | % { $Cards += Get-KanbanizeAllTasks -BoardID $_ }
    $Cards | Mixin-TervisKanbanizeCardProperties
    $CardsWithTrackITIDs = $Cards | where trackitid
    
    $WorkOrders = Get-TervisTrackITUnOfficialWorkOrder

    $CardsWithTrackITIDs |
    where TrackITID -NotIn $WorkOrders.WOID
}

Function Remove-KanbanizeCardsForClosedTrackITs {
    $CardsThatNeedToBeClosed = Find-CardsClosedInTrackITButOpenInKanbanize
    foreach ($Card in $CardsThatNeedToBeClosed) {
        Remove-KanbanizeTask -BoardID $Card.BoardParent -TaskID $Card.TaskID
    }
}
