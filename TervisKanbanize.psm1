#Requires -Modules KanbanizePowerShell, TrackITUnOfficial, TrackITWebAPIPowerShell, get-MultipleChoiceQuestionAnswered, TervisTrackITWebAPIPowerShell
#Requires -Version 4

function Install-TervisKanbanize {
    param (
        $Email,
        $Pass
    )

    Set-KanbanizeSubDomain -SubDomain tervis -Permanent
    Invoke-KanbanizeLogin -Permanent @PSBoundParameters
}

function Add-TervisKanbanizeCardProperties {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$Card,
        [Switch]$PassThru
    )
    process {
        $Card |        
        Add-Member -MemberType ScriptProperty -Name TrackITIDFromTitle -Force -Value { 
            if ($this.Title -match " - ") { [int]$($this.Title -split " - ")[0] } 
        } -PassThru |
        Add-Member -MemberType ScriptProperty -Name ScheduledDate -Force -Value { 
            $this."Scheduled Date"
        } -PassThru | 
        Add-Member -MemberType ScriptProperty -Name PositionInt -Force -Value { 
            [int]$this.position 
        } -PassThru | 
        Add-Member -MemberType ScriptProperty -Name PriorityInt -Force -Value { 
            switch($this.color) {
                "#cc1a33" {1} #Red for priority 1
                "#f37325" {2} #Orange for priority 2
                "#77569b" {3} #Purple for priority 3
                "#067db7" {4} #Blue for priority 4
            }
        } -PassThru | 
        Add-Member -MemberType ScriptProperty -Name BoardID -Force -Value { 
            $this.BoardParent 
        } -PassThru | 
        Add-Member -MemberType ScriptProperty -Name CreatedAtDateTime -Force -Value { 
            Get-Date $this.CreatedAt 
        }
        
        foreach ($CustomField in $Card.CustomFields) {
            $Card | Add-Member -MemberType NoteProperty -Name $CustomField.Name -Force -Value (
                $CustomField | Get-CustomFieldValue
            )
        }

        if ($PassThru) { $Card }
    }
}

$CustomFieldTypeMapping = @{
    number = "int"
    date = "DateTime"  
}

function Get-CustomFieldValue {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$CustomField
    )
    process {
        $PowerShellType = $CustomFieldTypeMapping[$CustomField.type]
        if ($PowerShellType) {
            Invoke-Expression "[$PowerShellType]`$CustomField.Value"
        } else {
            $CustomField.Value
        }        
    }
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

        $Cards += $CardsFromBoard
    }

    if($ExcludeDoneAndArchive) {
        $Cards = $Cards |
        where columnpath -NotIn "Done","Archive"
    }

    $Cards | 
    Add-TervisKanbanizeCardProperties -PassThru
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

    $AllCardDetails | Add-TervisKanbanizeCardDetailsProperties -PassThru
}

Function Add-TervisKanbanizeCardDetailsProperties {
    param(
        [Parameter(Mandatory,ValueFromPipeline)]$Card,
        [Switch]$PassThru
    )
    process {
        $Card | 
        Add-Member -Name CreatedDate -MemberType ScriptProperty -Value { 
            $This.HistoryDetails | 
            where historyevent -eq "Task created" |
            Select -ExpandProperty entrydate |
            Get-Date
        } -PassThru | 
        Add-Member -Name CompletedDate -MemberType ScriptProperty -Value { 
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
        } -PassThru |
        Add-Member -Name CompletedYearAndWeek -MemberType ScriptProperty -Value {
            if($this.completedDate) {
                Get-YearAndWeekFromDate $This.CompletedDate
            }
        } -PassThru |
        Add-Member -Name CycleTimeTimeSpan -MemberType ScriptProperty -Value {
            if ($This.CompletedDate) {
                $this.CompletedDate - $this.CreatedDate
            } else {
                $(get-date) - $this.CreatedDate
            }
        }

        $Card.HistoryDetails | 
        Add-Member -Name EntryDateTime -Type ScriptProperty -Value {
            Get-Date $This.EntryDate
        } -PassThru |
        where eventtype -EQ Transitions |
        Add-Member -Name TransitionFromColumn -Type ScriptProperty -Value {
            $This.Details | 
            Select-StringBetween -After "From '" -Before "' "
        } -PassThru |
        Add-Member -Name TransitionToColumn -Type ScriptProperty -Value {
            $This.Details | 
            Select-StringBetween -After "' to '" -Before "'"
        }

        if ($Passthru) {
            $Card
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

Function New-KanbanizeCardFromTrackITWorkOrder {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$WorkOrder,
        $DestinationBoardID,
        $DestinationColumn
    )
    process {
        $CardName = "" + $WorkOrder.Wo_Num + " -  " + $WorkOrder.Task
        $Response = New-KanbanizeTask -BoardID $DestinationBoardID -Title $CardName -CustomFields @{"trackitid"=$WorkOrder.Wo_Num;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w$($WorkOrder.Wo_Num)"} -Column $DestinationColumn -Lane "Planned Work"
        
        Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot
        Edit-TervisTrackITWorkOrder -WorkOrderNumber $WorkOrder.Wo_Num -KanbanizeCardID $Response.id | Out-Null
    }
}

function Get-ApprovedWorkInstructionsInEvernote {
    $ProjectsAndBoards = Get-KanbanizeProjectsAndBoards
    $Project = $projectsAndBoards.projects | where name -eq "Technical Services"
    $Board = $Project.boards | where name -EQ "Help Desk Service Request Types"
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
    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess |
    where columnpath -EQ "Requested.Ready to be worked on" |
    where type -NotIn (Get-ApprovedWorkInstructionsInEvernote)
    
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
    $Card | Select TaskID, Title, Description, Type, deadline, PriorityInt, Reporter | FL
    
    if ($Card.TrackITID) {
        $WorkOrder = Get-TervisTrackITWorkOrder -WorkOrderNumber $Card.TrackITID    
    
        $WorkOrder.AllNotes | 
        sort createddateDate -Descending |
        select createddateDate, CreatedBy, FullText |
        FL
    }
    
    $Task = Get-KanbanizeTaskDetails -BoardID $Card.BoardID -TaskID $Card.TaskID -History yes -Event comment

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

Function Get-WorkOrdersThatHaveKanbanizeCards {
    $KanbanizeProjedctsAndBoards = Get-KanbanizeProjectsAndBoards
    $BoardIDs = $KanbanizeProjedctsAndBoards.projects.boards.ID

    $Cards = $null
    $BoardIDs | % { $Cards += Get-KanbanizeAllTasks -BoardID $_ }
    
    $CardsWithTrackITIDs = $Cards | 
    Add-TervisKanbanizeCardProperties -PassThru | 
    where trackitid
    
    $WorkOrders = Get-TervisTrackITUnOfficialWorkOrder

    $WorkOrders |
    where WOID -In $CardsWithTrackITIDs.TrackITID
}

Function Find-CardsClosedInTrackITButOpenInKanbanize {
    $KanbanizeProjedctsAndBoards = Get-KanbanizeProjectsAndBoards
    $BoardIDs = $KanbanizeProjedctsAndBoards.projects.boards.ID

    $Cards = $null
    $BoardIDs | % { $Cards += Get-KanbanizeAllTasks -BoardID $_ }

    $CardsWithTrackITIDs = $Cards | 
    Add-TervisKanbanizeCardProperties -PassThru | 
    where trackitid
    
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

Function Import-TrackITWorkOrdersBasedOnAssignedTechnician {
    param(
        [Parameter(Mandatory)]$AssignedTechnician,
        [Parameter(Mandatory)]$DestinationBoardID
    )
    $workorders = Get-TervisTrackITUnOfficialWorkOrder
    $hashTable = $workorders | group RESPONS -AsHashTable
    $hashTable.$AssignedTechnician | New-KanbanizeCardFromTrackITWorkOrder -DestinationBoardID $DestinationBoardID
}

Function Get-TervisKanbnaizeAllTasksFromAllBoards {
    $KanbanizeProjedctsAndBoards = Get-KanbanizeProjectsAndBoards
    $BoardIDs = $KanbanizeProjedctsAndBoards.projects.boards.ID

    $Cards = $null
    $BoardIDs | % { $Cards += Get-KanbanizeAllTasks -BoardID $_ }
    $Cards | Add-TervisKanbanizeCardProperties -PassThru
}

Function Sync-KanbanizeTaskIDToTracKITWorkOrder {
    $Cards = Get-TervisKanbnaizeAllTasksFromAllBoards
    $CardsWithTrackITIDs = $Cards | where trackitid
    
    $WorkOrders = Get-TervisTrackITUnOfficialWorkOrder
    $WorkOrdersWithOutKanbanizeIDs = $WorkOrders | where { -not $_.KanbanizeID }

    $CardsWithTrackITIDsOpenInTrackIT = $CardsWithTrackITIDs | where trackitid -in $($WorkOrdersWithOutKanbanizeIDs.WOID)
    
    foreach ($Card in $CardsWithTrackITIDsOpenInTrackIT) {
        Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot
        Edit-TervisTrackITWorkOrder -WorkOrderNumber $Card.TrackITID -KanbanizeCardID $Card.taskid
    }

}

Function Remove-CardsNoLongerOpenInTrackIT {

}

Function Get-WorkInstructionURI {
    param (
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)]$Type,
        $Cards = $(Get-KanbanizeAllTasks -BoardID 33 | Add-TervisKanbanizeCardProperties -PassThru)
    )
    $Cards | 
    where Type -eq $Type | 
    select -ExpandProperty "Work Instruction"
}

Function Get-TervisKanbanizePowerShellTypeMetaData {
    param (
        $Cards = $(Get-KanbanizeAllTasks -BoardID 33 | Add-TervisKanbanizeCardProperties -PassThru)
    )
    $Cards | 
    where BoardID -EQ 33
}

Function Invoke-SortCardsOnHelpDeskProcess {
    $Cards = Get-TervisKanbnaizeAllTasksFromAllBoards

    $CardsThatNeedToBeSorted = $Cards |
    where boardid -eq 32 |
    where columnpath -Match requested |
    where lanename -eq "Planned Work" |
    sort positionint

    $SortedCards = $CardsThatNeedToBeSorted |
    sort priorityint, trackitid


    foreach ($Card in $SortedCards) {
        Move-KanbanizeTask -BoardID 32 -TaskID $Card.taskid -Column "Requested.Ready to be worked on" -Lane "Planned Work" -Position ($SortedCards.IndexOf($Card))
    }
}

Function Measure-HelpDeskKanbanizeBoardHealth {
    $KanbanizeBoards = Get-KanbanizeProjectsAndBoards
    $HelpDeskBoardID = $KanbanizeBoards.projects.boards | where name -EQ "Help Desk Process" | select -ExpandProperty ID

    $Cards = Get-KanbanizeAllTasks -BoardID $HelpDeskBoardID
    $Types = Get-KanbanizeFullBoardSettings -BoardID $HelpDeskBoardID | select -ExpandProperty types

    $Cards | Add-Member -MemberType ScriptProperty -Name PositionInt -Value { [int]$this.position }

    foreach ($Card in $Cards) {
        $Card.CustomFields | % { 
            $Card | Add-Member -MemberType NoteProperty -Name $_.Name -Value $_.value
        }
    }

    $CardsInScheduledDateColumn = $Cards |
        where columnpath -eq "In Progress.Waiting for Scheduled Date"

    $CardsWithScheduledDate = $CardsInScheduledDateColumn | where {$_."Scheduled Date"}

    $PercentageOfCardsInTheWaitingForScheduledDateColumnWithScheduledDateSpecififed = [int](($CardsWithScheduledDate.Count / $CardsInScheduledDateColumn.Count) * 100)
    $MetricMessage = "$PercentageOfCardsInTheWaitingForScheduledDateColumnWithScheduledDateSpecififed% of cards in the Waiting for scheduled date column that have a scheduled date set"
    
    $MetricColor = switch ($PercentageOfCardsInTheWaitingForScheduledDateColumnWithScheduledDateSpecififed) {
        {$_ -in $(0..80)} {"Red"}
        {$_ -in $(81..99)} {"Yellow"}
        {$_ -eq 100} {"Green"}
    }

    Write-Host -ForegroundColor $MetricColor $MetricMessage
    
    $CardsWhereTrackITIDNotSet = $Cards | where TrackITID -LT 1
    $MetricValue = $CardsWhereTrackITIDNotSet.count
    $MetricColor = switch ($MetricValue) {
        {$_ -gt 0} {"Red"}
        {$_ -eq 0} {"Green"}
        {$_ -eq $Null} {"Green"; $MetricValue = 0}
    }

    Write-Host -ForegroundColor $MetricColor "$MetricValue cards don't have a TrackITID set"

    #Metric: What is the date of the oldest most recent comment on a card
    $CardsInWaitingForUserFeedback = $Cards |
        where columnpath -eq "In Progress.Waiting for user feedback" |
        where lanename -EQ "Unplanned Work"
         
    $CardsInWaitingForUserFeedbackDetails = $CardsInWaitingForUserFeedback | % {
        Get-KanbanizeTaskDetails -BoardID $HelpDeskBoardID -TaskID $_.taskid -History yes -Event comment
    }
    
    $CardsInWaitingForUserFeedbackDetails.historydetails | Add-Member -MemberType ScriptProperty -Name entrydateDate -Value { get-date $this.entrydate }

    $DateOfTheOldestLastCommentOnAllCardsInWaitingForUserFeedback = $CardsInWaitingForUserFeedbackDetails | % {
        $_.historydetails.entrydatedate | sort -Descending | select -First 1
    } | sort | select -first 1

    $NumberOfDays = $($(Get-Date) - $DateOfTheOldestLastCommentOnAllCardsInWaitingForUserFeedback | select -ExpandProperty days)
    
    $MetricColor = switch ($NumberOfDays) {
        { $_ -gt 7 } {"Red"}
        { $_ -le 7 -and $_ -gt 3} {"Yellow"}
        default  {"Green"}
    }

    Write-Host -ForegroundColor $MetricColor "$NumberOfDays is the longest number of days since we reached out to a customer on a card we are waiting for user feedback in unplanned work"

}

Function Write-PercentageMetric {
    param (
        $MetricValuesThatMeanTerrible,
        $MetricValuesThatMeanBad,
        $MetricValuesThatMeanGood,
        $MetricValue,
        $MetricMessage
    )

    $MetricColor = switch ($MetricValue) {
        {$_ -in $MetricValuesThatMeanTerrible} {"Red"}
        {$_ -in $MetricValuesThatMeanBad} {"Yellow"}
        {$_ -in $MetricValuesThatMeanGood} {"Green"}
    }

    Write-Host -ForegroundColor $MetricColor $MetricMessage
}
