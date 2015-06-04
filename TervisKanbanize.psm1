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
    $CardsInScheduledDateThatDontHaveScheduledDateSet = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess -HelpDeskTechnicianProcess | 
    where columnpath -Match "Waiting for Scheduled date" | 
    where scheduleddate -eq $null
    
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
            #throw "Testing blowing up the automated track it to kanbnaize importer"
            if($UnassignedWorkOrder.Wo_Num -in $($Cards.TrackITID)) {throw "There is already a card for this Track IT"}

            $Response = New-KanbanizeTask -BoardID $TriageProcessBoardID -Title $CardName -CustomFields @{"trackitid"=$UnassignedWorkOrder.Wo_Num;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w$($UnassignedWorkOrder.Wo_Num)"} -Column $TriageProcessStartingColumn -Lane "Planned Work"
            Edit-TrackITWorkOrder -WorkOrderNumber $UnassignedWorkOrder.Wo_Num -AssignedTechnician "Backlog" | Out-Null
        } catch {            
            $ErrorMessage = "Error importing: " + $CardName
            Send-MailMessage -From HelpDeskBot@tervis.com -to HelpDeskDispatch@tervis.com -subject $ErrorMessage -SmtpServer cudaspam.tervis.com -Body $_.Exception|format-list -force
        }
    }
}

$CurrentWorkInstructions = @"
Printer toner swap out
Printer waste toner box swap out
Termination
Whitelist email address
Internet explorer browser settings reset
Distribution list member add
Distribution list member remove
Distribution list create
Monitor swap or add
Layer 1 equipment get
Layer 1 equipment install
Personal phone email install and TervisWifi install
Software chrome install
Software paint.net install
Uninstall software
iPhone Swap
Iphone get, initialize, install
iPhone work email access grant
Capex new
Active directory user photo update
EBS user responisbilities update
Software oracle sql developer install
Mailbox access grant
Termination IT
EBS rapid planning user responsibilities update
Active directory user password reset
Active directory user phone number update
CRM password reset
"@

function Invoke-PrioritizeConfirmTypeAndMoveCardTechnicainBoard {
    [CmdletBinding()]
    param()

    $VerbosePreference = "continue"

    Import-Module KanbanizePowerShell -Force
    Import-module TrackItWebAPIPowerShell -Force

    Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot

    $KanbanizeBoards = Get-KanbanizeProjectsAndBoards
    $HelpDeskProcessBoardID = $KanbanizeBoards.projects.boards | 
    where name -EQ "Help Desk Process" | 
    select -ExpandProperty ID

    $Types = Get-KanbanizeFullBoardSettings -BoardID $HelpDeskProcessBoardID | select -ExpandProperty types

    $Cards = Get-KanbanizeTervisHelpDeskCards -HelpDeskProcess

    $WaitingToBePrioritized = $Cards |
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
            $Priority = get-MultipleChoiceQuestionAnswered -Question "What priority level should this request have?" -Choices 1,2,3,4
            $color = switch($Priority) {
                1 { "cc1a33" } #Red for priority 1
                2 { "f37325" } #Orange for priority 2
                3 { "77569b" } #Yello for priority 3
                4 { "067db7" } #Blue for priority 4
            }
            Write-Verbose "Color: $color"
            Edit-KanbanizeTask -BoardID $Card.BoardID -TaskID $Task.taskid -Color $color
        }

        $WorkInstructionsForThisRequest = get-MultipleChoiceQuestionAnswered -Question "Are there work instructions to complete this request?" -Choices "Yes","No" | 
        ConvertTo-Boolean
        
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