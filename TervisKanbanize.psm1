filter Mixin-TervisKanbanizeCardProperties {
    $_ | Add-Member -MemberType ScriptProperty -Name TrackITID -Value { [int]$($this.customfields | Where name -eq "trackitid" | select -ExpandProperty value) }
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

function Move-CompletedCardsThatHaveAllInformationToArchive {
    $KanbanizeBoards = Get-KanbanizeProjectsAndBoards
    $HelpDeskBoardID = $KanbanizeBoards.projects.boards | where name -EQ "Help Desk Technician Process" | select -ExpandProperty ID
    $TriageBoardID = $KanbanizeBoards.projects.boards | where name -EQ "Help Desk Triage Process" | select -ExpandProperty ID

    $TechnicianProcessCards = Get-KanbanizeAllTasks -BoardID $HelpDeskBoardID
    $TechnicianProcessCards | Add-Member -MemberType NoteProperty -Name BoardID -Value $HelpDeskBoardID
    
    $TriageProcessCards = Get-KanbanizeAllTasks -BoardID $TriageBoardID
    $TriageProcessCards | Add-Member -MemberType NoteProperty -Name BoardID -Value $TriageBoardID

    $Cards = $TechnicianProcessCards + $TriageProcessCards

    $Cards | Mixin-TervisKanbanizeCardProperties

    $OpenTrackITWorkOrders = get-TrackITWorkOrders

    $CardsThatCanBeArchived = $Cards | 
    where columnpath -Match "Done" |
    where type -ne "None" |
    where assignee -NE "None" |
    where color -in ("#cc1a33","#f37325","#77569b","#067db7") |
    where TrackITID |
    where TrackITID -NotIn $($OpenTrackITWorkOrders.woid)

    foreach ($Card in $CardsThatCanBeArchived){
        Move-KanbanizeTask -BoardID $Card.BoardID -TaskID $Card.TaskID -Column "Archive"
    }
}