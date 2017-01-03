function Sync-TrackITAndKanbanize {
    $KanbanizeProjedctsAndBoards = Get-KanbanizeProjectsAndBoards
    $BoardIDs = $KanbanizeProjedctsAndBoards.projects.boards.ID

    $Cards = $null
    $BoardIDs | % { $Cards += Get-KanbanizeAllTasks -BoardID $_ }
    $Cards | Mixin-TervisKanbanizeCardProperties
    $CardsWithTrackITIDs = $Cards | where trackitid
    
    $WorkOrders = Get-TervisTrackITUnOfficialWorkOrder
    $WorkOrdersWithOutKanbanizeIDs = $WorkOrders | where { -not $_.KanbanizeID }

    $CardsWithTrackITIDsOpenInTrackIT = $CardsWithTrackITIDs | where trackitid -in $($WorkOrders.WOID)

    $CardsWithTrackITIDsOpenInTrackIT = $CardsWithTrackITIDs | where trackitid -in $($WorkOrdersWithOutKanbanizeIDs.WOID)
    
    foreach ($Card in $CardsWithTrackITIDsOpenInTrackIT) {
        $Board = $KanbanizeProjedctsAndBoards.projects.boards | where {$_.id -eq $Card.boardparent}
        $Project = $KanbanizeProjedctsAndBoards.projects| where {$_.Boards.contains($Board)}

        Invoke-TrackITLogin -Username helpdeskbot -Pwd helpdeskbot
        Edit-TervisTrackITWorkOrder -WorkOrderNumber $Card.TrackITID -KanbanizeCardID $Card.taskid
    }
    <#
    foreach ($WorkOrder in $WorkOrdersWithOutKanbanizeIDs) {
        $CardName = "" + $WorkOrder.Wo_Num + " -  " + $WorkOrder.Task 
        New-KanbanizeTask -BoardID $TriageProcessBoardID -Title $CardName -CustomFields @{"trackitid"=$WorkOrder.Wo_Num;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w$($WorkOrder.Wo_Num)"} -Column $TriageProcessStartingColumn -Lane "Planned Work"
    }

    #>
    #$WorkOrdersWithOutKanbanizeIDs | group KanbanizeProjectBasedOnAssignedTechnician
    
    $WorkOrdersThatShouldLikelyBeProcessedByTechnicalServices = $WorkOrdersWithOutKanbanizeIDs | 
        group KanbanizeProjectBasedOnAssignedTechnician | 
        where name -Match "Technical Services" | 
        select -ExpandProperty group

    #$WorkOrdersThatShouldLikelyBeProcessedByTechnicalServices | group respons

    foreach ($WorkOrder in $WorkOrdersThatShouldLikelyBeProcessedByTechnicalServices) {
        $ExistingCardBasedOnTrackITIDInTitleOfCard = $Cards | where TrackITIDFromTitle -EQ $WorkOrder.WOID
        if ($ExistingCardBasedOnTrackITIDInTitleOfCard){
            $ExistingCardBasedOnTrackITIDInTitleOfCard.TrackITIDFromTitle
            Edit-KanbanizeTask -BoardID $ExistingCardBasedOnTrackITIDInTitleOfCard.boardparent -TaskID $ExistingCardBasedOnTrackITIDInTitleOfCard.taskid -CustomFields @{"trackitid"=$WorkOrder.Wo_Num;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w$($WorkOrder.Wo_Num)"}
        }
    }

    foreach ($WorkOrder in $WorkOrdersThatShouldLikelyBeProcessedByTechnicalServices) {
        $ExistingCardBasedOnTitleFormat = $Cards | where title -eq $WorkOrder.TitleInKanbanizeCardFormat
        if ($ExistingCardBasedOnTitleFormat){
            $ExistingCardBasedOnTitleFormat.TrackITIDFromTitle
            Edit-KanbanizeTask -BoardID $ExistingCardBasedOnTitleFormat.boardparent -TaskID $ExistingCardBasedOnTitleFormat.taskid -CustomFields @{"trackitid"=$WorkOrder.Wo_Num;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w$($WorkOrder.Wo_Num)"}
        }
    }


    
    foreach ($WorkOrder in $WorkOrdersThatShouldLikelyBeProcessedByTechnicalServices) {
        New-KanbanizeTask -BoardID 49 -Title $WorkOrder.TitleInKanbanizeCardFormat -CustomFields @{"trackitid"=$WorkOrder.Wo_Num;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w$($WorkOrder.Wo_Num)"} -Column "Requested" -Assignee $WorkOrder.Respons
    }



    $WorkOrdersWithOutKanbanizeIDsForItemManagement = $WorkOrdersWithOutKanbanizeIDs | 
        group KanbanizeProjectBasedOnAssignedTechnician | 
        where name -Match "Item Management" | 
        select -ExpandProperty group

    foreach ($WorkOrder in $WorkOrdersWithOutKanbanizeIDsForItemManagement) {
        New-KanbanizeTask -BoardID 52 -Title $WorkOrder.TitleInKanbanizeCardFormat -CustomFields @{"trackitid"=$WorkOrder.Wo_Num;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w$($WorkOrder.Wo_Num)"} -Column "Requested" -Assignee $WorkOrder.Respons
    }

    $WorkOrdersWithoutKanbanizeIDsAssignedToBacklog = $WorkOrdersWithOutKanbanizeIDs | 
        group respons | 
        where name -Match "Backlog" | 
        select -ExpandProperty group

    $WorkOrdersWithoutKanbanizeIDsAssignedToBacklog | where type -Match "Technical Services"


    <# New strategy, filter out everything that isn't possibly in scope for Technical Services
    $WorkOrdersThatShouldLikelyBeProcessedByBusinessServices = $WorkOrdersWithOutKanbanizeIDs | 
        group KanbanizeProjectBasedOnAssignedTechnician | 
        where name -Match "Business Services" | 
        select -ExpandProperty group

    foreach ($WorkOrder in $WorkOrdersThatShouldLikelyBeProcessedByTechnicalServices) {
        $ExistingCardBasedOnTrackITIDInTitleOfCard = $Cards | where TrackITIDFromTitle -EQ $WorkOrder.WOID
        if ($ExistingCardBasedOnTrackITIDInTitleOfCard){
            $ExistingCardBasedOnTrackITIDInTitleOfCard.TrackITIDFromTitle
            Edit-KanbanizeTask -BoardID $ExistingCardBasedOnTrackITIDInTitleOfCard.boardparent -TaskID $ExistingCardBasedOnTrackITIDInTitleOfCard.taskid -CustomFields @{"trackitid"=$WorkOrder.Wo_Num;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w$($WorkOrder.Wo_Num)"}
        }
    }
    #>

}
Sync-TrackITAndKanbanize

#$Cards | where title -eq "51118 - Provide a server as a dedicated FTP sending resource for Customizer"
#Edit-KanbanizeTask -BoardID 8 -TaskID 1012 -CustomFields @{"trackitid"=51118;"trackiturl"="http://trackit/TTHelpdesk/Application/Main?tabs=w51118"}