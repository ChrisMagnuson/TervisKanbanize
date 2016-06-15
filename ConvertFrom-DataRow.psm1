

$WorkOrdersAsDataTable = get-TrackITWorkOrders
$WorkOrders = $WorkOrdersAsDataTable | ConvertFrom-DataRow
