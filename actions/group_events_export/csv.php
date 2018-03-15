<?php

$groupGuid = (int) get_input("groupGuid");
$csvExportString = generate_group_events_spreadsheet($groupGuid);
header("Content-type: application/vnd.ms-excel");
header("Content-Disposition: Attachment; filename=export.csv");
header('Pragma: public');
echo $csvExportString;
exit;
