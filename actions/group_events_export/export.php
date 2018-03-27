<?php

$groupGuid = (int) get_input("groupGuid");
/*
$csvExportString = generate_group_events_csv($groupGuid);
header("Content-type: text/csv");
header("Content-Disposition: Attachment; filename=export.csv");
header('Pragma: public');
echo $csvExportString;
exit;
*/
$resultEntities = $_SESSION['eventEntities'];
$spreadsheetExportString = generate_export_spreadsheet($resultEntities);
header("Content-type: text/xml");
header("Content-Disposition: Attachment; filename=export.xml");
header('Pragma: public');
echo $spreadsheetExportString;
exit;
