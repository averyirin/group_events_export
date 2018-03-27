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
echo "Action----"."<br/>";
var_dump($_SESSION['eventHook']);
var_dump($_SESSION['eventEntityType']);
var_dump($_SESSION['eventValue']);
var_dump($_SESSION['eventParams']);
exit();
/*
$spreadsheetExportString = generate_export_spreadsheet($groupGuid);
header("Content-type: text/xml");
header("Content-Disposition: Attachment; filename=export.xml");
header('Pragma: public');
echo $spreadsheetExportString;
exit;
*/
