<?php

$groupGuid = (int) get_input("groupGuid");

$csvExportString = generate_group_events_spreadsheet($groupGuid);
header("Content-type: text/csv");
header("Content-Disposition: Attachment; filename=export.csv");
header('Pragma: public');
echo $csvExportString;
exit;

/*
$csvExportString = getGroupEventXML();
header("Content-type: text/xml");
header("Content-Disposition: Attachment; filename=export.xml");
header('Pragma: public');
echo $csvExportString;
exit;
*/
