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

$csvExportString = generate_export_xml($groupGuid);
header("Content-type: text/xml");
header("Content-Disposition: Attachment; filename=export.xml");
header('Pragma: public');
echo $csvExportString;
exit;
