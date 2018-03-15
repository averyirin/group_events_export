<?php


$subtype = filter_var(get_input('subtype'),FILTER_SANITIZE_STRING);
$groupGuid = filter_var(get_input('groupGuid'),FILTER_SANITIZE_STRING);

$_SESSION['subtype'] = $subtype;
if($groupGuid != NULL){
    forward("/group_statistics/dashboard/".$groupGuid);
}

