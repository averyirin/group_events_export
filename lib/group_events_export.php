<?php
/**
 * Created by PhpStorm.
 * User: Irin A
 * Date: 11/21/2017
 * Time: 10:17 AM
 */

function get_events_from_group($groupGuid = NULL)
{

    $return = array();

    $listEvents = array();

    $return['title'] = "Group Events Export";
    $return['content'] .= elgg_view('group_events_export/list_events',
        array('listEvents' => $listEvents));
    return $return;
}

function get_events_by_group_guid($groupGuid)
{
    //Loop through subtypes of what content group has
    $groupSubtypes = elgg_get_entities(array(
        'type' => 'object',
        'container_guid' => $groupGuid,
        'group_by' => 'e.subtype',
        count => false,
        limit => 0
    ));

    $groupSubtypesClean = array();
    foreach ($groupSubtypes as $subtype) {
        array_push($groupSubtypesClean, get_subtype_from_id($subtype->subtype));
    }
    //filter out invalid subtypes
    $invalidSubtypeArray = array('widget', 'folder', 'task_top', 'messages', 'hjcategory');
    $validSubtypeArray = array_diff($groupSubtypesClean, $invalidSubtypeArray);


    $subtypeToTop = 'page_top';
    //put the page at the top of the list
    if (in_array($subtypeToTop, $validSubtypeArray)) {
        $validSubtypeArray = array_diff($validSubtypeArray, [$subtypeToTop]);
        array_unshift($validSubtypeArray, $subtypeToTop);
    }

    return $validSubtypeArray;
}

?>
