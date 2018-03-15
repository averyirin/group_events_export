<?php
/**
 * Author: Irin Avery
 * Desc: Start file for group_statistics plugin
 */
//redirect action handler
$action_path = elgg_get_plugins_path() . 'group_events_export/actions/group_events_export';
//register redirection and forms
elgg_register_action("group_events_export/csv", $action_path . "/csv.php");
elgg_register_action("subtypeForm", $action_path . "/subtypeForm.php");
//register plugin init
elgg_register_event_handler('init', 'system', 'group_events_export_init');


//when the plugin is active
function group_events_export_init()
{
    //page setup handler to add the group export events button
    elgg_register_event_handler("pagesetup", "system", "group_events_export_pagesetup", 550);
    //register group_statistics library
    elgg_register_library('group_events_export:lib', elgg_get_plugins_path() . 'group_events_export/lib/group_events_export.php');
    //register page handler
    //load library
    elgg_load_library('group_events_export:lib');


    //if we are in the group statistics register libraries
    if (get_context() == 'group_events_export') {
        elgg_register_page_handler('group_events_export', 'group_events_export_page_handler');
    }
}


//adds the group statistics button for group admins
function group_events_export_pagesetup()
{
    //if we can edit the group
    $page_owner = elgg_get_page_owner_entity();
    //add the group statistics button to the group admin menu
    if(elgg_in_context("events")&& ($page_owner instanceof ElggGroup)) {
      elgg_register_menu_item('title', array(
								'name' => "export",
								'href' => "action/group_events_export/csv?groupGuid=1742051",
								'text' => "Export Group Events",
								'link_class' => 'elgg-button elgg-button-action',
                'is_action' => true
								));
    }
}


/**
 *Directs group admin to their statistics page
 * group_statistics/dashboard/Group ID/Item ID
 *
 * @param array $page
 * @return bool
 */
function group_events_export_page_handler($page)
{

    //set page, groupID and itemID
    $pageNav = $page[0];

    $groupID = $page[1];
    //Looking for
    //group_statistics/dashboard/GroupID/ItemId
    switch ($pageNav) {
        case "export_events":
            //if we have itemID, go to item view with back btn
            if ($groupID!= NULL) {
                $params = get_events_from_group($groupID);
                $body = elgg_view_layout('admin', $params);
            }
            break;
        default:
            return false;
    }

    //show the created page
    echo elgg_view_page($params['title'], $body);
    return true;
}
