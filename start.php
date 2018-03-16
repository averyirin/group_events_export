<?php
/**
 * Author: Irin Avery
 * Desc: Start file for group_statistics plugin
 */
//group events export action handler
$action_path = elgg_get_plugins_path() . 'group_events_export/actions/group_events_export';
//register export action
elgg_register_action("group_events_export/export", $action_path . "/export.php");
//register plugin init
elgg_register_event_handler('init', 'system', 'group_events_export_init');


//when the plugin is active
function group_events_export_init()
{


    //page setup handler to add the group export events button
    elgg_register_event_handler("pagesetup", "system", "group_events_export_pagesetup", 550);
    //register group_events_export library
    elgg_register_library('group_events_export:lib', elgg_get_plugins_path() . 'group_events_export/lib/group_events_export.php');
    //load library
    elgg_load_library('group_events_export:lib');
}


//adds the group statistics button for group admins
function group_events_export_pagesetup()
{
    //if we can edit the group
    $page_owner = elgg_get_page_owner_entity();
    $who_create_group_events = elgg_get_plugin_setting('who_create_group_events', 'event_manager'); // group_admin, members
    $user = elgg_get_logged_in_user_entity();
    //add the group statistics button to the group events menu
    //! Does this need security?
    if(elgg_in_context("events")&& ($page_owner instanceof ElggGroup)) {
      if((($who_create_group_events == "group_admin") && $page_owner->canEdit()) || (($who_create_group_events == "members") && $page_owner->isMember($user))){

      elgg_register_menu_item('title', array(
                'name' => "export",
                'href' => "action/group_events_export/export?groupGuid=".$page_owner->getGuid(),
                'text' => "Export Group Events",
                'link_class' => 'elgg-button elgg-button-action',
                'is_action' => true
                ));
    }

  }

}
