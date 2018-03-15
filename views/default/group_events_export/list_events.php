<?php
$listEvents = $vars['listEvents'];
?>
<div>

    <div  class="col col-xs-12">
      <textarea name="csv" rows="8" cols="80"><?php echo $listEvents;?></textarea>
    </div>
<?php
 elgg_view("output/url",
 array("is_action" => true,
 "href" => "action/group_events_export/group_events_export/csv?groupGuid=1742051",
 "text" => "Export to CSV"));
?>

</div>
