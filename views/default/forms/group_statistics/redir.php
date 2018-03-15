<?php
/**
 * Created by PhpStorm.
 * User: Irin A
 * Date: 11/10/2017
 * Time: 11:37 AM
 */

$subtypeSelected = $vars['subtype'];
$yearSelected = $vars['year'];
$groupGuid = $vars['groupGuid'];
$validSubtypeArray = $vars['validSubtypeArray'];

$groupCreateDate = get_entity($groupGuid)->time_created;
//echo date("F d Y",$groupCreateDate);
?>

<div class="row">
    <div class="col col-xs-12 ">
        <div class="dashNav" style="text-align: center">

            <div
                style="margin-top:5px;display: inline-block;
vertical-align: middle;text-align: center">
                <?php echo elgg_echo('group_statistics:showMe');?>
                <select style="background-color: white" name="subtype" id="subtypeDropdown">
                    <?php
                    foreach ($validSubtypeArray as $subtype){

                        $prettifiedSubtype = elgg_echo("item:object:" . $subtype);
                        $selected = ($subtype==$subtypeSelected)? "selected" : "";
                        ?>

                        <option value="<?php  echo $subtype;?>" <?php echo $selected;?> ><?php echo $prettifiedSubtype; ?></option>

                        <?php

                    }

                    ?>
                </select>

                <?php echo elgg_echo('group_statistics:statisticsFrom');?>


            </div>
            <?php echo elgg_view('input/submit', array('name' => 'submit', 'value' => elgg_echo('group_statistics:go')));
            ?>

        </div>
    </div>

</div>
