<?php
/**
 * Created by PhpStorm.
 * User: Irin A
 * Date: 11/21/2017
 * Time: 10:17 AM
 */

function generate_group_events_spreadsheet($groupGuid = NULL){

  $event_options = array();
  $event_options["container_guid"] =$groupGuid;
  $events = event_manager_search_events($event_options);
  $entities = $events["entities"];
  $eventGuids = array();
  $csvExportString = "";
  $headerString .= '"'."Event Title".'","'.elgg_echo('name').'","'.elgg_echo('email').'","'.elgg_echo('Relationship').'"'."\r\n";
  $csvExportString.=$headerString;

  foreach ($entities as $event) {
    array_push($eventGuids, $event->guid);
    $csvExportString .= group_events_export_comma($event);
  }
  return $csvExportString;
}

function group_events_export_comma($event) {
		$old_ia = elgg_get_ignore_access();
		elgg_set_ignore_access(true);
    $EOL = "\r\n";
/*
    //To do, see what register event and with program are needed
		if($event->registration_needed) {
			if($registration_form = $event->getRegistrationFormQuestions()) {
				foreach($registration_form as $question) {
					$headerString .= ',"'.$question->title.'"';
				}
			}
		}

		if($event->with_program) {
			if($eventDays = $event->getEventDays()) {
				foreach($eventDays as $eventDay) {
					$date = date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY, $eventDay->date);
					if($eventSlots = $eventDay->getEventSlots()) {
						foreach($eventSlots as $eventSlot) {
							$start_time = $eventSlot->start_time;
							$end_time = $eventSlot->end_time;

							$start_time_hour = date('H', $start_time);
							$start_time_minutes = date('i', $start_time);

							$end_time_hour = date('H', $end_time);
							$end_time_minutes = date('i', $end_time);

							$headerString .= ',"Event activity: \''.$eventSlot->title.'\' '.$date. ' ('.$start_time_hour.':'.$start_time_minutes.' - '.$end_time_hour.':'.$end_time_minutes.')"';
						}
					}
				}
			}
		}
*/

    //Loop Through relationship options
    $event_relationship_options = event_manager_event_get_relationship_options();
    reset($event_relationship_options);
    foreach($event_relationship_options as $relationship) {
        $old_ia = elgg_set_ignore_access(true);
        $peopleResponded = elgg_get_entities_from_relationship(array(
          'relationship' => $relationship,
          'relationship_guid' => $event->getGUID(),
          'inverse_relationship' => FALSE,
          'site_guids' => false,
          'limit' => false
        ));
        elgg_set_ignore_access($old_ia);

        if($peopleResponded) {
          reset($peopleResponded);
    			foreach($peopleResponded as $attendee) {
    				$answerString = '';

//time_updated

    				$dataString .= '"'.$event->title.'","'.$attendee->name.'","'.$attendee->email.'","'.$relationship.'","'.$attendee->time_created.'"';
/*
            if($event->registration_needed) {
    					if($registration_form = $event->getRegistrationFormQuestions()) {
    						foreach($registration_form as $question) {
    							$answer = $question->getAnswerFromUser($attendee->getGUID());

    							$answerString .= '"'.addslashes($answer->value).'",';
    						}
    					}
    					$dataString .= ','.substr($answerString, 0, (strlen($answerString) -1));
    				}

    				if($event->with_program) {
    					if($eventDays = $event->getEventDays()) {
    						foreach($eventDays as $eventDay) {
    							if($eventSlots = $eventDay->getEventSlots()) {
    								foreach($eventSlots as $eventSlot) {
    									if(check_entity_relationship($attendee->getGUID(), EVENT_MANAGER_RELATION_SLOT_REGISTRATION, $eventSlot->getGUID())) {
    										$dataString .= ',"V"';
    									} else {
    										$dataString .= ',""';
    									}
    								}
    							}
    						}
    					}
    				}
*/
    				$dataString .= $EOL;
    			}
    		}
    }
		elgg_set_ignore_access($old_ia);
		return $dataString;
	}
function getGroupEventXML(){
  $xml = '<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>Jack Herrington</Author>
  <LastAuthor>Jack Herrington</LastAuthor>
  <Created>2005-08-02T04:06:26Z</Created>
  <LastSaved>2005-08-02T04:30:11Z</LastSaved>
  <Company>My Software Company, Inc.</Company>
  <Version>11.6360</Version>
  </DocumentProperties>
  <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>8535</WindowHeight>
  <WindowWidth>12345</WindowWidth>
  <WindowTopX>480</WindowTopX>
  <WindowTopY>90</WindowTopY>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
  </ExcelWorkbook>
  <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
  <Alignment ss:Vertical="Bottom"/>
  <Borders/>
  <Font/>
  <Interior/>
  <NumberFormat/>
  <Protection/>
  </Style>
  <Style ss:ID="s21" ss:Name="Hyperlink">
  <Font ss:Color="#0000FF" ss:Underline="Single"/>
  </Style>
  <Style ss:ID="s23">
  <Font x:Family="Swiss" ss:Bold="1"/>
  </Style>
  </Styles>
  <Worksheet ss:Name="Sheet1">
  <Table ss:ExpandedColumnCount="4"
  ss:ExpandedRowCount="5" x:FullColumns="1"
  x:FullRows="1">
  <Column ss:Index="4" ss:AutoFitWidth="0" ss:Width="154.5"/>
  <Row ss:StyleID="s23">
  <Cell><Data ss:Type="String">First</Data></Cell>
  <Cell><Data ss:Type="String">Middle</Data></Cell>
  <Cell><Data ss:Type="String">Last</Data></Cell>
  <Cell><Data ss:Type="String">Email</Data></Cell>
  </Row>
  <Row>
  <Cell><Data ss:Type="String">Molly</Data></Cell>
  <Cell ss:Index="3"><Data
  ss:Type="String">Katzen</Data></Cell>
  <Cell ss:StyleID="s21" ss:HRef="mailto:molly@katzen.com">
  <Data ss:Type="String">molly@katzen.com</Data></Cell>
  </Row>
  ...
  </Table>
  <WorksheetOptions
  xmlns="urn:schemas-microsoft-com:office:excel">
  <Print>
  <ValidPrinterInfo/>
  <HorizontalResolution>300</HorizontalResolution>
  <VerticalResolution>300</VerticalResolution>
  </Print>
  <Selected/>
  <Panes>
  <Pane>
  <Number>3</Number>
  <ActiveRow>5</ActiveRow>
  </Pane>
  </Panes>
  <ProtectObjects>False</ProtectObjects>
  <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
  </Worksheet>
  <Worksheet ss:Name="Sheet2">
  <WorksheetOptions
  xmlns="urn:schemas-microsoft-com:office:excel">
  <ProtectObjects>False</ProtectObjects>
  <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
  </Worksheet>
  <Worksheet ss:Name="Sheet3">
  <WorksheetOptions
  xmlns="urn:schemas-microsoft-com:office:excel">
  <ProtectObjects>False</ProtectObjects>
  <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
  </Worksheet>
  </Workbook>';
  return $xml;
}



?>
