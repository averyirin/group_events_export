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
		$headerString .= '"'.$event->title.'","'.elgg_echo('guid').'","'.elgg_echo('name').'","'.elgg_echo('email').'","'.elgg_echo('username').'","'.elgg_echo('Relationship').'"';


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

							$headerString .= ';"Event activity: \''.$eventSlot->title.'\' '.$date. ' ('.$start_time_hour.':'.$start_time_minutes.' - '.$end_time_hour.':'.$end_time_minutes.')"';
						}
					}
				}
			}
		}


    //Loop Through relationship options
    $event_relationship_options = event_manager_event_get_relationship_options();
    reset($event_relationship_options);
    foreach($event_relationship_options as $relationship) {
      if($relationship == EVENT_MANAGER_RELATION_ATTENDING){
                $dataString .= "Found ".$relationship;
      }

      $dataString .= "Looking for ".$relationship;

      if($event->$relationship){

            				$dataString .= " Has ".$relationship;
            				$dataString .= $EOL;

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

    				$dataString .= '"'.$attendee->guid.'","'.$attendee->name.'","'.$attendee->email.'","'.$attendee->username.'","'.$relationship.'"';

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

    				$dataString .= $EOL;
    			}
    		}

      }




    }



		$headerString .= $EOL;
		elgg_set_ignore_access($old_ia);

		return $headerString . $dataString;
	}


function event_manager_export_attendees_comma($event, $file = false) {
		$old_ia = elgg_get_ignore_access();
		elgg_set_ignore_access(true);

		if($file) {
			$EOL = "\r\n";
		} else {
			$EOL = PHP_EOL;
		}

		$headerString .= '"'.elgg_echo('guid').'","'.elgg_echo('name').'","'.elgg_echo('email').'","'.elgg_echo('username').'"';

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

							$headerString .= ';"Event activity: \''.$eventSlot->title.'\' '.$date. ' ('.$start_time_hour.':'.$start_time_minutes.' - '.$end_time_hour.':'.$end_time_minutes.')"';
						}
					}
				}
			}
		}

		if($attendees = $event->exportAttendees()) {
			foreach($attendees as $attendee) {
				$answerString = '';

				$dataString .= '"'.$attendee->guid.'","'.$attendee->name.'","'.$attendee->email.'","'.$attendee->username.'"';

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

				$dataString .= $EOL;
			}
		}

		$headerString .= $EOL;
		elgg_set_ignore_access($old_ia);

		return $headerString . $dataString;
	}

?>
