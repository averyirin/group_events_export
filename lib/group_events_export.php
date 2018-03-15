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

    $event_options = array();
    $event_options["container_guid"] =$groupGuid;
    $events = event_manager_search_events($event_options);
	  $entities = $events["entities"];
    $eventGuids = array();
    $csvExportString = "";

    foreach ($entities as $event) {
      array_push($eventGuids, $event->guid);
      $csvExportString .= event_manager_export_attendees_original($event, true);
    }
    //echo event_manager_export_attendees_original($event, true);



    $return['title'] = "Group Events Export";
    $return['content'] .= elgg_view('group_events_export/list_events',
        array('listEvents' => $csvExportString));
    return $return;
}


function event_manager_export_attendees_original($event, $file = false) {
		$old_ia = elgg_get_ignore_access();
		elgg_set_ignore_access(true);

		if($file) {
			$EOL = "\r\n";
		} else {
			$EOL = PHP_EOL;
		}

		$headerString .= '"'.elgg_echo('guid').'";"'.elgg_echo('name').'";"'.elgg_echo('email').'";"'.elgg_echo('username').'"';

		if($event->registration_needed) {
			if($registration_form = $event->getRegistrationFormQuestions()) {
				foreach($registration_form as $question) {
					$headerString .= ';"'.$question->title.'"';
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

				$dataString .= '"'.$attendee->guid.'";"'.$attendee->name.'";"'.$attendee->email.'";"'.$attendee->username.'"';

				if($event->registration_needed) {
					if($registration_form = $event->getRegistrationFormQuestions()) {
						foreach($registration_form as $question) {
							$answer = $question->getAnswerFromUser($attendee->getGUID());

							$answerString .= '"'.addslashes($answer->value).'";';
						}
					}
					$dataString .= ';'.substr($answerString, 0, (strlen($answerString) -1));
				}

				if($event->with_program) {
					if($eventDays = $event->getEventDays()) {
						foreach($eventDays as $eventDay) {
							if($eventSlots = $eventDay->getEventSlots()) {
								foreach($eventSlots as $eventSlot) {
									if(check_entity_relationship($attendee->getGUID(), EVENT_MANAGER_RELATION_SLOT_REGISTRATION, $eventSlot->getGUID())) {
										$dataString .= ';"V"';
									} else {
										$dataString .= ';""';
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
