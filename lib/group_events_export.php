<?php
/**
 * Created by PhpStorm.
 * User: Irin A
 * Date: 11/21/2017
 * Time: 10:17 AM
 */
function generate_export_spreadsheet($resultEventGuids, $groupGuid){
      //Create excel formatted xml spreadsheet
      $spreadsheetExportString = '<?xml version="1.0"?>
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
      <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
      </Borders>
      <Font ss:Color="#0000FF" ss:Underline="Single"/>
      </Style>
      <Style ss:ID="s23">
      <Font x:Family="Swiss" ss:Bold="1"/>
      </Style>
      <Style ss:ID="s26">
       <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
       </Borders>
      <NumberFormat ss:Format="General Date"/>
     </Style>
     <Style ss:ID="s27">
      <Borders>
       <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
       <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
      </Borders>
      <NumberFormat ss:Format="[$-1009]d\-mmm\-yy;@"/>
    </Style>
      <Style ss:ID="s28">
       <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
       <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="16" ss:Color="#FFFFFF"
        ss:Bold="1"/>
       <Interior ss:Color="#000000" ss:Pattern="Solid"/>
      </Style>
      <Style ss:ID="s29">
       <Alignment ss:Vertical="Bottom"/>
       <Font ss:FontName="Arial" x:Family="Swiss" ss:Color="#FFFFFF" ss:Bold="1"/>
       <Interior ss:Color="#808080" ss:Pattern="Solid"/>
      </Style>
      <Style ss:ID="s30">
       <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
       </Borders>
      </Style>
      <Style ss:ID="s31">
       <Alignment ss:Horizontal="Center" ss:Vertical="Bottom"/>
       <Font ss:FontName="Arial" x:Family="Swiss" ss:Color="#FFFFFF" ss:Bold="1"/>
       <Interior ss:Color="#808080" ss:Pattern="Solid"/>
      </Style>
      <Style ss:ID="s32">
       <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
       <Font ss:FontName="Arial" x:Family="Swiss" ss:Size="11" ss:Color="#FFFFFF"
        ss:Bold="0"/>
       <Interior ss:Color="#000000" ss:Pattern="Solid"/>
      </Style>
      </Styles>
      ';
      //rows is number of events plus header filter row plus top heading
      $overviewRowTotal = count($resultEventGuids)+2;
      //col is number of data cols
      $overviewColTotal = 5;

      //get name of group
      $nameOfGroup = get_entity($groupGuid)->name;
      //Set Overview Headers
      $overviewHeaderRow .= '
       <Row ss:StyleID="s23">
       <Cell ss:StyleID="s29"><Data ss:Type="String">Event</Data></Cell>
       <Cell ss:StyleID="s29"><Data ss:Type="String">Location</Data></Cell>
       <Cell ss:StyleID="s29"><Data ss:Type="String">Venue</Data></Cell>
       <Cell ss:StyleID="s29"><Data ss:Type="String">Start</Data></Cell>
       <Cell ss:StyleID="s29"><Data ss:Type="String">End</Data></Cell>
       ';
       $event_relationship_options = event_manager_event_get_relationship_options();
       reset($event_relationship_options);
       foreach($event_relationship_options as $relationship) {
         //Add types of attendance header (attended/interested/organizing/exhibiting)
         $overviewColTotal++;
         $overviewHeaderRow .= '<Cell ss:StyleID="s29"><Data ss:Type="String">'.ucfirst(substr($relationship,6)).'</Data></Cell>';
      }
      $overviewHeaderRow .= '</Row>';
      $overviewHeaderTitle = '
       <Row>
       <Cell ss:MergeAcross="'.($overviewColTotal-1).'" ss:StyleID="s28"><Data ss:Type="String">'.$nameOfGroup.' Overview</Data></Cell>
       </Row>';



      //Create Overview Sheet
      $overviewSheet = '
      <Worksheet ss:Name="'."Overview".'">
      <Names>
       <NamedRange ss:Name="_FilterDatabase" ss:RefersTo="=Overview!R2C1:R'.$overviewRowTotal.'C'.$overviewColTotal.'"
        ss:Hidden="1"/>
      </Names>
      <Table
      x:FullColumns="1"
      x:FullRows="1">
      <Column ss:AutoFitWidth="1"  />';

      //add overview header information with filter
      $spreadsheetExportString .= $overviewSheet .$overviewHeaderTitle. $overviewHeaderRow;

      //Populate Overview Data
      foreach ($resultEventGuids as $eventGuid) {
        $xml = group_events_export_overview(get_entity($eventGuid));
        $spreadsheetExportString .= $xml;
      }
      //End overview Spreadsheet
      $spreadsheetExportString .= '</Table>
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
      <AutoFilter x:Range="R2C1:R'.$overviewRowTotal.'C'.$overviewColTotal.'"
       xmlns="urn:schemas-microsoft-com:office:excel">
      </AutoFilter>
      </Worksheet>';
      //Create Individual Event Sheet
      foreach ($resultEventGuids as $eventGuid) {
        $xml = group_events_export_sheet(get_entity($eventGuid));
        $spreadsheetExportString .= $xml;
      }
      $spreadsheetExportString .='
        </Workbook>
           ';
      return $spreadsheetExportString;
}

function group_events_export_overview($event){
  $old_ia = elgg_get_ignore_access();
  elgg_set_ignore_access(true);
   $eventDataXml = '<Row>
   <Cell ss:StyleID="s30"><Data ss:Type="String">'.(string)$event->title.'</Data></Cell>
   <Cell ss:StyleID="s30"><Data ss:Type="String">'.(string)$event->location.'</Data></Cell>
   <Cell ss:StyleID="s30"><Data ss:Type="String">'.(string)$event->venue.'</Data></Cell>
   <Cell ss:StyleID="s27"><Data ss:Type="DateTime">'.(string)date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY, $event->start_day) . "T". date('H', $event->start_time) . ':' . date('i', $event->start_time).'</Data></Cell>
   <Cell ss:StyleID="s27"><Data ss:Type="DateTime">'.(string)date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY, $event->end_ts) . "T". date('H', $event->end_ts) . ':' . date('i', $event->end_ts) .'</Data></Cell>
   ';
   $event_relationship_options = event_manager_event_get_relationship_options();
   reset($event_relationship_options);
   foreach($event_relationship_options as $relationship) {
     //Add types of attendance header (attended/interested/organizing/exhibiting)
       $old_ia = elgg_set_ignore_access(true);
       $peopleResponded = elgg_get_entities_from_relationship(array(
         'relationship' => $relationship,
         'relationship_guid' => $event->getGUID(),
         'inverse_relationship' => FALSE,
         'site_guids' => false,
         'limit' => false
       ));
       //Add number of people who are each attendance type
       $eventDataXml .='<Cell ss:StyleID="s30"><Data ss:Type="Number">'.(int)count($peopleResponded).'</Data></Cell>';
   }
  $eventDataXml .= '</Row>';
  //return of event info
  return $eventDataXml;
}

//generates a sheet for each event with details
function group_events_export_sheet($event){
  $beginXml = '
  <Worksheet ss:Name="'.$event->title.'">
   <Table
   x:FullColumns="1"
   x:FullRows="1">
   <Column />';
   $endXml = '</Table>
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
   </Worksheet>';

    //tables
    $eventTable = getEventTable($event);
    $infoTable = getInfoTable($event);
    $contactTable = getContactTable($event);
    $descTable = getDescriptionTable($event);
    $activityTable = getActivityTable($event);
    $attendeeTable = getAttendeeTable($event);

    //return sheet of event info
    return $beginXml.$eventTable.$infoTable.$descTable.$activityTable.$attendeeTable.$endXml;
}


//generates the general event information table
function getEventTable($event){
  $eventColTotal = 5;
  $eventHeaderXml = '
   <Row ss:StyleID="s23">
   <Cell ss:StyleID="s29"><Data ss:Type="String">Event</Data></Cell>
   <Cell ss:StyleID="s29"><Data ss:Type="String">Location</Data></Cell>
   <Cell ss:StyleID="s29"><Data ss:Type="String">Venue</Data></Cell>
   <Cell ss:StyleID="s29"><Data ss:Type="String">Start</Data></Cell>
   <Cell ss:StyleID="s29"><Data ss:Type="String">End</Data></Cell>
   ';


   $eventGeneralHeaderXml = '
    <Row ss:StyleID="s23">
    <Cell ss:StyleID="s29"></Cell>
    <Cell ss:StyleID="s29"></Cell>
    <Cell ss:StyleID="s29"></Cell>
    <Cell ss:StyleID="s29"></Cell>
    <Cell ss:StyleID="s29"></Cell>
    <Cell ss:MergeAcross="6" ss:StyleID="s31"><Data ss:Type="String">Status</Data></Cell>
    </Row>';

    $eventDataXml = '<Row>
    <Cell ss:StyleID="s30"><Data ss:Type="String">'.(string)$event->title.'</Data></Cell>
    <Cell ss:StyleID="s30"><Data ss:Type="String">'.(string)$event->location.'</Data></Cell>
    <Cell ss:StyleID="s30"><Data ss:Type="String">'.(string)$event->venue.'</Data></Cell>
    <Cell ss:StyleID="s27"><Data ss:Type="DateTime">'.(string)date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY, $event->start_day) . "T". date('H', $event->start_time) . ':' . date('i', $event->start_time).'</Data></Cell>
    <Cell ss:StyleID="s27"><Data ss:Type="DateTime">'.(string)date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY, $event->end_ts) . "T". date('H', $event->end_ts) . ':' . date('i', $event->end_ts) .'</Data></Cell>
    ';
    $event_relationship_options = event_manager_event_get_relationship_options();
    foreach($event_relationship_options as $relationship) {
        //Add types of attendance header (attended/interested/organizing/exhibiting)
        //Add number of people who are each attendance type
        $peopleResponded = elgg_get_entities_from_relationship(array(
          'relationship' => $relationship,
          'relationship_guid' => $event->getGUID(),
          'inverse_relationship' => FALSE,
          'site_guids' => false,
          'limit' => false
        ));
        $eventColTotal ++;
        $eventHeaderXml .=  '<Cell ss:StyleID="s29"><Data ss:Type="String">'.ucfirst(substr($relationship,6)).'</Data></Cell>';
        $eventDataXml .='<Cell ss:StyleID="s30"><Data ss:Type="Number">'.(int)count($peopleResponded).'</Data></Cell>';
    }
    $eventHeaderXml .= '</Row>';
    $eventDataXml .= '</Row>';
    $eventHeaderTitle = '
     <Row>
     <Cell ss:MergeAcross="'.($eventColTotal-1).'" ss:StyleID="s28"><Data ss:Type="String">'.$event->title.' Overview</Data></Cell>
     </Row>';

     $eventShortDescTitle = '
      <Row ">
      <Cell ss:MergeAcross="'.($eventColTotal-1).'" ss:StyleID="s32"><Data ss:Type="String">'.$event->shortdescription.'</Data></Cell>
      </Row>';

     return $eventHeaderTitle.$eventShortDescTitle.$eventGeneralHeaderXml.$eventHeaderXml.$eventDataXml.'<Row></Row>';
}

//returns the complex description table generated from LP Table
function getDescriptionTable($event){
  //Description
  $descHeaderXml = '';
  $descDataXml = '';
  $descColTotal = 0;
  $descGeneralHeaderXml = '';
  $descHeaderTitle = '';

  $data = (string)($event->description);
   //Filter out the event description table into headers
   if($data != ""){
      $descHeaderXml .= '<Row>';
      $descDataXml .= '<Row>';
      $descGeneralHeaderXml = '<Row>';
     $dom = new DOMDocument();
     @$dom->loadHTML($data);
     $dom->preserveWhiteSpace = false;
     $xpath = new DOMXPath($dom);

     $results = $xpath->query('/html/body/table/tbody/tr');
     if($results->length > 0){
       //Found tables
       foreach ($results as $result){
         $cells = $result -> getElementsByTagName('td');
         $internalTables = $result -> getElementsByTagName('table');
         if($internalTables->length > 0){
            foreach ($internalTables as $it) {
               $icells = $it -> getElementsByTagName('tr');
               $iColTotal = 0;
               foreach($icells as $val){
                 $iDataCells = $val -> getElementsByTagName('td');
                 $descColTotal++;
                 $iColTotal++;
                 $descHeaderXml .=  '<Cell  ss:StyleID="s29"><Data ss:Type="String">'.($iDataCells->item(0)->nodeValue).'</Data></Cell>';
                 $descDataXml .='<Cell ss:StyleID="s30"><Data ss:Type="String">'.($iDataCells->item(1)->nodeValue).'</Data></Cell>';
               }
               $descGeneralHeaderXml .= '<Cell ss:MergeAcross="'.($iColTotal-1).'" ss:StyleID="s31"><Data ss:Type="String">'.($cells->item(0)->nodeValue).'</Data></Cell>';
           }
          }else{
            $descColTotal++;
          $descGeneralHeaderXml .= '<Cell ss:StyleID="s29"></Cell>';
           $descHeaderXml .=  '<Cell  ss:StyleID="s29"><Data ss:Type="String">'.($cells->item(0)->nodeValue).'</Data></Cell>';
           $descDataXml .='<Cell ss:StyleID="s30"><Data ss:Type="String">'.($cells->item(1)->nodeValue).'</Data></Cell>';
         }
       }
     }else{
       //No tables in description
       $descColTotal++;
       $descGeneralHeaderXml .= '<Cell ss:StyleID="s29"></Cell>';
       $descHeaderXml .=  '<Cell  ss:StyleID="s29"><Data ss:Type="String">'."Description".'</Data></Cell>';
       $descDataXml .='<Cell ss:StyleID="s30"><Data ss:Type="String">'.$data.'</Data></Cell>';
     }
     //end description data
     $descHeaderXml .= '</Row>';
    $descDataXml .= '</Row>';
    $descGeneralHeaderXml .= '</Row>';
    $descHeaderTitle = '
     <Row ss:StyleID="s23">
    <Cell ss:MergeAcross="'.($descColTotal-1).'" ss:StyleID="s28"><Data ss:Type="String">Description</Data></Cell>
    </Row>';
  }
  //optional desc table spacing
  $descTable = $descHeaderTitle.$descGeneralHeaderXml.$descHeaderXml.$descDataXml;
  if($descTable != ""){
       $descTable .= '<Row></Row>';
  }

  return $descTable;
}


function getInfoTable($event){
        //Info and Contact Table
        $infoHeaderXml = '
        <Row>
        <Cell ss:StyleID="s29"><Data ss:Type="String">registration_ended</Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String">notify_onsignup</Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String">max_attendees</Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String">waiting_list</Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String">twitter_hash</Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String">region</Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String">website</Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String">event_type</Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String">contact</Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String">fee</Data></Cell>
        </Row>';

        /*
          $event->shortdescription = $shortdescription;
          $event->comments_on = $comments_on;
          $event->registration_ended = $registration_ended;
          $event->registration_needed = $registration_needed;
          $event->show_attendees = $show_attendees;
          $event->hide_owner_block = $hide_owner_block;
          $event->notify_onsignup = $notify_onsignup;
          $event->max_attendees = $max_attendees;
          $event->waiting_list = $waiting_list;
          $event->twitter_hash = $twitter_hash;
          $event->region = $region;
          $event->website = $website;
          $event->event_type = $event_type;
          $event->organizer = $organizer;
          $event->fee = $fee;

          <Cell ss:StyleID="s27"><Data ss:Type="DateTime">'.(string)date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY,$eventDay->date) . "T". date('H', $eventSlot->start_time) . ':' . date('i', $eventSlot->start_time).'</Data></Cell>

          */
        $infoDataXml = '<Row>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->registration_ended.'</Data></Cell>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->notify_onsignup.'</Data></Cell>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->max_attendees.'</Data></Cell>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->waiting_list.'</Data></Cell>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->twitter_hash.'</Data></Cell>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->region.'</Data></Cell>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->website.'</Data></Cell>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->event_type.'</Data></Cell>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->contact_details.'</Data></Cell>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->fee.'</Data></Cell>
        </Row>';
        $infoColTotal = 11;
        $infoGeneralHeaderXml = '<Row><Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
        <Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell>
</Row>';
        $infoHeaderTitle = '<Row ss:StyleID="s23">
       <Cell ss:MergeAcross="'.($infoColTotal-1).'" ss:StyleID="s28"><Data ss:Type="String">Info</Data></Cell>
       </Row>';
       return $infoHeaderTitle. $infoGeneralHeaderXml.$infoHeaderXml.$infoDataXml.'<Row></Row>';

}

function getContactTable($event){
      //Contact Table
      $contactHeaderXml = '
      <Row><Cell ss:StyleID="s29"><Data ss:Type="String">Contact Details</Data></Cell></Row>';
      $contactDataXml = '<Row>
      <Cell ss:StyleID="s30"><Data ss:Type="String">'.$event->contact_details.'</Data></Cell>
      </Row>';
      $contactColTotal = 1;
      $contactGeneralHeaderXml = '<Row><Cell ss:StyleID="s29"><Data ss:Type="String"></Data></Cell></Row>';
      $contactHeaderTitle = '<Row ss:StyleID="s23">
     <Cell ss:MergeAcross="'.($contactColTotal-1).'" ss:StyleID="s28"><Data ss:Type="String">Contact</Data></Cell>
     </Row>';
     return $contactHeaderTitle. $contactGeneralHeaderXml.$contactHeaderXml.$contactDataXml.'<Row></Row>';
}

//get attendees responses and Status
function getAttendeeTable($event){
  //Default attendee data
  $attendeeHeaderXml = '
  <Row ss:StyleID="s23">
  <Cell ss:StyleID="s29"><Data ss:Type="String">Name</Data></Cell>
  <Cell ss:StyleID="s29"><Data ss:Type="String">Email</Data></Cell>
  <Cell ss:StyleID="s29"><Data ss:Type="String">Status</Data></Cell>
  ';

  $attendeeColTotal = 3;
  $attendeeHeaderTitle = '';

  //Build registration question headers
  if($event->registration_needed) {
    if($registration_form = $event->getRegistrationFormQuestions()) {
      foreach($registration_form as $question) {
        $attendeeColTotal++;
        $attendeeHeaderXml .= '<Cell ss:StyleID="s29"><Data ss:Type="String">'.$question->title.'</Data></Cell>';
      }
    }
  }
  //Build program headers with the events that the attendee can join
  if($event->with_program) {
    if($eventDays = $event->getEventDays()) {
      foreach($eventDays as $eventDay) {
        if($eventSlots = $eventDay->getEventSlots()) {
          foreach($eventSlots as $eventSlot) {
            $attendeeColTotal++;
            $attendeeHeaderXml .= '<Cell ss:StyleID="s29"><Data ss:Type="String">'.$eventSlot->title.'</Data></Cell>';
          }
        }
      }
    }
  }
  //End Fields
  $attendeeHeaderXml .= '</Row>';
  $attendeeDataXml = '';
  $event_relationship_options = event_manager_event_get_relationship_options();
  foreach($event_relationship_options as $relationship) {
      $peopleResponded = elgg_get_entities_from_relationship(array(
        'relationship' => $relationship,
        'relationship_guid' => $event->getGUID(),
        'inverse_relationship' => FALSE,
        'site_guids' => false,
        'limit' => false
      ));
      //add individual attendee status, their registration question responses, and chosen activities
      foreach ($peopleResponded as $attendee) {
        $attendeeDataXml .=  '<Row>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.(string)$attendee->name.'</Data></Cell>
        <Cell ss:StyleID="s21" ss:HRef="mailto:molly@katzen.com">
        <Data ss:Type="String">'.(string)$attendee->email.'</Data></Cell>
        <Cell ss:StyleID="s30"><Data ss:Type="String">'.(string)$relationship.'</Data></Cell>
        ';
        //Registration question answers
        $answerString = '';
        if($event->registration_needed) {
          if($registration_form = $event->getRegistrationFormQuestions()) {
            foreach($registration_form as $question) {
              $answer = $question->getAnswerFromUser($attendee->getGUID());
              $attendeeDataXml .= '<Cell ss:StyleID="s30"><Data ss:Type="String">'.($answer->value).'</Data></Cell>';
            }
          }
        }

        //[V] Checked - Joined a program within event
        if($event->with_program) {
          if($eventDays = $event->getEventDays()) {
            foreach($eventDays as $eventDay) {
              if($eventSlots = $eventDay->getEventSlots()) {
                foreach($eventSlots as $eventSlot) {
                  if(check_entity_relationship($attendee->getGUID(), EVENT_MANAGER_RELATION_SLOT_REGISTRATION, $eventSlot->getGUID())) {
                    $attendeeDataXml .= '<Cell ss:StyleID="s30"><Data ss:Type="String">'.'x'.'</Data></Cell>';
                  } else {
                    $attendeeDataXml .= '<Cell ss:StyleID="s30"><Data ss:Type="String"></Data></Cell>';
                  }
                }
              }
            }
          }
        }
        $attendeeDataXml .= '</Row>';
      }
  }
  $attendeeHeaderTitle = '
         <Row ss:StyleID="s23">
        <Cell ss:MergeAcross="'.($attendeeColTotal-1).'" ss:StyleID="s28"><Data ss:Type="String">Attendees</Data></Cell>
        </Row>';

  $attendeeTable = $attendeeHeaderTitle.$attendeeHeaderXml.$attendeeDataXml.'<Row></Row>';
  return $attendeeTable;
}

//get list of activities and total members attended them
function getActivityTable($event){
  $activityHeaderXml = '';
  $activityDataXml = '';
  $activityColTotal = 5;
  $activityHeaderTitle = '';

  //Build program headers with the events that the attendee can join
  if($event->with_program) {
    $activityHeaderXml .= '<Row ss:StyleID="s23">
    <Cell ss:StyleID="s29"><Data ss:Type="String">Name</Data></Cell>
    <Cell ss:StyleID="s29"><Data ss:Type="String">Description</Data></Cell>
    <Cell ss:StyleID="s29"><Data ss:Type="String">Start</Data></Cell>
    <Cell ss:StyleID="s29"><Data ss:Type="String">End</Data></Cell>
    <Cell ss:StyleID="s29"><Data ss:Type="String">Participants</Data></Cell>
    </Row>';
    if($eventDays = $event->getEventDays()) {
      foreach($eventDays as $eventDay) {
        $date = date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY, $eventDay->date);
        if($eventSlots = $eventDay->getEventSlots()) {
          foreach($eventSlots as $eventSlot) {
             $activityDataXml .= '<Row>
             <Cell ss:StyleID="s30"><Data ss:Type="String">'.$eventSlot->title.'</Data></Cell>
             <Cell ss:StyleID="s30"><Data ss:Type="String">'.$eventSlot->description.'</Data></Cell>
             <Cell ss:StyleID="s27"><Data ss:Type="DateTime">'.(string)date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY,$eventDay->date) . "T". date('H', $eventSlot->start_time) . ':' . date('i', $eventSlot->start_time).'</Data></Cell>
             <Cell ss:StyleID="s27"><Data ss:Type="DateTime">'.(string)date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY,$eventDay->date) . "T". date('H', $eventSlot->end_time) . ':' . date('i', $eventSlot->end_time) .'</Data></Cell>
             <Cell ss:StyleID="s30"><Data ss:Type="Number">'.(int)$eventSlot->countRegistrations().'</Data></Cell>
             </Row>';
          }
        }
      }
    }
    $activityHeaderTitle = '
     <Row ss:StyleID="s23">
    <Cell ss:MergeAcross="'.($activityColTotal-1).'" ss:StyleID="s28"><Data ss:Type="String">Activities</Data></Cell>
    </Row>';
  }
  //optional activity data table spacing
  $activityTable = $activityHeaderTitle.$activityHeaderXml.$activityDataXml;
  if($activityTable != ""){
       $activityTable .= '<Row></Row>';
  }
  return $activityTable;
}


function getGroupEventXMLOriginal(){
  $xml = '<?xml version="1.0"?>
  <?mso-application progid="Excel.Sheet"?>
  <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
   xmlns:o="urn:schemas-microsoft-com:office:office"
   xmlns:x="urn:schemas-microsoft-com:office:excel"
   xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
   xmlns:html="http://www.w3.org/TR/REC-html40">
   <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
    <Author>Irin Avery</Author>
    <LastAuthor>Irin A</LastAuthor>
    <Created>2005-08-02T04:06:26Z</Created>
    <LastSaved>2005-08-02T04:30:11Z</LastSaved>
    <Company>MPG LSC</Company>
    <Version>14.00</Version>
   </DocumentProperties>
   <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
    <AllowPNG/>
   </OfficeDocumentSettings>
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
     <Font ss:FontName="Arial"/>
     <Interior/>
     <NumberFormat/>
     <Protection/>
    </Style>
    <Style ss:ID="s62" ss:Name="Hyperlink">
     <Font ss:FontName="Arial" ss:Color="#0000FF" ss:Underline="Single"/>
    </Style>
    <Style ss:ID="s63">
     <Font ss:FontName="Arial" x:Family="Swiss" ss:Bold="1"/>
    </Style>
   </Styles>
   <Worksheet ss:Name="Sheet1">
    <Names>
     <NamedRange ss:Name="_FilterDatabase" ss:RefersTo="=Sheet1!R1C1:R2C4"
      ss:Hidden="1"/>
    </Names>
    <Table ss:ExpandedColumnCount="4" ss:ExpandedRowCount="2" x:FullColumns="1"
     x:FullRows="1">
     <Column ss:Index="4" ss:AutoFitWidth="0" ss:Width="154.5"/>
     <Row ss:StyleID="s63">
      <Cell><Data ss:Type="String">First</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
      <Cell><Data ss:Type="String">Middle</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
      <Cell><Data ss:Type="String">Last</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
      <Cell><Data ss:Type="String">Email</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
     </Row>
     <Row>
      <Cell><Data ss:Type="String">Molly</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
      <Cell ss:Index="3"><Data ss:Type="String">Katzen</Data><NamedCell
        ss:Name="_FilterDatabase"/></Cell>
      <Cell ss:StyleID="s62" ss:HRef="mailto:molly@katzen.com"><Data ss:Type="String">molly@katzen.com</Data><NamedCell
        ss:Name="_FilterDatabase"/></Cell>
     </Row>
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
     <Print>
      <ValidPrinterInfo/>
      <HorizontalResolution>300</HorizontalResolution>
      <VerticalResolution>300</VerticalResolution>
     </Print>
     <Selected/>
     <ProtectObjects>False</ProtectObjects>
     <ProtectScenarios>False</ProtectScenarios>
    </WorksheetOptions>
    <AutoFilter x:Range="R1C1:R2C4" xmlns="urn:schemas-microsoft-com:office:excel">
    </AutoFilter>
   </Worksheet>
   <Worksheet ss:Name="Sheet2">
    <Table ss:ExpandedColumnCount="1" ss:ExpandedRowCount="1" x:FullColumns="1"
     x:FullRows="1">
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
     <ProtectObjects>False</ProtectObjects>
     <ProtectScenarios>False</ProtectScenarios>
    </WorksheetOptions>
   </Worksheet>
   <Worksheet ss:Name="Sheet3">
    <Table ss:ExpandedColumnCount="1" ss:ExpandedRowCount="1" x:FullColumns="1"
     x:FullRows="1">
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
     <ProtectObjects>False</ProtectObjects>
     <ProtectScenarios>False</ProtectScenarios>
    </WorksheetOptions>
   </Worksheet>
  </Workbook>
  ';
    return $xml;
}




?>
