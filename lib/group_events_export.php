<?php
/**
 * Created by PhpStorm.
 * User: Irin A
 * Date: 11/21/2017
 * Time: 10:17 AM
 */
function generate_export_spreadsheet($resultEventGuids){
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
  <Font ss:Color="#0000FF" ss:Underline="Single"/>
  </Style>
  <Style ss:ID="s23">
  <Font x:Family="Swiss" ss:Bold="1"/>
  </Style>
  </Styles>
  ';
  //Create Overview Sheet


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


function group_events_export_sheet($event){
  $old_ia = elgg_get_ignore_access();
  elgg_set_ignore_access(true);
  $headerXml = '
   <Worksheet ss:Name="'.$event->title.'">
   <Table
   x:FullColumns="1"
   x:FullRows="1">
   <Column />
   <Row ss:StyleID="s23">
   <Cell><Data ss:Type="String">Event</Data></Cell>
   <Cell><Data ss:Type="String">Location</Data></Cell>
   <Cell><Data ss:Type="String">Venue</Data></Cell>
   <Cell><Data ss:Type="String">Start</Data></Cell>
   <Cell><Data ss:Type="String">End</Data></Cell>
   ';
   //Default attendee data
   $attendeeHeaderXml = '<Row></Row>
   <Row ss:StyleID="s23">
   <Cell><Data ss:Type="String">Name</Data></Cell>
   <Cell><Data ss:Type="String">Email</Data></Cell>
   <Cell><Data ss:Type="String">Status</Data></Cell>
   ';
   //Build registration question headers
   $headerString .= '"'.elgg_echo('name').'","'.elgg_echo('email').'","'.elgg_echo('Status').'"';
   if($event->registration_needed) {
     if($registration_form = $event->getRegistrationFormQuestions()) {
       foreach($registration_form as $question) {
         $attendeeHeaderXml .= '<Cell><Data ss:Type="String">'.$question->title.'</Data></Cell>';
       }
     }
   }

   //Build program headers with the events that the attendee can join
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

             $attendeeHeaderXml .= '<Cell><Data ss:Type="String">Event activity: \''.$eventSlot->title.'\' '.$date. ' ('.$start_time_hour.':'.$start_time_minutes.' - '.$end_time_hour.':'.$end_time_minutes.')</Data></Cell>';
           }
         }
       }
     }
   }


   //End Fields
   $attendeeHeaderXml .= '</Row>';


   $attendeeDataXml = '';

   $eventXml = '<Row>
   <Cell><Data ss:Type="String">'.(string)$event->title.'</Data></Cell>
   <Cell><Data ss:Type="String">'.(string)$event->location.'</Data></Cell>
   <Cell><Data ss:Type="String">'.(string)$event->venue.'</Data></Cell>
   <Cell><Data ss:Type="String">'.(string)date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY, $event->start_day) . " ". date('H', $event->start_time) . ':' . date('i', $event->start_time).'</Data></Cell>
   <Cell><Data ss:Type="String">'.date(EVENT_MANAGER_FORMAT_DATE_EVENTDAY, $event->end_ts) . " ". date('H', $event->end_ts) . ':' . date('i', $event->end_ts) .'</Data></Cell>
   ';




   $event_relationship_options = event_manager_event_get_relationship_options();
   reset($event_relationship_options);
   foreach($event_relationship_options as $relationship) {
     //Add types of attendance header (attended/interested/organizing/exhibiting)
      $headerXml .=  '<Cell><Data ss:Type="String">'.$relationship.'</Data></Cell>';
       $old_ia = elgg_set_ignore_access(true);
       $peopleResponded = elgg_get_entities_from_relationship(array(
         'relationship' => $relationship,
         'relationship_guid' => $event->getGUID(),
         'inverse_relationship' => FALSE,
         'site_guids' => false,
         'limit' => false
       ));
       //Add number of people who are each attendance type
       $eventXml .='<Cell><Data ss:Type="Number">'.(int)count($peopleResponded).'</Data></Cell>';

       //add individual attendee status, their registration question responses, and chosen activities
       foreach ($peopleResponded as $attendee) {
         $attendeeDataXml .=  '<Row>
         <Cell><Data ss:Type="String">'.(string)$attendee->name.'</Data></Cell>
         <Cell ss:StyleID="s21" ss:HRef="mailto:molly@katzen.com">
         <Data ss:Type="String">'.(string)$attendee->email.'</Data></Cell>
         <Cell><Data ss:Type="String">'.(string)$relationship.'</Data></Cell>
         ';
         //Registration question answers
         $answerString = '';
         if($event->registration_needed) {
           if($registration_form = $event->getRegistrationFormQuestions()) {
             foreach($registration_form as $question) {
               $answer = $question->getAnswerFromUser($attendee->getGUID());
               $attendeeDataXml .= '<Cell><Data ss:Type="String">'.($answer->value).'</Data></Cell>';
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
                     $attendeeDataXml .= '<Cell><Data ss:Type="String">'.'joined'.'</Data></Cell>';
                   } else {
                     $attendeeDataXml .= '<Cell><Data ss:Type="String"></Data></Cell>';
                   }
                 }
               }
             }
           }
         }
         $attendeeDataXml .= '</Row>';

       }
   }




   //Filter out the event description table into headers
   $dataXml = '';
   $data = (string)($event->description);

   $dom = new DOMDocument();
   @$dom->loadHTML($data);
   $dom->preserveWhiteSpace = false;
   $xpath = new DOMXPath($dom);

   $results = $xpath->query('/html/body/table/tbody/tr');
   //  echo htmlentities($results);
   foreach ($results as $result){
     $cells = $result -> getElementsByTagName('td');
     $internalTables = $result -> getElementsByTagName('table');
     if($internalTables->length > 0){
        foreach ($internalTables as $it) {
           $icells = $it -> getElementsByTagName('tr');
           foreach($icells as $val){
             $cells = $val -> getElementsByTagName('td');
             $headerXml .=  '<Cell><Data ss:Type="String">'.($cells->item(0)->nodeValue).'</Data></Cell>';
             $eventXml .='<Cell><Data ss:Type="String">'.($cells->item(1)->nodeValue).'</Data></Cell>';
           }
       }
      }else{
       $headerXml .=  '<Cell><Data ss:Type="String">'.($cells->item(0)->nodeValue).'</Data></Cell>';
       $eventXml .='<Cell><Data ss:Type="String">'.($cells->item(1)->nodeValue).'</Data></Cell>';
     }
   }
  $eventXml .= '</Row>';
  $headerXml .= '</Row>';
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

//return sheet of event info
  return $headerXml.$eventXml.$attendeeHeaderXml.$attendeeDataXml.$endXml;
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
