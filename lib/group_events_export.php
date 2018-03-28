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

function group_events_export_sheet($event){
  $old_ia = elgg_get_ignore_access();
  elgg_set_ignore_access(true);

  $rowSpace = '<Row></Row>';
  $beginXml = '
  <Worksheet ss:Name="'.$event->title.'">
   <Table
   x:FullColumns="1"
   x:FullRows="1">
   <Column />';
  $eventHeaderXml = '
   <Row ss:StyleID="s23">
   <Cell ss:StyleID="s29"><Data ss:Type="String">Event</Data></Cell>
   <Cell ss:StyleID="s29"><Data ss:Type="String">Location</Data></Cell>
   <Cell ss:StyleID="s29"><Data ss:Type="String">Venue</Data></Cell>
   <Cell ss:StyleID="s29"><Data ss:Type="String">Start</Data></Cell>
   <Cell ss:StyleID="s29"><Data ss:Type="String">End</Data></Cell>
   ';
   //Default attendee data
   $attendeeHeaderXml = '
   <Row ss:StyleID="s23">
   <Cell ss:StyleID="s29"><Data ss:Type="String">Name</Data></Cell>
   <Cell ss:StyleID="s29"><Data ss:Type="String">Email</Data></Cell>
   <Cell ss:StyleID="s29"><Data ss:Type="String">Status</Data></Cell>
   ';
   //Build registration question headers
   if($event->registration_needed) {
     if($registration_form = $event->getRegistrationFormQuestions()) {
       foreach($registration_form as $question) {
         $attendeeHeaderXml .= '<Cell ss:StyleID="s29"><Data ss:Type="String">'.$question->title.'</Data></Cell>';
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
             /*
             $start_time = $eventSlot->start_time;
             $end_time = $eventSlot->end_time;

             $start_time_hour = date('H', $start_time);
             $start_time_minutes = date('i', $start_time);

             $end_time_hour = date('H', $end_time);
             $end_time_minutes = date('i', $end_time);
            '\' '.$date. ' ('.$start_time_hour.':'.$start_time_minutes.' - '.$end_time_hour.':'.$end_time_minutes.')
             */
             $attendeeHeaderXml .= '<Cell ss:StyleID="s29"><Data ss:Type="String">'.$eventSlot->title.'</Data></Cell>';
           }
         }
       }
     }
   }
   //End Fields
   $attendeeHeaderXml .= '</Row>';
   $attendeeDataXml = '';
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
      $eventHeaderXml .=  '<Cell ss:StyleID="s29"><Data ss:Type="String">'.ucfirst(substr($relationship,6)).'</Data></Cell>';
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
  $descHeaderXml = '';
  $descDataXml = '';
  $data = (string)($event->description);
   //Filter out the event description table into headers
   if($data != ""){
      $descHeaderXml .= '<Row>';
      $descDataXml .= '<Row>';
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
               foreach($icells as $val){
                 $cells = $val -> getElementsByTagName('td');
                 $descHeaderXml .=  '<Cell  ss:StyleID="s29"><Data ss:Type="String">'.($cells->item(0)->nodeValue).'</Data></Cell>';
                 $descDataXml .='<Cell ss:StyleID="s30"><Data ss:Type="String">'.($cells->item(1)->nodeValue).'</Data></Cell>';
               }
           }
          }else{
           $descHeaderXml .=  '<Cell  ss:StyleID="s29"><Data ss:Type="String">'.($cells->item(0)->nodeValue).'</Data></Cell>';
           $descDataXml .='<Cell ss:StyleID="s30"><Data ss:Type="String">'.($cells->item(1)->nodeValue).'</Data></Cell>';
         }
       }
     }else{
       //No tables in description
       $descHeaderXml .=  '<Cell  ss:StyleID="s29"><Data ss:Type="String">'."Description".'</Data></Cell>';
       $descDataXml .='<Cell ss:StyleID="s30"><Data ss:Type="String">'.$data.'</Data></Cell>';
     }
     //end description data
     $descHeaderXml .= '</Row>';
    $descDataXml .= '</Row>';
  }

  $eventDataXml .= '</Row>';
  $eventHeaderXml .= '</Row>';
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
//event table
$eventTable = $eventHeaderXml.$eventDataXml.$rowSpace;
//optional desc table spacing
$descTable = $descHeaderXml.$descDataXml;
echo var_dump($descTable);
exit();
if($descTable != ""){
     $descHeaderXml .= $rowSpace;
}
$attendeeTable = $attendeeHeaderXml.$attendeeDataXml.$rowSpace;

//return sheet of event info
  return $beginXml.$eventTable.$descTable.$attendeeTable.$endXml;
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
