<?php
/**
 * Created by PhpStorm.
 * User: Irin A
 * Date: 11/21/2017
 * Time: 10:17 AM
 */
function generate_export_spreadsheet($event){
  $event_options = array();
  $event_options["container_guid"] =$groupGuid;
  $events = event_manager_search_events($event_options);
  $eventEntities = $events["entities"];
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
  foreach ($eventEntities as $event) {
    $xml = group_events_export_sheet($event);
    $spreadsheetExportString .= $xml;
  }
  $spreadsheetExportString .='
    </Workbook>
       ';

  return $spreadsheetExportString;
}






function group_events_export_sheet($event){



  /*

$xml = new DOMDocument();
//$xml->validateOnParse = true;
$xml->loadHTML($html);

$xpath = new DOMXPath($xml);
$table =$xpath->query("table")->item(0);
echo var_dump($table);

exit();
*/

  $old_ia = elgg_get_ignore_access();
  elgg_set_ignore_access(true);
  $EOL = "\r\n";


  $headerXml = '
   <Worksheet ss:Name="'.$event->title.'">
   <Table
   x:FullColumns="1"
   x:FullRows="1">
   <Column ss:Index="4" ss:AutoFitWidth="0" ss:Width="154.5"/>
   <Row ss:StyleID="s23">
   <Cell><Data ss:Type="String">Name</Data></Cell>
   <Cell><Data ss:Type="String">Email</Data></Cell>
   <Cell><Data ss:Type="String">Status</Data></Cell>
   </Row>';


/*
  //Title
  $titleString .= '"'.$event->title.'"';
  //Fields
  $headerString .= '"'.elgg_echo('name').'","'.elgg_echo('email').'","'.elgg_echo('Status').'"';
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
  //End Fields
  $headerString .= $EOL;
*/
  $includeEvent = false;
  //generate event data
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
        $includeEvent = true;
        reset($peopleResponded);
        foreach($peopleResponded as $attendee) {

          $dataXml .=  '<Row>
          <Cell><Data ss:Type="String">'.(string)$attendee->name.'</Data></Cell>
          <Cell ss:StyleID="s21" ss:HRef="mailto:molly@katzen.com">
          <Data ss:Type="String">'.(string)$attendee->email.'</Data></Cell>
          <Cell><Data ss:Type="String">'.(string)htmlentities($event->description).'</Data></Cell>
          </Row>';


          $data = (string)($event->description);
          $dom = new DOMDocument();
          @$dom->loadHTML($data);
          $dom->preserveWhiteSpace = false;
          $xpath = new DOMXPath($dom);

          echo $data."<br/>";
          $res = $xpath->query('/table/tbody/tr');
          foreach ($res as $row) {
            $cells = $row -> getElementsByTagName('td');
            echo var_dump($cells->item(0)->nodeValue)." , ".var_dump($cells->item(1)->nodeValue)."<br/>";
          }

          echo "End Parent Table <br/>";






          $tables = $dom->getElementsByTagName('table');
          foreach($tables as $table){
            $rows = $table->getElementsByTagName("tr");
            //var_dump($tables->item(1));
            foreach ($rows as $row) {
                        $cells = $row -> getElementsByTagName('td');
                        echo var_dump($cells->item(0)->nodeValue)." , ".var_dump($cells->item(1)->nodeValue)."<br/>";
              }

              echo "-----<br/>";
          }



          /*

          foreach($firstCol as $rows[0]){
            $cells = $firstCol -> getElementsByTagName('td');
            echo $cells."<br/>";
          }
          foreach ($rows as $row) {
                      $cells = $row -> getElementsByTagName('td');
                      foreach ($cells as $cell) {
                        echo $cell->nodeValue; // print cells' content as 124578
            }
                      echo "<br/>";
          }
          */
          exit();




          /*
          //Registration question answers
          $answerString = '';
          if($event->registration_needed) {
            if($registration_form = $event->getRegistrationFormQuestions()) {
              foreach($registration_form as $question) {
                $answer = $question->getAnswerFromUser($attendee->getGUID());

                $answerString .= '"'.addslashes($answer->value).'",';
              }
            }
            $dataString .= ','.substr($answerString, 0, (strlen($answerString) -1));
          }

          //[V] Checked - Joined a program within event
          if($event->with_program) {
            if($eventDays = $event->getEventDays()) {
              foreach($eventDays as $eventDay) {
                if($eventSlots = $eventDay->getEventSlots()) {
                  foreach($eventSlots as $eventSlot) {
                    if(check_entity_relationship($attendee->getGUID(), EVENT_MANAGER_RELATION_SLOT_REGISTRATION, $eventSlot->getGUID())) {
                      $dataString .= ',"joined"';
                    } else {
                      $dataString .= ',""';
                    }
                  }
                }
              }
            }
          }
          $dataString .= $EOL;
          */
          //end of data, move on to next person
        }
      }
  }

//  $titleString .= $EOL;
  elgg_set_ignore_access($old_ia);

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
       /*

      */
  //return $headerXml.$dataXml.$endXml;
  if($includeEvent){
  return $headerXml.$dataXml.$endXml;
  }
  else{
    return '';
  }
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
