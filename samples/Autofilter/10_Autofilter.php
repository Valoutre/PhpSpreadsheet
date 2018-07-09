<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
require __DIR__ . '/../Header.php';
// Create new Spreadsheet object
$helper->log('Create new Spreadsheet object');
$spreadsheet = new Spreadsheet();
// Set document properties
$helper->log('Set document properties');
$spreadsheet->getProperties()->setCreator('Maarten Balliauw')->setLastModifiedBy('Maarten Balliauw')->setTitle('PhpSpreadsheet Test Document')->setSubject('PhpSpreadsheet Test Document')->setDescription('Test document for PhpSpreadsheet, generated using PHP classes.')->setKeywords('office PhpSpreadsheet php')->setCategory('Test result file');
// Create the worksheet
$helper->log('Add data');
$spreadsheet->setActiveSheetIndex(0);
$spreadsheet->getActiveSheet()->setCellValue('A1', 'Year')->setCellValue('B1', 'Quarter')->setCellValue('C1', 'Country')->setCellValue('D1', 'Sales');
$dataArray = array(array('2010', 'Q1', 'United States', 790), array('2010', 'Q2', 'United States', 730), array('2010', 'Q3', 'United States', 860), array('2010', 'Q4', 'United States', 850), array('2011', 'Q1', 'United States', 800), array('2011', 'Q2', 'United States', 700), array('2011', 'Q3', 'United States', 900), array('2011', 'Q4', 'United States', 950), array('2010', 'Q1', 'Belgium', 380), array('2010', 'Q2', 'Belgium', 390), array('2010', 'Q3', 'Belgium', 420), array('2010', 'Q4', 'Belgium', 460), array('2011', 'Q1', 'Belgium', 400), array('2011', 'Q2', 'Belgium', 350), array('2011', 'Q3', 'Belgium', 450), array('2011', 'Q4', 'Belgium', 500), array('2010', 'Q1', 'UK', 690), array('2010', 'Q2', 'UK', 610), array('2010', 'Q3', 'UK', 620), array('2010', 'Q4', 'UK', 600), array('2011', 'Q1', 'UK', 720), array('2011', 'Q2', 'UK', 650), array('2011', 'Q3', 'UK', 580), array('2011', 'Q4', 'UK', 510), array('2010', 'Q1', 'France', 510), array('2010', 'Q2', 'France', 490), array('2010', 'Q3', 'France', 460), array('2010', 'Q4', 'France', 590), array('2011', 'Q1', 'France', 620), array('2011', 'Q2', 'France', 650), array('2011', 'Q3', 'France', 415), array('2011', 'Q4', 'France', 570), array('2010', 'Q1', 'Germany', 720), array('2010', 'Q2', 'Germany', 680), array('2010', 'Q3', 'Germany', 640), array('2010', 'Q4', 'Germany', 660), array('2011', 'Q1', 'Germany', 680), array('2011', 'Q2', 'Germany', 620), array('2011', 'Q3', 'Germany', 710), array('2011', 'Q4', 'Germany', 690), array('2010', 'Q1', 'Spain', 510), array('2010', 'Q2', 'Spain', 490), array('2010', 'Q3', 'Spain', 470), array('2010', 'Q4', 'Spain', 420), array('2011', 'Q1', 'Spain', 460), array('2011', 'Q2', 'Spain', 390), array('2011', 'Q3', 'Spain', 430), array('2011', 'Q4', 'Spain', 415), array('2010', 'Q1', 'Italy', 440), array('2010', 'Q2', 'Italy', 410), array('2010', 'Q3', 'Italy', 420), array('2010', 'Q4', 'Italy', 450), array('2011', 'Q1', 'Italy', 430), array('2011', 'Q2', 'Italy', 370), array('2011', 'Q3', 'Italy', 350), array('2011', 'Q4', 'Italy', 335));
$spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A2');
// Set title row bold
$helper->log('Set title row bold');
$spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setBold(true);
// Set autofilter
$helper->log('Set autofilter');
// Always include the complete filter range!
// Excel does support setting only the caption
// row, but that's not a best practise...
$spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());
// Save
$helper->write($spreadsheet, __FILE__);