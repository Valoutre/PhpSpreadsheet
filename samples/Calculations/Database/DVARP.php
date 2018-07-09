<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
require __DIR__ . '/../../Header.php';
$helper->log('Calculates variance based on the entire population of selected database entries,');
// Create new PhpSpreadsheet object
$spreadsheet = new Spreadsheet();
$worksheet = $spreadsheet->getActiveSheet();
// Add some data
$database = array(array('Tree', 'Height', 'Age', 'Yield', 'Profit'), array('Apple', 18, 20, 14, 105.0), array('Pear', 12, 12, 10, 96.0), array('Cherry', 13, 14, 9, 105.0), array('Apple', 14, 15, 10, 75.0), array('Pear', 9, 8, 8, 76.8), array('Apple', 8, 9, 6, 45.0));
$criteria = array(array('Tree', 'Height', 'Age', 'Yield', 'Profit', 'Height'), array('="=Apple"', '>10', null, null, null, '<16'), array('="=Pear"', null, null, null, null, null));
$worksheet->fromArray($criteria, null, 'A1');
$worksheet->fromArray($database, null, 'A4');
$worksheet->setCellValue('A12', 'The variance in the yield of Apple and Pear trees');
$worksheet->setCellValue('B12', '=DVARP(A4:E10,"Yield",A1:A3)');
$worksheet->setCellValue('A13', 'The variance in height of Apple and Pear trees');
$worksheet->setCellValue('B13', '=DVARP(A4:E10,2,A1:A3)');
$helper->log('Database');
$databaseData = $worksheet->rangeToArray('A4:E10', null, true, true, true);
var_dump($databaseData);
// Test the formulae
$helper->log('Criteria');
$criteriaData = $worksheet->rangeToArray('A1:A3', null, true, true, true);
var_dump($criteriaData);
$helper->log($worksheet->getCell('A12')->getValue());
$helper->log('DVARP() Result is ' . $worksheet->getCell('B12')->getCalculatedValue());
$helper->log('Criteria');
$criteriaData = $worksheet->rangeToArray('A1:A3', null, true, true, true);
var_dump($criteriaData);
$helper->log($worksheet->getCell('A13')->getValue());
$helper->log('DVARP() Result is ' . $worksheet->getCell('B13')->getCalculatedValue());