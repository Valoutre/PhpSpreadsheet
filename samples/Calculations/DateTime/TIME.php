<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
require __DIR__ . '/../../Header.php';
$helper->log('Returns the serial number of a particular time.');
// Create new PhpSpreadsheet object
$spreadsheet = new Spreadsheet();
$worksheet = $spreadsheet->getActiveSheet();
// Add some data
$testDates = array(array(3, 15), array(13, 15), array(15, 15, 15), array(3, 15, 30), array(15, 15, 15), array(5), array(9, 15, 0), array(9, 15, -1), array(13, -14, -15), array(0, 0, -1));
$testDateCount = count($testDates);
$worksheet->fromArray($testDates, null, 'A1', true);
for ($row = 1; $row <= $testDateCount; ++$row) {
    $worksheet->setCellValue('D' . $row, '=TIME(A' . $row . ',B' . $row . ',C' . $row . ')');
    $worksheet->setCellValue('E' . $row, '=D' . $row);
}
$worksheet->getStyle('E1:E' . $testDateCount)->getNumberFormat()->setFormatCode('hh:mm:ss');
// Test the formulae
for ($row = 1; $row <= $testDateCount; ++$row) {
    $helper->log('Hour: ' . $worksheet->getCell('A' . $row)->getFormattedValue());
    $helper->log('Minute: ' . $worksheet->getCell('B' . $row)->getFormattedValue());
    $helper->log('Second: ' . $worksheet->getCell('C' . $row)->getFormattedValue());
    $helper->log('Formula: ' . $worksheet->getCell('D' . $row)->getValue());
    $helper->log('Excel TimeStamp: ' . $worksheet->getCell('D' . $row)->getFormattedValue());
    $helper->log('Formatted TimeStamp: ' . $worksheet->getCell('E' . $row)->getFormattedValue());
    $helper->log('');
}