<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
require __DIR__ . '/../../Header.php';
$helper->log('Returns the serial number of a particular date.');
// Create new PhpSpreadsheet object
$spreadsheet = new Spreadsheet();
$worksheet = $spreadsheet->getActiveSheet();
// Add some data
$testDates = array(array(2012, 3, 26), array(2012, 2, 29), array(2012, 4, 1), array(2012, 12, 25), array(2012, 10, 31), array(2012, 11, 5), array(2012, 1, 1), array(2012, 3, 17), array(2011, 2, 29), array(7, 5, 3), array(2012, 13, 1), array(2012, 11, 45), array(2012, 0, 0), array(2012, 1, 0), array(2012, 0, 1), array(2012, -2, 2), array(2012, 2, -2), array(2012, -2, -2));
$testDateCount = count($testDates);
$worksheet->fromArray($testDates, null, 'A1', true);
for ($row = 1; $row <= $testDateCount; ++$row) {
    $worksheet->setCellValue('D' . $row, '=DATE(A' . $row . ',B' . $row . ',C' . $row . ')');
    $worksheet->setCellValue('E' . $row, '=D' . $row);
}
$worksheet->getStyle('E1:E' . $testDateCount)->getNumberFormat()->setFormatCode('yyyy-mmm-dd');
// Test the formulae
for ($row = 1; $row <= $testDateCount; ++$row) {
    $helper->log('Year: ' . $worksheet->getCell('A' . $row)->getFormattedValue());
    $helper->log('Month: ' . $worksheet->getCell('B' . $row)->getFormattedValue());
    $helper->log('Day: ' . $worksheet->getCell('C' . $row)->getFormattedValue());
    $helper->log('Formula: ' . $worksheet->getCell('D' . $row)->getValue());
    $helper->log('Excel DateStamp: ' . $worksheet->getCell('D' . $row)->getFormattedValue());
    $helper->log('Formatted DateStamp: ' . $worksheet->getCell('E' . $row)->getFormattedValue());
    $helper->log('');
}