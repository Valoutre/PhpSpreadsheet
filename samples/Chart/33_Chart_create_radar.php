<?php

use PhpOffice\PhpSpreadsheet\Chart\Chart;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
use PhpOffice\PhpSpreadsheet\Chart\Layout;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
require __DIR__ . '/../Header.php';
$spreadsheet = new Spreadsheet();
$worksheet = $spreadsheet->getActiveSheet();
$worksheet->fromArray(array(array('', 2010, 2011, 2012), array('Jan', 47, 45, 71), array('Feb', 56, 73, 86), array('Mar', 52, 61, 69), array('Apr', 40, 52, 60), array('May', 42, 55, 71), array('Jun', 58, 63, 76), array('Jul', 53, 61, 89), array('Aug', 46, 69, 85), array('Sep', 62, 75, 81), array('Oct', 51, 70, 96), array('Nov', 55, 66, 89), array('Dec', 68, 62, 0)));
//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesLabels = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$C$1', null, 1), new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$D$1', null, 1));
//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$A$2:$A$13', null, 12), new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$A$2:$A$13', null, 12));
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$C$2:$C$13', null, 12), new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$D$2:$D$13', null, 12));
//	Build the dataseries
$series = new DataSeries(DataSeries::TYPE_RADARCHART, null, range(0, count($dataSeriesValues) - 1), $dataSeriesLabels, $xAxisTickValues, $dataSeriesValues, null, null, DataSeries::STYLE_MARKER);
//	Set up a layout object for the Pie chart
$layout = new Layout();
//	Set the series in the plot area
$plotArea = new PlotArea($layout, array($series));
//	Set the chart legend
$legend = new Legend(Legend::POSITION_RIGHT, null, false);
$title = new Title('Test Radar Chart');
//	Create the chart
$chart = new Chart('chart1', $title, $legend, $plotArea, true, 0, null, null);
//	Set the position where the chart should appear in the worksheet
$chart->setTopLeftPosition('F2');
$chart->setBottomRightPosition('M15');
//	Add the chart to the worksheet
$worksheet->addChart($chart);
// Save Excel 2007 file
$filename = $helper->getFilename(__FILE__);
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->setIncludeCharts(true);
$callStartTime = microtime(true);
$writer->save($filename);
$helper->logWrite($writer, $filename, $callStartTime);