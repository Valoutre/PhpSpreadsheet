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
$worksheet->fromArray(array(array('', 2010, 2011, 2012), array('Q1', 12, 15, 21), array('Q2', 56, 73, 86), array('Q3', 52, 61, 69), array('Q4', 30, 32, 0)));
//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesLabels1 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$C$1', null, 1));
//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues1 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$A$2:$A$5', null, 4));
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues1 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$C$2:$C$5', null, 4));
//	Build the dataseries
$series1 = new DataSeries(DataSeries::TYPE_PIECHART, null, range(0, count($dataSeriesValues1) - 1), $dataSeriesLabels1, $xAxisTickValues1, $dataSeriesValues1);
//	Set up a layout object for the Pie chart
$layout1 = new Layout();
$layout1->setShowVal(true);
$layout1->setShowPercent(true);
//	Set the series in the plot area
$plotArea1 = new PlotArea($layout1, array($series1));
//	Set the chart legend
$legend1 = new Legend(Legend::POSITION_RIGHT, null, false);
$title1 = new Title('Test Pie Chart');
//	Create the chart
$chart1 = new Chart('chart1', $title1, $legend1, $plotArea1, true, 0, null, null);
//	Set the position where the chart should appear in the worksheet
$chart1->setTopLeftPosition('A7');
$chart1->setBottomRightPosition('H20');
//	Add the chart to the worksheet
$worksheet->addChart($chart1);
//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesLabels2 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$C$1', null, 1));
//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues2 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$A$2:$A$5', null, 4));
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues2 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$C$2:$C$5', null, 4));
//	Build the dataseries
$series2 = new DataSeries(DataSeries::TYPE_DONUTCHART, null, range(0, count($dataSeriesValues2) - 1), $dataSeriesLabels2, $xAxisTickValues2, $dataSeriesValues2);
//	Set up a layout object for the Pie chart
$layout2 = new Layout();
$layout2->setShowVal(true);
$layout2->setShowCatName(true);
//	Set the series in the plot area
$plotArea2 = new PlotArea($layout2, array($series2));
$title2 = new Title('Test Donut Chart');
//	Create the chart
$chart2 = new Chart('chart2', $title2, null, $plotArea2, true, 0, null, null);
//	Set the position where the chart should appear in the worksheet
$chart2->setTopLeftPosition('I7');
$chart2->setBottomRightPosition('P20');
//	Add the chart to the worksheet
$worksheet->addChart($chart2);
// Save Excel 2007 file
$filename = $helper->getFilename(__FILE__);
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->setIncludeCharts(true);
$callStartTime = microtime(true);
$writer->save($filename);
$helper->logWrite($writer, $filename, $callStartTime);