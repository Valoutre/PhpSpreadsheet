<?php

use PhpOffice\PhpSpreadsheet\Chart\Chart;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;
use PhpOffice\PhpSpreadsheet\Chart\DataSeriesValues;
use PhpOffice\PhpSpreadsheet\Chart\Legend;
use PhpOffice\PhpSpreadsheet\Chart\PlotArea;
use PhpOffice\PhpSpreadsheet\Chart\Title;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
require __DIR__ . '/../Header.php';
$spreadsheet = new Spreadsheet();
$worksheet = $spreadsheet->getActiveSheet();
$worksheet->fromArray(array(array('', '', 'Budget', 'Forecast', 'Actual'), array('2010', 'Q1', 47, 44, 43), array('', 'Q2', 56, 53, 50), array('', 'Q3', 52, 46, 45), array('', 'Q4', 45, 40, 40), array('2011', 'Q1', 51, 42, 46), array('', 'Q2', 53, 58, 56), array('', 'Q3', 64, 66, 69), array('', 'Q4', 54, 55, 56), array('2012', 'Q1', 49, 52, 58), array('', 'Q2', 68, 73, 86), array('', 'Q3', 72, 78, 0), array('', 'Q4', 50, 60, 0)));
//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesLabels = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$C$1', null, 1), new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$D$1', null, 1), new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$E$1', null, 1));
//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$A$2:$B$13', null, 12));
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$C$2:$C$13', null, 12), new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$D$2:$D$13', null, 12), new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$E$2:$E$13', null, 12));
//	Build the dataseries
$series = new DataSeries(DataSeries::TYPE_BARCHART, DataSeries::GROUPING_CLUSTERED, range(0, count($dataSeriesValues) - 1), $dataSeriesLabels, $xAxisTickValues, $dataSeriesValues);
//	Set additional dataseries parameters
//		Make it a vertical column rather than a horizontal bar graph
$series->setPlotDirection(DataSeries::DIRECTION_COL);
//	Set the series in the plot area
$plotArea = new PlotArea(null, array($series));
//	Set the chart legend
$legend = new Legend(Legend::POSITION_BOTTOM, null, false);
$title = new Title('Test Grouped Column Chart');
$xAxisLabel = new Title('Financial Period');
$yAxisLabel = new Title('Value ($k)');
//	Create the chart
$chart = new Chart('chart1', $title, $legend, $plotArea, true, 0, $xAxisLabel, $yAxisLabel);
//	Set the position where the chart should appear in the worksheet
$chart->setTopLeftPosition('G2');
$chart->setBottomRightPosition('P20');
//	Add the chart to the worksheet
$worksheet->addChart($chart);
// Save Excel 2007 file
$filename = $helper->getFilename(__FILE__);
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->setIncludeCharts(true);
$callStartTime = microtime(true);
$writer->save($filename);
$helper->logWrite($writer, $filename, $callStartTime);