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
$worksheet->fromArray(array(array('', 'Rainfall (mm)', 'Temperature (°F)', 'Humidity (%)'), array('Jan', 78, 52, 61), array('Feb', 64, 54, 62), array('Mar', 62, 57, 63), array('Apr', 21, 62, 59), array('May', 11, 75, 60), array('Jun', 1, 75, 57), array('Jul', 1, 79, 56), array('Aug', 1, 79, 59), array('Sep', 10, 75, 60), array('Oct', 40, 68, 63), array('Nov', 69, 62, 64), array('Dec', 89, 57, 66)));
//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesLabels1 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$B$1', null, 1));
$dataSeriesLabels2 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$C$1', null, 1));
$dataSeriesLabels3 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$D$1', null, 1));
//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$xAxisTickValues = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_STRING, 'Worksheet!$A$2:$A$13', null, 12));
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues1 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$B$2:$B$13', null, 12));
//	Build the dataseries
$series1 = new DataSeries(DataSeries::TYPE_BARCHART, DataSeries::GROUPING_CLUSTERED, range(0, count($dataSeriesValues1) - 1), $dataSeriesLabels1, $xAxisTickValues, $dataSeriesValues1);
//	Set additional dataseries parameters
//		Make it a vertical column rather than a horizontal bar graph
$series1->setPlotDirection(DataSeries::DIRECTION_COL);
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues2 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$C$2:$C$13', null, 12));
//	Build the dataseries
$series2 = new DataSeries(DataSeries::TYPE_LINECHART, DataSeries::GROUPING_STANDARD, range(0, count($dataSeriesValues2) - 1), $dataSeriesLabels2, array(), $dataSeriesValues2);
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
$dataSeriesValues3 = array(new DataSeriesValues(DataSeriesValues::DATASERIES_TYPE_NUMBER, 'Worksheet!$D$2:$D$13', null, 12));
//	Build the dataseries
$series3 = new DataSeries(DataSeries::TYPE_AREACHART, DataSeries::GROUPING_STANDARD, range(0, count($dataSeriesValues2) - 1), $dataSeriesLabels3, array(), $dataSeriesValues3);
//	Set the series in the plot area
$plotArea = new PlotArea(null, array($series1, $series2, $series3));
//	Set the chart legend
$legend = new Legend(Legend::POSITION_RIGHT, null, false);
$title = new Title('Average Weather Chart for Crete');
//	Create the chart
$chart = new Chart('chart1', $title, $legend, $plotArea, true, 0, null, null);
//	Set the position where the chart should appear in the worksheet
$chart->setTopLeftPosition('F2');
$chart->setBottomRightPosition('O16');
//	Add the chart to the worksheet
$worksheet->addChart($chart);
// Save Excel 2007 file
$filename = $helper->getFilename(__FILE__);
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->setIncludeCharts(true);
$callStartTime = microtime(true);
$writer->save($filename);
$helper->logWrite($writer, $filename, $callStartTime);