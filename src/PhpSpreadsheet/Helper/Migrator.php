<?php

namespace PhpOffice\PhpSpreadsheet\Helper;

class Migrator
{
    /**
     * Return the ordered mapping from old PHPExcel class names to new PhpSpreadsheet one.
     *
     * @return string[]
     */
    public function getMapping()
    {
        // Order matters here, we should have the deepest namespaces first (the most "unique" strings)
        $classes = array('PHPExcel_Shared_Escher_DggContainer_BstoreContainer_BSE_Blip' => 'Blip', 'PHPExcel_Shared_Escher_DgContainer_SpgrContainer_SpContainer' => 'SpContainer', 'PHPExcel_Shared_Escher_DggContainer_BstoreContainer_BSE' => 'BSE', 'PHPExcel_Shared_Escher_DgContainer_SpgrContainer' => 'SpgrContainer', 'PHPExcel_Shared_Escher_DggContainer_BstoreContainer' => 'BstoreContainer', 'PHPExcel_Shared_OLE_PPS_File' => 'File', 'PHPExcel_Shared_OLE_PPS_Root' => 'Root', 'PHPExcel_Worksheet_AutoFilter_Column_Rule' => 'Rule', 'PHPExcel_Writer_OpenDocument_Cell_Comment' => 'Comment', 'PHPExcel_Calculation_Token_Stack' => 'Stack', 'PHPExcel_Chart_Renderer_jpgraph' => 'JpGraph', 'PHPExcel_Reader_Excel5_Escher' => 'Escher', 'PHPExcel_Reader_Excel5_MD5' => \PhpOffice\PhpSpreadsheet\Reader\Xls\MD5::class, 'PHPExcel_Reader_Excel5_RC4' => \PhpOffice\PhpSpreadsheet\Reader\Xls\RC4::class, 'PHPExcel_Reader_Excel2007_Chart' => 'Chart', 'PHPExcel_Reader_Excel2007_Theme' => 'Theme', 'PHPExcel_Shared_Escher_DgContainer' => 'DgContainer', 'PHPExcel_Shared_Escher_DggContainer' => 'DggContainer', 'CholeskyDecomposition' => 'CholeskyDecomposition', 'EigenvalueDecomposition' => 'EigenvalueDecomposition', 'PHPExcel_Shared_JAMA_LUDecomposition' => 'LUDecomposition', 'PHPExcel_Shared_JAMA_Matrix' => 'Matrix', 'QRDecomposition' => 'QRDecomposition', 'PHPExcel_Shared_JAMA_QRDecomposition' => 'QRDecomposition', 'SingularValueDecomposition' => 'SingularValueDecomposition', 'PHPExcel_Shared_OLE_ChainedBlockStream' => 'ChainedBlockStream', 'PHPExcel_Shared_OLE_PPS' => 'PPS', 'PHPExcel_Best_Fit' => 'BestFit', 'PHPExcel_Exponential_Best_Fit' => 'ExponentialBestFit', 'PHPExcel_Linear_Best_Fit' => 'LinearBestFit', 'PHPExcel_Logarithmic_Best_Fit' => 'LogarithmicBestFit', 'polynomialBestFit' => 'PolynomialBestFit', 'PHPExcel_Polynomial_Best_Fit' => 'PolynomialBestFit', 'powerBestFit' => 'PowerBestFit', 'PHPExcel_Power_Best_Fit' => 'PowerBestFit', 'trendClass' => 'Trend', 'PHPExcel_Worksheet_AutoFilter_Column' => 'Column', 'PHPExcel_Worksheet_Drawing_Shadow' => 'Shadow', 'PHPExcel_Writer_OpenDocument_Content' => 'Content', 'PHPExcel_Writer_OpenDocument_Meta' => 'Meta', 'PHPExcel_Writer_OpenDocument_MetaInf' => 'MetaInf', 'PHPExcel_Writer_OpenDocument_Mimetype' => 'Mimetype', 'PHPExcel_Writer_OpenDocument_Settings' => 'Settings', 'PHPExcel_Writer_OpenDocument_Styles' => 'Styles', 'PHPExcel_Writer_OpenDocument_Thumbnails' => 'Thumbnails', 'PHPExcel_Writer_OpenDocument_WriterPart' => 'WriterPart', 'PHPExcel_Writer_PDF_Core' => 'Pdf', 'PHPExcel_Writer_PDF_DomPDF' => 'Dompdf', 'PHPExcel_Writer_PDF_mPDF' => 'Mpdf', 'PHPExcel_Writer_PDF_tcPDF' => 'Tcpdf', 'PHPExcel_Writer_Excel5_BIFFwriter' => 'BIFFwriter', 'PHPExcel_Writer_Excel5_Escher' => 'Escher', 'PHPExcel_Writer_Excel5_Font' => 'Font', 'PHPExcel_Writer_Excel5_Parser' => 'Parser', 'PHPExcel_Writer_Excel5_Workbook' => 'Workbook', 'PHPExcel_Writer_Excel5_Worksheet' => 'Worksheet', 'PHPExcel_Writer_Excel5_Xf' => 'Xf', 'PHPExcel_Writer_Excel2007_Chart' => 'Chart', 'PHPExcel_Writer_Excel2007_Comments' => 'Comments', 'PHPExcel_Writer_Excel2007_ContentTypes' => 'ContentTypes', 'PHPExcel_Writer_Excel2007_DocProps' => 'DocProps', 'PHPExcel_Writer_Excel2007_Drawing' => 'Drawing', 'PHPExcel_Writer_Excel2007_Rels' => 'Rels', 'PHPExcel_Writer_Excel2007_RelsRibbon' => 'RelsRibbon', 'PHPExcel_Writer_Excel2007_RelsVBA' => 'RelsVBA', 'PHPExcel_Writer_Excel2007_StringTable' => 'StringTable', 'PHPExcel_Writer_Excel2007_Style' => 'Style', 'PHPExcel_Writer_Excel2007_Theme' => 'Theme', 'PHPExcel_Writer_Excel2007_Workbook' => 'Workbook', 'PHPExcel_Writer_Excel2007_Worksheet' => 'Worksheet', 'PHPExcel_Writer_Excel2007_WriterPart' => 'WriterPart', 'PHPExcel_CachedObjectStorage_CacheBase' => 'Cells', 'PHPExcel_CalcEngine_CyclicReferenceStack' => 'CyclicReferenceStack', 'PHPExcel_CalcEngine_Logger' => 'Logger', 'PHPExcel_Calculation_Functions' => 'Functions', 'PHPExcel_Calculation_Function' => 'Category', 'PHPExcel_Calculation_Database' => 'Database', 'PHPExcel_Calculation_DateTime' => 'DateTime', 'PHPExcel_Calculation_Engineering' => 'Engineering', 'PHPExcel_Calculation_Exception' => 'Exception', 'PHPExcel_Calculation_ExceptionHandler' => 'ExceptionHandler', 'PHPExcel_Calculation_Financial' => 'Financial', 'PHPExcel_Calculation_FormulaParser' => 'FormulaParser', 'PHPExcel_Calculation_FormulaToken' => 'FormulaToken', 'PHPExcel_Calculation_Logical' => 'Logical', 'PHPExcel_Calculation_LookupRef' => 'LookupRef', 'PHPExcel_Calculation_MathTrig' => 'MathTrig', 'PHPExcel_Calculation_Statistical' => 'Statistical', 'PHPExcel_Calculation_TextData' => 'TextData', 'PHPExcel_Cell_AdvancedValueBinder' => 'AdvancedValueBinder', 'PHPExcel_Cell_DataType' => 'DataType', 'PHPExcel_Cell_DataValidation' => 'DataValidation', 'PHPExcel_Cell_DefaultValueBinder' => 'DefaultValueBinder', 'PHPExcel_Cell_Hyperlink' => 'Hyperlink', 'PHPExcel_Cell_IValueBinder' => 'IValueBinder', 'PHPExcel_Chart_Axis' => 'Axis', 'PHPExcel_Chart_DataSeries' => 'DataSeries', 'PHPExcel_Chart_DataSeriesValues' => 'DataSeriesValues', 'PHPExcel_Chart_Exception' => 'Exception', 'PHPExcel_Chart_GridLines' => 'GridLines', 'PHPExcel_Chart_Layout' => 'Layout', 'PHPExcel_Chart_Legend' => 'Legend', 'PHPExcel_Chart_PlotArea' => 'PlotArea', 'PHPExcel_Properties' => 'Properties', 'PHPExcel_Chart_Title' => 'Title', 'PHPExcel_DocumentProperties' => 'Properties', 'PHPExcel_DocumentSecurity' => 'Security', 'PHPExcel_Helper_HTML' => 'Html', 'PHPExcel_Reader_Abstract' => 'BaseReader', 'PHPExcel_Reader_CSV' => 'Csv', 'PHPExcel_Reader_DefaultReadFilter' => 'DefaultReadFilter', 'PHPExcel_Reader_Excel2003XML' => 'Xml', 'PHPExcel_Reader_Exception' => 'Exception', 'PHPExcel_Reader_Gnumeric' => 'Gnumeric', 'PHPExcel_Reader_HTML' => 'Html', 'PHPExcel_Reader_IReadFilter' => 'IReadFilter', 'PHPExcel_Reader_IReader' => 'IReader', 'PHPExcel_Reader_OOCalc' => 'Ods', 'PHPExcel_Reader_SYLK' => 'Slk', 'PHPExcel_Reader_Excel5' => 'Xls', 'PHPExcel_Reader_Excel2007' => 'Xlsx', 'PHPExcel_RichText_ITextElement' => 'ITextElement', 'PHPExcel_RichText_Run' => 'Run', 'PHPExcel_RichText_TextElement' => 'TextElement', 'PHPExcel_Shared_CodePage' => 'CodePage', 'PHPExcel_Shared_Date' => 'Date', 'PHPExcel_Shared_Drawing' => 'Drawing', 'PHPExcel_Shared_Escher' => 'Escher', 'PHPExcel_Shared_File' => 'File', 'PHPExcel_Shared_Font' => 'Font', 'PHPExcel_Shared_OLE' => 'OLE', 'PHPExcel_Shared_OLERead' => 'OLERead', 'PHPExcel_Shared_PasswordHasher' => 'PasswordHasher', 'PHPExcel_Shared_String' => 'StringHelper', 'PHPExcel_Shared_TimeZone' => 'TimeZone', 'PHPExcel_Shared_XMLWriter' => 'XMLWriter', 'PHPExcel_Shared_Excel5' => 'Xls', 'PHPExcel_Style_Alignment' => 'Alignment', 'PHPExcel_Style_Border' => 'Border', 'PHPExcel_Style_Borders' => 'Borders', 'PHPExcel_Style_Color' => 'Color', 'PHPExcel_Style_Conditional' => 'Conditional', 'PHPExcel_Style_Fill' => 'Fill', 'PHPExcel_Style_Font' => 'Font', 'PHPExcel_Style_NumberFormat' => 'NumberFormat', 'PHPExcel_Style_Protection' => 'Protection', 'PHPExcel_Style_Supervisor' => 'Supervisor', 'PHPExcel_Worksheet_AutoFilter' => 'AutoFilter', 'PHPExcel_Worksheet_BaseDrawing' => 'BaseDrawing', 'PHPExcel_Worksheet_CellIterator' => 'CellIterator', 'PHPExcel_Worksheet_Column' => 'Column', 'PHPExcel_Worksheet_ColumnCellIterator' => 'ColumnCellIterator', 'PHPExcel_Worksheet_ColumnDimension' => 'ColumnDimension', 'PHPExcel_Worksheet_ColumnIterator' => 'ColumnIterator', 'PHPExcel_Worksheet_Drawing' => 'Drawing', 'PHPExcel_Worksheet_HeaderFooter' => 'HeaderFooter', 'PHPExcel_Worksheet_HeaderFooterDrawing' => 'HeaderFooterDrawing', 'PHPExcel_WorksheetIterator' => 'Iterator', 'PHPExcel_Worksheet_MemoryDrawing' => 'MemoryDrawing', 'PHPExcel_Worksheet_PageMargins' => 'PageMargins', 'PHPExcel_Worksheet_PageSetup' => 'PageSetup', 'PHPExcel_Worksheet_Protection' => 'Protection', 'PHPExcel_Worksheet_Row' => 'Row', 'PHPExcel_Worksheet_RowCellIterator' => 'RowCellIterator', 'PHPExcel_Worksheet_RowDimension' => 'RowDimension', 'PHPExcel_Worksheet_RowIterator' => 'RowIterator', 'PHPExcel_Worksheet_SheetView' => 'SheetView', 'PHPExcel_Writer_Abstract' => 'BaseWriter', 'PHPExcel_Writer_CSV' => 'Csv', 'PHPExcel_Writer_Exception' => 'Exception', 'PHPExcel_Writer_HTML' => 'Html', 'PHPExcel_Writer_IWriter' => 'IWriter', 'PHPExcel_Writer_OpenDocument' => 'Ods', 'PHPExcel_Writer_PDF' => 'Pdf', 'PHPExcel_Writer_Excel5' => 'Xls', 'PHPExcel_Writer_Excel2007' => 'Xlsx', 'PHPExcel_CachedObjectStorageFactory' => 'CellsFactory', 'PHPExcel_Calculation' => 'Calculation', 'PHPExcel_Cell' => 'Cell', 'PHPExcel_Chart' => 'Chart', 'PHPExcel_Comment' => 'Comment', 'PHPExcel_Exception' => 'Exception', 'PHPExcel_HashTable' => 'HashTable', 'PHPExcel_IComparable' => 'IComparable', 'PHPExcel_IOFactory' => 'IOFactory', 'PHPExcel_NamedRange' => 'NamedRange', 'PHPExcel_ReferenceHelper' => 'ReferenceHelper', 'PHPExcel_RichText' => 'RichText', 'PHPExcel_Settings' => 'Settings', 'PHPExcel_Style' => 'Style', 'PHPExcel_Worksheet' => 'Worksheet', 'PHPExcel' => 'Spreadsheet');
        $methods = array('MINUTEOFHOUR' => 'MINUTE', 'SECONDOFMINUTE' => 'SECOND', 'DAYOFWEEK' => 'WEEKDAY', 'WEEKOFYEAR' => 'WEEKNUM', 'ExcelToPHPObject' => 'excelToDateTimeObject', 'ExcelToPHP' => 'excelToTimestamp', 'FormattedPHPToExcel' => 'formattedPHPToExcel', 'Cell::absoluteCoordinate' => 'Coordinate::absoluteCoordinate', 'Cell::absoluteReference' => 'Coordinate::absoluteReference', 'Cell::buildRange' => 'Coordinate::buildRange', 'Cell::columnIndexFromString' => 'Coordinate::columnIndexFromString', 'Cell::coordinateFromString' => 'Coordinate::coordinateFromString', 'Cell::extractAllCellReferencesInRange' => 'Coordinate::extractAllCellReferencesInRange', 'Cell::getRangeBoundaries' => 'Coordinate::getRangeBoundaries', 'Cell::mergeRangesInCollection' => 'Coordinate::mergeRangesInCollection', 'Cell::rangeBoundaries' => 'Coordinate::rangeBoundaries', 'Cell::rangeDimension' => 'Coordinate::rangeDimension', 'Cell::splitRange' => 'Coordinate::splitRange', 'Cell::stringFromColumnIndex' => 'Coordinate::stringFromColumnIndex');
        // Keep '\' prefix for class names
        $prefixedClasses = array();
        foreach ($classes as $key => &$value) {
            $value = str_replace('PhpOffice\\', '\\PhpOffice\\', $value);
            $prefixedClasses['\\' . $key] = $value;
        }
        $mapping = $prefixedClasses + $classes + $methods;
        return $mapping;
    }
    /**
     * Search in all files in given directory.
     *
     * @param string $path
     */
    private function recursiveReplace($path)
    {
        $patterns = array('/*.md', '/*.php', '/*.phtml', '/*.txt', '/*.TXT');
        $from = array_keys($this->getMapping());
        $to = array_values($this->getMapping());
        foreach ($patterns as $pattern) {
            foreach (glob($path . $pattern) as $file) {
                $original = file_get_contents($file);
                $converted = str_replace($from, $to, $original);
                if ($original !== $converted) {
                    echo $file . ' converted
';
                    file_put_contents($file, $converted);
                }
            }
        }
        // Do the recursion in subdirectory
        foreach (glob($path . '/*', GLOB_ONLYDIR) as $subpath) {
            if (strpos($subpath, $path . '/') === 0) {
                $this->recursiveReplace($subpath);
            }
        }
    }
    public function migrate()
    {
        $path = realpath(getcwd());
        echo 'This will search and replace recursively in ' . $path . PHP_EOL;
        echo 'You MUST backup your files first, or you risk losing data.' . PHP_EOL;
        echo 'Are you sure ? (y/n)';
        $confirm = fread(STDIN, 1);
        if ($confirm === 'y') {
            $this->recursiveReplace($path);
        }
    }
}