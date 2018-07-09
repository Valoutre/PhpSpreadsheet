<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Xls;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\RichText\Run;
use PhpOffice\PhpSpreadsheet\Shared\StringHelper;
use PhpOffice\PhpSpreadsheet\Shared\Xls;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Conditional;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Worksheet\SheetView;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;
// Original file header of PEAR::Spreadsheet_Excel_Writer_Worksheet (used as the base for this class):
// -----------------------------------------------------------------------------------------
// /*
// *  Module written/ported by Xavier Noguer <xnoguer@rezebra.com>
// *
// *  The majority of this is _NOT_ my code.  I simply ported it from the
// *  PERL Spreadsheet::WriteExcel module.
// *
// *  The author of the Spreadsheet::WriteExcel module is John McNamara
// *  <jmcnamara@cpan.org>
// *
// *  I _DO_ maintain this code, and John McNamara has nothing to do with the
// *  porting of this code to PHP.  Any questions directly related to this
// *  class library should be directed to me.
// *
// *  License Information:
// *
// *    Spreadsheet_Excel_Writer:  A library for generating Excel Spreadsheets
// *    Copyright (c) 2002-2003 Xavier Noguer xnoguer@rezebra.com
// *
// *    This library is free software; you can redistribute it and/or
// *    modify it under the terms of the GNU Lesser General Public
// *    License as published by the Free Software Foundation; either
// *    version 2.1 of the License, or (at your option) any later version.
// *
// *    This library is distributed in the hope that it will be useful,
// *    but WITHOUT ANY WARRANTY; without even the implied warranty of
// *    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
// *    Lesser General Public License for more details.
// *
// *    You should have received a copy of the GNU Lesser General Public
// *    License along with this library; if not, write to the Free Software
// *    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
// */
class Worksheet extends BIFFwriter
{
    /**
     * Formula parser.
     *
     * @var \PhpOffice\PhpSpreadsheet\Writer\Xls\Parser
     */
    private $parser;
    /**
     * Maximum number of characters for a string (LABEL record in BIFF5).
     *
     * @var int
     */
    private $xlsStringMaxLength;
    /**
     * Array containing format information for columns.
     *
     * @var array
     */
    private $columnInfo;
    /**
     * Array containing the selected area for the worksheet.
     *
     * @var array
     */
    private $selection;
    /**
     * The active pane for the worksheet.
     *
     * @var int
     */
    private $activePane;
    /**
     * Whether to use outline.
     *
     * @var int
     */
    private $outlineOn;
    /**
     * Auto outline styles.
     *
     * @var bool
     */
    private $outlineStyle;
    /**
     * Whether to have outline summary below.
     *
     * @var bool
     */
    private $outlineBelow;
    /**
     * Whether to have outline summary at the right.
     *
     * @var bool
     */
    private $outlineRight;
    /**
     * Reference to the total number of strings in the workbook.
     *
     * @var int
     */
    private $stringTotal;
    /**
     * Reference to the number of unique strings in the workbook.
     *
     * @var int
     */
    private $stringUnique;
    /**
     * Reference to the array containing all the unique strings in the workbook.
     *
     * @var array
     */
    private $stringTable;
    /**
     * Color cache.
     */
    private $colors;
    /**
     * Index of first used row (at least 0).
     *
     * @var int
     */
    private $firstRowIndex;
    /**
     * Index of last used row. (no used rows means -1).
     *
     * @var int
     */
    private $lastRowIndex;
    /**
     * Index of first used column (at least 0).
     *
     * @var int
     */
    private $firstColumnIndex;
    /**
     * Index of last used column (no used columns means -1).
     *
     * @var int
     */
    private $lastColumnIndex;
    /**
     * Sheet object.
     *
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public $phpSheet;
    /**
     * Count cell style Xfs.
     *
     * @var int
     */
    private $countCellStyleXfs;
    /**
     * Escher object corresponding to MSODRAWING.
     *
     * @var \PhpOffice\PhpSpreadsheet\Shared\Escher
     */
    private $escher;
    /**
     * Array of font hashes associated to FONT records index.
     *
     * @var array
     */
    public $fontHashIndex;
    /**
     * @var bool
     */
    private $preCalculateFormulas;
    /**
     * @var int
     */
    private $printHeaders;
    /**
     * Constructor.
     *
     * @param int $str_total Total number of strings
     * @param int $str_unique Total number of unique strings
     * @param array &$str_table String Table
     * @param array &$colors Colour Table
     * @param Parser $parser The formula parser created for the Workbook
     * @param bool $preCalculateFormulas Flag indicating whether formulas should be calculated or just written
     * @param \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpSheet The worksheet to write
     */
    public function __construct(&$str_total, &$str_unique, &$str_table, &$colors, Parser $parser, $preCalculateFormulas, \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $phpSheet)
    {
        // It needs to call its parent's constructor explicitly
        parent::__construct();
        $this->preCalculateFormulas = $preCalculateFormulas;
        $this->stringTotal =& $str_total;
        $this->stringUnique =& $str_unique;
        $this->stringTable =& $str_table;
        $this->colors =& $colors;
        $this->parser = $parser;
        $this->phpSheet = $phpSheet;
        $this->xlsStringMaxLength = 255;
        $this->columnInfo = array();
        $this->selection = array(0, 0, 0, 0);
        $this->activePane = 3;
        $this->printHeaders = 0;
        $this->outlineStyle = 0;
        $this->outlineBelow = 1;
        $this->outlineRight = 1;
        $this->outlineOn = 1;
        $this->fontHashIndex = array();
        // calculate values for DIMENSIONS record
        $minR = 1;
        $minC = 'A';
        $maxR = $this->phpSheet->getHighestRow();
        $maxC = $this->phpSheet->getHighestColumn();
        // Determine lowest and highest column and row
        $this->lastRowIndex = $maxR > 65535 ? 65535 : $maxR;
        $this->firstColumnIndex = Coordinate::columnIndexFromString($minC);
        $this->lastColumnIndex = Coordinate::columnIndexFromString($maxC);
        //        if ($this->firstColumnIndex > 255) $this->firstColumnIndex = 255;
        if ($this->lastColumnIndex > 255) {
            $this->lastColumnIndex = 255;
        }
        $this->countCellStyleXfs = count($phpSheet->getParent()->getCellStyleXfCollection());
    }
    /**
     * Add data to the beginning of the workbook (note the reverse order)
     * and to the end of the workbook.
     *
     * @see \PhpOffice\PhpSpreadsheet\Writer\Xls\Workbook::storeWorkbook()
     */
    public function close()
    {
        $phpSheet = $this->phpSheet;
        // Write BOF record
        $this->storeBof(16);
        // Write PRINTHEADERS
        $this->writePrintHeaders();
        // Write PRINTGRIDLINES
        $this->writePrintGridlines();
        // Write GRIDSET
        $this->writeGridset();
        // Calculate column widths
        $phpSheet->calculateColumnWidths();
        // Column dimensions
        if (($defaultWidth = $phpSheet->getDefaultColumnDimension()->getWidth()) < 0) {
            $defaultWidth = \PhpOffice\PhpSpreadsheet\Shared\Font::getDefaultColumnWidthByFont($phpSheet->getParent()->getDefaultStyle()->getFont());
        }
        $columnDimensions = $phpSheet->getColumnDimensions();
        $maxCol = $this->lastColumnIndex - 1;
        for ($i = 0; $i <= $maxCol; ++$i) {
            $hidden = 0;
            $level = 0;
            $xfIndex = 15;
            // there are 15 cell style Xfs
            $width = $defaultWidth;
            $columnLetter = Coordinate::stringFromColumnIndex($i + 1);
            if (isset($columnDimensions[$columnLetter])) {
                $columnDimension = $columnDimensions[$columnLetter];
                if ($columnDimension->getWidth() >= 0) {
                    $width = $columnDimension->getWidth();
                }
                $hidden = $columnDimension->getVisible() ? 0 : 1;
                $level = $columnDimension->getOutlineLevel();
                $xfIndex = $columnDimension->getXfIndex() + 15;
            }
            // Components of columnInfo:
            // $firstcol first column on the range
            // $lastcol  last column on the range
            // $width    width to set
            // $xfIndex  The optional cell style Xf index to apply to the columns
            // $hidden   The optional hidden atribute
            // $level    The optional outline level
            $this->columnInfo[] = array($i, $i, $width, $xfIndex, $hidden, $level);
        }
        // Write GUTS
        $this->writeGuts();
        // Write DEFAULTROWHEIGHT
        $this->writeDefaultRowHeight();
        // Write WSBOOL
        $this->writeWsbool();
        // Write horizontal and vertical page breaks
        $this->writeBreaks();
        // Write page header
        $this->writeHeader();
        // Write page footer
        $this->writeFooter();
        // Write page horizontal centering
        $this->writeHcenter();
        // Write page vertical centering
        $this->writeVcenter();
        // Write left margin
        $this->writeMarginLeft();
        // Write right margin
        $this->writeMarginRight();
        // Write top margin
        $this->writeMarginTop();
        // Write bottom margin
        $this->writeMarginBottom();
        // Write page setup
        $this->writeSetup();
        // Write sheet protection
        $this->writeProtect();
        // Write SCENPROTECT
        $this->writeScenProtect();
        // Write OBJECTPROTECT
        $this->writeObjectProtect();
        // Write sheet password
        $this->writePassword();
        // Write DEFCOLWIDTH record
        $this->writeDefcol();
        // Write the COLINFO records if they exist
        if (!empty($this->columnInfo)) {
            $colcount = count($this->columnInfo);
            for ($i = 0; $i < $colcount; ++$i) {
                $this->writeColinfo($this->columnInfo[$i]);
            }
        }
        $autoFilterRange = $phpSheet->getAutoFilter()->getRange();
        if (!empty($autoFilterRange)) {
            // Write AUTOFILTERINFO
            $this->writeAutoFilterInfo();
        }
        // Write sheet dimensions
        $this->writeDimensions();
        // Row dimensions
        foreach ($phpSheet->getRowDimensions() as $rowDimension) {
            $xfIndex = $rowDimension->getXfIndex() + 15;
            // there are 15 cellXfs
            $this->writeRow($rowDimension->getRowIndex() - 1, $rowDimension->getRowHeight(), $xfIndex, $rowDimension->getVisible() ? '0' : '1', $rowDimension->getOutlineLevel());
        }
        // Write Cells
        foreach ($phpSheet->getCoordinates() as $coordinate) {
            $cell = $phpSheet->getCell($coordinate);
            $row = $cell->getRow() - 1;
            $column = Coordinate::columnIndexFromString($cell->getColumn()) - 1;
            // Don't break Excel break the code!
            if ($row > 65535 || $column > 255) {
                throw new WriterException('Rows or columns overflow! Excel5 has limit to 65535 rows and 255 columns. Use XLSX instead.');
            }
            // Write cell value
            $xfIndex = $cell->getXfIndex() + 15;
            // there are 15 cell style Xfs
            $cVal = $cell->getValue();
            if ($cVal instanceof RichText) {
                $arrcRun = array();
                $str_len = StringHelper::countCharacters($cVal->getPlainText(), 'UTF-8');
                $str_pos = 0;
                $elements = $cVal->getRichTextElements();
                foreach ($elements as $element) {
                    // FONT Index
                    if ($element instanceof Run) {
                        $str_fontidx = $this->fontHashIndex[$element->getFont()->getHashCode()];
                    } else {
                        $str_fontidx = 0;
                    }
                    $arrcRun[] = array('strlen' => $str_pos, 'fontidx' => $str_fontidx);
                    // Position FROM
                    $str_pos += StringHelper::countCharacters($element->getText(), 'UTF-8');
                }
                $this->writeRichTextString($row, $column, $cVal->getPlainText(), $xfIndex, $arrcRun);
            } else {
                switch ($cell->getDatatype()) {
                    case DataType::TYPE_STRING:
                    case DataType::TYPE_NULL:
                        if ($cVal === '' || $cVal === null) {
                            $this->writeBlank($row, $column, $xfIndex);
                        } else {
                            $this->writeString($row, $column, $cVal, $xfIndex);
                        }
                        break;
                    case DataType::TYPE_NUMERIC:
                        $this->writeNumber($row, $column, $cVal, $xfIndex);
                        break;
                    case DataType::TYPE_FORMULA:
                        $calculatedValue = $this->preCalculateFormulas ? $cell->getCalculatedValue() : null;
                        $this->writeFormula($row, $column, $cVal, $xfIndex, $calculatedValue);
                        break;
                    case DataType::TYPE_BOOL:
                        $this->writeBoolErr($row, $column, $cVal, 0, $xfIndex);
                        break;
                    case DataType::TYPE_ERROR:
                        $this->writeBoolErr($row, $column, self::mapErrorCode($cVal), 1, $xfIndex);
                        break;
                }
            }
        }
        // Append
        $this->writeMsoDrawing();
        // Write WINDOW2 record
        $this->writeWindow2();
        // Write PLV record
        $this->writePageLayoutView();
        // Write ZOOM record
        $this->writeZoom();
        if ($phpSheet->getFreezePane()) {
            $this->writePanes();
        }
        // Write SELECTION record
        $this->writeSelection();
        // Write MergedCellsTable Record
        $this->writeMergedCells();
        // Hyperlinks
        foreach ($phpSheet->getHyperLinkCollection() as $coordinate => $hyperlink) {
            list($column, $row) = Coordinate::coordinateFromString($coordinate);
            $url = $hyperlink->getUrl();
            if (strpos($url, 'sheet://') !== false) {
                // internal to current workbook
                $url = str_replace('sheet://', 'internal:', $url);
            } elseif (preg_match('/^(http:|https:|ftp:|mailto:)/', $url)) {
            } else {
                // external (local file)
                $url = 'external:' . $url;
            }
            $this->writeUrl($row - 1, Coordinate::columnIndexFromString($column) - 1, $url);
        }
        $this->writeDataValidity();
        $this->writeSheetLayout();
        // Write SHEETPROTECTION record
        $this->writeSheetProtection();
        $this->writeRangeProtection();
        $arrConditionalStyles = $phpSheet->getConditionalStylesCollection();
        if (!empty($arrConditionalStyles)) {
            $arrConditional = array();
            // @todo CFRule & CFHeader
            // Write CFHEADER record
            $this->writeCFHeader();
            // Write ConditionalFormattingTable records
            foreach ($arrConditionalStyles as $cellCoordinate => $conditionalStyles) {
                foreach ($conditionalStyles as $conditional) {
                    if ($conditional->getConditionType() == Conditional::CONDITION_EXPRESSION || $conditional->getConditionType() == Conditional::CONDITION_CELLIS) {
                        if (!isset($arrConditional[$conditional->getHashCode()])) {
                            // This hash code has been handled
                            $arrConditional[$conditional->getHashCode()] = true;
                            // Write CFRULE record
                            $this->writeCFRule($conditional);
                        }
                    }
                }
            }
        }
        $this->storeEof();
    }
    /**
     * Write a cell range address in BIFF8
     * always fixed range
     * See section 2.5.14 in OpenOffice.org's Documentation of the Microsoft Excel File Format.
     *
     * @param string $range E.g. 'A1' or 'A1:B6'
     *
     * @return string Binary data
     */
    private function writeBIFF8CellRangeAddressFixed($range)
    {
        $explodes = explode(':', $range);
        // extract first cell, e.g. 'A1'
        $firstCell = $explodes[0];
        // extract last cell, e.g. 'B6'
        if (count($explodes) == 1) {
            $lastCell = $firstCell;
        } else {
            $lastCell = $explodes[1];
        }
        $firstCellCoordinates = Coordinate::coordinateFromString($firstCell);
        // e.g. [0, 1]
        $lastCellCoordinates = Coordinate::coordinateFromString($lastCell);
        // e.g. [1, 6]
        return pack('vvvv', $firstCellCoordinates[1] - 1, $lastCellCoordinates[1] - 1, Coordinate::columnIndexFromString($firstCellCoordinates[0]) - 1, Coordinate::columnIndexFromString($lastCellCoordinates[0]) - 1);
    }
    /**
     * Retrieves data from memory in one chunk, or from disk in $buffer
     * sized chunks.
     *
     * @return string The data
     */
    public function getData()
    {
        $buffer = 4096;
        // Return data stored in memory
        if (isset($this->_data)) {
            $tmp = $this->_data;
            unset($this->_data);
            return $tmp;
        }
        // No data to return
        return false;
    }
    /**
     * Set the option to print the row and column headers on the printed page.
     *
     * @param int $print Whether to print the headers or not. Defaults to 1 (print).
     */
    public function printRowColHeaders($print = 1)
    {
        $this->printHeaders = $print;
    }
    /**
     * This method sets the properties for outlining and grouping. The defaults
     * correspond to Excel's defaults.
     *
     * @param bool $visible
     * @param bool $symbols_below
     * @param bool $symbols_right
     * @param bool $auto_style
     */
    public function setOutline($visible = true, $symbols_below = true, $symbols_right = true, $auto_style = false)
    {
        $this->outlineOn = $visible;
        $this->outlineBelow = $symbols_below;
        $this->outlineRight = $symbols_right;
        $this->outlineStyle = $auto_style;
        // Ensure this is a boolean vale for Window2
        if ($this->outlineOn) {
            $this->outlineOn = 1;
        }
    }
    /**
     * Write a double to the specified row and column (zero indexed).
     * An integer can be written as a double. Excel will display an
     * integer. $format is optional.
     *
     * Returns  0 : normal termination
     *         -2 : row or column out of range
     *
     * @param int $row Zero indexed row
     * @param int $col Zero indexed column
     * @param float $num The number to write
     * @param mixed $xfIndex The optional XF format
     *
     * @return int
     */
    private function writeNumber($row, $col, $num, $xfIndex)
    {
        $record = 515;
        // Record identifier
        $length = 14;
        // Number of bytes to follow
        $header = pack('vv', $record, $length);
        $data = pack('vvv', $row, $col, $xfIndex);
        $xl_double = pack('d', $num);
        if (self::getByteOrder()) {
            // if it's Big Endian
            $xl_double = strrev($xl_double);
        }
        $this->append($header . $data . $xl_double);
        return 0;
    }
    /**
     * Write a LABELSST record or a LABEL record. Which one depends on BIFF version.
     *
     * @param int $row Row index (0-based)
     * @param int $col Column index (0-based)
     * @param string $str The string
     * @param int $xfIndex Index to XF record
     */
    private function writeString($row, $col, $str, $xfIndex)
    {
        $this->writeLabelSst($row, $col, $str, $xfIndex);
    }
    /**
     * Write a LABELSST record or a LABEL record. Which one depends on BIFF version
     * It differs from writeString by the writing of rich text strings.
     *
     * @param int $row Row index (0-based)
     * @param int $col Column index (0-based)
     * @param string $str The string
     * @param int $xfIndex The XF format index for the cell
     * @param array $arrcRun Index to Font record and characters beginning
     */
    private function writeRichTextString($row, $col, $str, $xfIndex, $arrcRun)
    {
        $record = 253;
        // Record identifier
        $length = 10;
        // Bytes to follow
        $str = StringHelper::UTF8toBIFF8UnicodeShort($str, $arrcRun);
        // check if string is already present
        if (!isset($this->stringTable[$str])) {
            $this->stringTable[$str] = $this->stringUnique++;
        }
        ++$this->stringTotal;
        $header = pack('vv', $record, $length);
        $data = pack('vvvV', $row, $col, $xfIndex, $this->stringTable[$str]);
        $this->append($header . $data);
    }
    /**
     * Write a string to the specified row and column (zero indexed).
     * This is the BIFF8 version (no 255 chars limit).
     * $format is optional.
     *
     * @param int $row Zero indexed row
     * @param int $col Zero indexed column
     * @param string $str The string to write
     * @param mixed $xfIndex The XF format index for the cell
     */
    private function writeLabelSst($row, $col, $str, $xfIndex)
    {
        $record = 253;
        // Record identifier
        $length = 10;
        // Bytes to follow
        $str = StringHelper::UTF8toBIFF8UnicodeLong($str);
        // check if string is already present
        if (!isset($this->stringTable[$str])) {
            $this->stringTable[$str] = $this->stringUnique++;
        }
        ++$this->stringTotal;
        $header = pack('vv', $record, $length);
        $data = pack('vvvV', $row, $col, $xfIndex, $this->stringTable[$str]);
        $this->append($header . $data);
    }
    /**
     * Write a blank cell to the specified row and column (zero indexed).
     * A blank cell is used to specify formatting without adding a string
     * or a number.
     *
     * A blank cell without a format serves no purpose. Therefore, we don't write
     * a BLANK record unless a format is specified.
     *
     * Returns  0 : normal termination (including no format)
     *         -1 : insufficient number of arguments
     *         -2 : row or column out of range
     *
     * @param int $row Zero indexed row
     * @param int $col Zero indexed column
     * @param mixed $xfIndex The XF format index
     *
     * @return int
     */
    public function writeBlank($row, $col, $xfIndex)
    {
        $record = 513;
        // Record identifier
        $length = 6;
        // Number of bytes to follow
        $header = pack('vv', $record, $length);
        $data = pack('vvv', $row, $col, $xfIndex);
        $this->append($header . $data);
        return 0;
    }
    /**
     * Write a boolean or an error type to the specified row and column (zero indexed).
     *
     * @param int $row Row index (0-based)
     * @param int $col Column index (0-based)
     * @param int $value
     * @param bool $isError Error or Boolean?
     * @param int $xfIndex
     *
     * @return int
     */
    private function writeBoolErr($row, $col, $value, $isError, $xfIndex)
    {
        $record = 517;
        $length = 8;
        $header = pack('vv', $record, $length);
        $data = pack('vvvCC', $row, $col, $xfIndex, $value, $isError);
        $this->append($header . $data);
        return 0;
    }
    /**
     * Write a formula to the specified row and column (zero indexed).
     * The textual representation of the formula is passed to the parser in
     * Parser.php which returns a packed binary string.
     *
     * Returns  0 : normal termination
     *         -1 : formula errors (bad formula)
     *         -2 : row or column out of range
     *
     * @param int $row Zero indexed row
     * @param int $col Zero indexed column
     * @param string $formula The formula text string
     * @param mixed $xfIndex The XF format index
     * @param mixed $calculatedValue Calculated value
     *
     * @return int
     */
    private function writeFormula($row, $col, $formula, $xfIndex, $calculatedValue)
    {
        $record = 6;
        // Record identifier
        // Initialize possible additional value for STRING record that should be written after the FORMULA record?
        $stringValue = null;
        // calculated value
        if (isset($calculatedValue)) {
            // Since we can't yet get the data type of the calculated value,
            // we use best effort to determine data type
            if (is_bool($calculatedValue)) {
                // Boolean value
                $num = pack('CCCvCv', 1, 0, (int) $calculatedValue, 0, 0, 65535);
            } elseif (is_int($calculatedValue) || is_float($calculatedValue)) {
                // Numeric value
                $num = pack('d', $calculatedValue);
            } elseif (is_string($calculatedValue)) {
                $errorCodes = DataType::getErrorCodes();
                if (isset($errorCodes[$calculatedValue])) {
                    // Error value
                    $num = pack('CCCvCv', 2, 0, self::mapErrorCode($calculatedValue), 0, 0, 65535);
                } elseif ($calculatedValue === '') {
                    // Empty string (and BIFF8)
                    $num = pack('CCCvCv', 3, 0, 0, 0, 0, 65535);
                } else {
                    // Non-empty string value (or empty string BIFF5)
                    $stringValue = $calculatedValue;
                    $num = pack('CCCvCv', 0, 0, 0, 0, 0, 65535);
                }
            } else {
                // We are really not supposed to reach here
                $num = pack('d', 0);
            }
        } else {
            $num = pack('d', 0);
        }
        $grbit = 3;
        // Option flags
        $unknown = 0;
        // Must be zero
        // Strip the '=' or '@' sign at the beginning of the formula string
        if ($formula[0] == '=') {
            $formula = substr($formula, 1);
        } else {
            // Error handling
            $this->writeString($row, $col, 'Unrecognised character for formula', 0);
            return -1;
        }
        // Parse the formula using the parser in Parser.php
        try {
            $error = $this->parser->parse($formula);
            $formula = $this->parser->toReversePolish();
            $formlen = strlen($formula);
            // Length of the binary string
            $length = 22 + $formlen;
            // Length of the record data
            $header = pack('vv', $record, $length);
            $data = pack('vvv', $row, $col, $xfIndex) . $num . pack('vVv', $grbit, $unknown, $formlen);
            $this->append($header . $data . $formula);
            // Append also a STRING record if necessary
            if ($stringValue !== null) {
                $this->writeStringRecord($stringValue);
            }
            return 0;
        } catch (PhpSpreadsheetException $e) {
        }
    }
    /**
     * Write a STRING record. This.
     *
     * @param string $stringValue
     */
    private function writeStringRecord($stringValue)
    {
        $record = 519;
        // Record identifier
        $data = StringHelper::UTF8toBIFF8UnicodeLong($stringValue);
        $length = strlen($data);
        $header = pack('vv', $record, $length);
        $this->append($header . $data);
    }
    /**
     * Write a hyperlink.
     * This is comprised of two elements: the visible label and
     * the invisible link. The visible label is the same as the link unless an
     * alternative string is specified. The label is written using the
     * writeString() method. Therefore the 255 characters string limit applies.
     * $string and $format are optional.
     *
     * The hyperlink can be to a http, ftp, mail, internal sheet (not yet), or external
     * directory url.
     *
     * Returns  0 : normal termination
     *         -2 : row or column out of range
     *         -3 : long string truncated to 255 chars
     *
     * @param int $row Row
     * @param int $col Column
     * @param string $url URL string
     *
     * @return int
     */
    private function writeUrl($row, $col, $url)
    {
        // Add start row and col to arg list
        return $this->writeUrlRange($row, $col, $row, $col, $url);
    }
    /**
     * This is the more general form of writeUrl(). It allows a hyperlink to be
     * written to a range of cells. This function also decides the type of hyperlink
     * to be written. These are either, Web (http, ftp, mailto), Internal
     * (Sheet1!A1) or external ('c:\temp\foo.xls#Sheet1!A1').
     *
     * @see writeUrl()
     *
     * @param int $row1 Start row
     * @param int $col1 Start column
     * @param int $row2 End row
     * @param int $col2 End column
     * @param string $url URL string
     *
     * @return int
     */
    public function writeUrlRange($row1, $col1, $row2, $col2, $url)
    {
        // Check for internal/external sheet links or default to web link
        if (preg_match('[^internal:]', $url)) {
            return $this->writeUrlInternal($row1, $col1, $row2, $col2, $url);
        }
        if (preg_match('[^external:]', $url)) {
            return $this->writeUrlExternal($row1, $col1, $row2, $col2, $url);
        }
        return $this->writeUrlWeb($row1, $col1, $row2, $col2, $url);
    }
    /**
     * Used to write http, ftp and mailto hyperlinks.
     * The link type ($options) is 0x03 is the same as absolute dir ref without
     * sheet. However it is differentiated by the $unknown2 data stream.
     *
     * @see writeUrl()
     *
     * @param int $row1 Start row
     * @param int $col1 Start column
     * @param int $row2 End row
     * @param int $col2 End column
     * @param string $url URL string
     *
     * @return int
     */
    public function writeUrlWeb($row1, $col1, $row2, $col2, $url)
    {
        $record = 440;
        // Record identifier
        $length = 0;
        // Bytes to follow
        // Pack the undocumented parts of the hyperlink stream
        $unknown1 = pack('H*', 'D0C9EA79F9BACE118C8200AA004BA90B02000000');
        $unknown2 = pack('H*', 'E0C9EA79F9BACE118C8200AA004BA90B');
        // Pack the option flags
        $options = pack('V', 3);
        // Convert URL to a null terminated wchar string
        $url = implode(' ', preg_split('\'\'', $url, -1, PREG_SPLIT_NO_EMPTY));
        $url = $url . '   ';
        // Pack the length of the URL
        $url_len = pack('V', strlen($url));
        // Calculate the data length
        $length = 52 + strlen($url);
        // Pack the header data
        $header = pack('vv', $record, $length);
        $data = pack('vvvv', $row1, $row2, $col1, $col2);
        // Write the packed data
        $this->append($header . $data . $unknown1 . $options . $unknown2 . $url_len . $url);
        return 0;
    }
    /**
     * Used to write internal reference hyperlinks such as "Sheet1!A1".
     *
     * @see writeUrl()
     *
     * @param int $row1 Start row
     * @param int $col1 Start column
     * @param int $row2 End row
     * @param int $col2 End column
     * @param string $url URL string
     *
     * @return int
     */
    public function writeUrlInternal($row1, $col1, $row2, $col2, $url)
    {
        $record = 440;
        // Record identifier
        $length = 0;
        // Bytes to follow
        // Strip URL type
        $url = preg_replace('/^internal:/', '', $url);
        // Pack the undocumented parts of the hyperlink stream
        $unknown1 = pack('H*', 'D0C9EA79F9BACE118C8200AA004BA90B02000000');
        // Pack the option flags
        $options = pack('V', 8);
        // Convert the URL type and to a null terminated wchar string
        $url .= ' ';
        // character count
        $url_len = StringHelper::countCharacters($url);
        $url_len = pack('V', $url_len);
        $url = StringHelper::convertEncoding($url, 'UTF-16LE', 'UTF-8');
        // Calculate the data length
        $length = 36 + strlen($url);
        // Pack the header data
        $header = pack('vv', $record, $length);
        $data = pack('vvvv', $row1, $row2, $col1, $col2);
        // Write the packed data
        $this->append($header . $data . $unknown1 . $options . $url_len . $url);
        return 0;
    }
    /**
     * Write links to external directory names such as 'c:\foo.xls',
     * c:\foo.xls#Sheet1!A1', '../../foo.xls'. and '../../foo.xls#Sheet1!A1'.
     *
     * Note: Excel writes some relative links with the $dir_long string. We ignore
     * these cases for the sake of simpler code.
     *
     * @see writeUrl()
     *
     * @param int $row1 Start row
     * @param int $col1 Start column
     * @param int $row2 End row
     * @param int $col2 End column
     * @param string $url URL string
     *
     * @return int
     */
    public function writeUrlExternal($row1, $col1, $row2, $col2, $url)
    {
        // Network drives are different. We will handle them separately
        // MS/Novell network drives and shares start with \\
        if (preg_match('[^external:\\\\]', $url)) {
            return;
        }
        $record = 440;
        // Record identifier
        $length = 0;
        // Bytes to follow
        // Strip URL type and change Unix dir separator to Dos style (if needed)
        //
        $url = preg_replace('/^external:/', '', $url);
        $url = preg_replace('/\\//', '\\', $url);
        // Determine if the link is relative or absolute:
        //   relative if link contains no dir separator, "somefile.xls"
        //   relative if link starts with up-dir, "..\..\somefile.xls"
        //   otherwise, absolute
        $absolute = 0;
        // relative path
        if (preg_match('/^[A-Z]:/', $url)) {
            $absolute = 2;
        }
        $link_type = 1 | $absolute;
        // Determine if the link contains a sheet reference and change some of the
        // parameters accordingly.
        // Split the dir name and sheet name (if it exists)
        $dir_long = $url;
        if (preg_match('/\\#/', $url)) {
            $link_type |= 8;
        }
        // Pack the link type
        $link_type = pack('V', $link_type);
        // Calculate the up-level dir count e.g.. (..\..\..\ == 3)
        $up_count = preg_match_all('/\\.\\.\\\\/', $dir_long, $useless);
        $up_count = pack('v', $up_count);
        // Store the short dos dir name (null terminated)
        $dir_short = preg_replace('/\\.\\.\\\\/', '', $dir_long) . ' ';
        // Store the long dir name as a wchar string (non-null terminated)
        $dir_long = $dir_long . ' ';
        // Pack the lengths of the dir strings
        $dir_short_len = pack('V', strlen($dir_short));
        $dir_long_len = pack('V', strlen($dir_long));
        $stream_len = pack('V', 0);
        //strlen($dir_long) + 0x06);
        // Pack the undocumented parts of the hyperlink stream
        $unknown1 = pack('H*', 'D0C9EA79F9BACE118C8200AA004BA90B02000000');
        $unknown2 = pack('H*', '0303000000000000C000000000000046');
        $unknown3 = pack('H*', 'FFFFADDE000000000000000000000000000000000000000');
        $unknown4 = pack('v', 3);
        // Pack the main data stream
        $data = pack('vvvv', $row1, $row2, $col1, $col2) . $unknown1 . $link_type . $unknown2 . $up_count . $dir_short_len . $dir_short . $unknown3 . $stream_len;
        /*.
          $dir_long_len .
          $unknown4     .
          $dir_long     .
          $sheet_len    .
          $sheet        ;*/
        // Pack the header data
        $length = strlen($data);
        $header = pack('vv', $record, $length);
        // Write the packed data
        $this->append($header . $data);
        return 0;
    }
    /**
     * This method is used to set the height and format for a row.
     *
     * @param int $row The row to set
     * @param int $height Height we are giving to the row.
     *                        Use null to set XF without setting height
     * @param int $xfIndex The optional cell style Xf index to apply to the columns
     * @param bool $hidden The optional hidden attribute
     * @param int $level The optional outline level for row, in range [0,7]
     */
    private function writeRow($row, $height, $xfIndex, $hidden = false, $level = 0)
    {
        $record = 520;
        // Record identifier
        $length = 16;
        // Number of bytes to follow
        $colMic = 0;
        // First defined column
        $colMac = 0;
        // Last defined column
        $irwMac = 0;
        // Used by Excel to optimise loading
        $reserved = 0;
        // Reserved
        $grbit = 0;
        // Option flags
        $ixfe = $xfIndex;
        if ($height < 0) {
            $height = null;
        }
        // Use writeRow($row, null, $XF) to set XF format without setting height
        if ($height != null) {
            $miyRw = $height * 20;
        } else {
            $miyRw = 255;
        }
        // Set the options flags. fUnsynced is used to show that the font and row
        // heights are not compatible. This is usually the case for WriteExcel.
        // The collapsed flag 0x10 doesn't seem to be used to indicate that a row
        // is collapsed. Instead it is used to indicate that the previous row is
        // collapsed. The zero height flag, 0x20, is used to collapse a row.
        $grbit |= $level;
        if ($hidden) {
            $grbit |= 48;
        }
        if ($height !== null) {
            $grbit |= 64;
        }
        if ($xfIndex !== 15) {
            $grbit |= 128;
        }
        $grbit |= 256;
        $header = pack('vv', $record, $length);
        $data = pack('vvvvvvvv', $row, $colMic, $colMac, $miyRw, $irwMac, $reserved, $grbit, $ixfe);
        $this->append($header . $data);
    }
    /**
     * Writes Excel DIMENSIONS to define the area in which there is data.
     */
    private function writeDimensions()
    {
        $record = 512;
        // Record identifier
        $length = 14;
        $data = pack('VVvvv', $this->firstRowIndex, $this->lastRowIndex + 1, $this->firstColumnIndex, $this->lastColumnIndex + 1, 0);
        // reserved
        $header = pack('vv', $record, $length);
        $this->append($header . $data);
    }
    /**
     * Write BIFF record Window2.
     */
    private function writeWindow2()
    {
        $record = 574;
        // Record identifier
        $length = 18;
        $grbit = 182;
        // Option flags
        $rwTop = 0;
        // Top row visible in window
        $colLeft = 0;
        // Leftmost column visible in window
        // The options flags that comprise $grbit
        $fDspFmla = 0;
        // 0 - bit
        $fDspGrid = $this->phpSheet->getShowGridlines() ? 1 : 0;
        // 1
        $fDspRwCol = $this->phpSheet->getShowRowColHeaders() ? 1 : 0;
        // 2
        $fFrozen = $this->phpSheet->getFreezePane() ? 1 : 0;
        // 3
        $fDspZeros = 1;
        // 4
        $fDefaultHdr = 1;
        // 5
        $fArabic = $this->phpSheet->getRightToLeft() ? 1 : 0;
        // 6
        $fDspGuts = $this->outlineOn;
        // 7
        $fFrozenNoSplit = 0;
        // 0 - bit
        // no support in PhpSpreadsheet for selected sheet, therefore sheet is only selected if it is the active sheet
        $fSelected = $this->phpSheet === $this->phpSheet->getParent()->getActiveSheet() ? 1 : 0;
        $fPaged = 1;
        // 2
        $fPageBreakPreview = $this->phpSheet->getSheetView()->getView() === SheetView::SHEETVIEW_PAGE_BREAK_PREVIEW;
        $grbit = $fDspFmla;
        $grbit |= $fDspGrid << 1;
        $grbit |= $fDspRwCol << 2;
        $grbit |= $fFrozen << 3;
        $grbit |= $fDspZeros << 4;
        $grbit |= $fDefaultHdr << 5;
        $grbit |= $fArabic << 6;
        $grbit |= $fDspGuts << 7;
        $grbit |= $fFrozenNoSplit << 8;
        $grbit |= $fSelected << 9;
        $grbit |= $fPaged << 10;
        $grbit |= $fPageBreakPreview << 11;
        $header = pack('vv', $record, $length);
        $data = pack('vvv', $grbit, $rwTop, $colLeft);
        // FIXME !!!
        $rgbHdr = 64;
        // Row/column heading and gridline color index
        $zoom_factor_page_break = $fPageBreakPreview ? $this->phpSheet->getSheetView()->getZoomScale() : 0;
        $zoom_factor_normal = $this->phpSheet->getSheetView()->getZoomScaleNormal();
        $data .= pack('vvvvV', $rgbHdr, 0, $zoom_factor_page_break, $zoom_factor_normal, 0);
        $this->append($header . $data);
    }
    /**
     * Write BIFF record DEFAULTROWHEIGHT.
     */
    private function writeDefaultRowHeight()
    {
        $defaultRowHeight = $this->phpSheet->getDefaultRowDimension()->getRowHeight();
        if ($defaultRowHeight < 0) {
            return;
        }
        // convert to twips
        $defaultRowHeight = (int) 20 * $defaultRowHeight;
        $record = 549;
        // Record identifier
        $length = 4;
        // Number of bytes to follow
        $header = pack('vv', $record, $length);
        $data = pack('vv', 1, $defaultRowHeight);
        $this->append($header . $data);
    }
    /**
     * Write BIFF record DEFCOLWIDTH if COLINFO records are in use.
     */
    private function writeDefcol()
    {
        $defaultColWidth = 8;
        $record = 85;
        // Record identifier
        $length = 2;
        // Number of bytes to follow
        $header = pack('vv', $record, $length);
        $data = pack('v', $defaultColWidth);
        $this->append($header . $data);
    }
    /**
     * Write BIFF record COLINFO to define column widths.
     *
     * Note: The SDK says the record length is 0x0B but Excel writes a 0x0C
     * length record.
     *
     * @param array $col_array This is the only parameter received and is composed of the following:
     *                0 => First formatted column,
     *                1 => Last formatted column,
     *                2 => Col width (8.43 is Excel default),
     *                3 => The optional XF format of the column,
     *                4 => Option flags.
     *                5 => Optional outline level
     */
    private function writeColinfo($col_array)
    {
        if (isset($col_array[0])) {
            $colFirst = $col_array[0];
        }
        if (isset($col_array[1])) {
            $colLast = $col_array[1];
        }
        if (isset($col_array[2])) {
            $coldx = $col_array[2];
        } else {
            $coldx = 8.43;
        }
        if (isset($col_array[3])) {
            $xfIndex = $col_array[3];
        } else {
            $xfIndex = 15;
        }
        if (isset($col_array[4])) {
            $grbit = $col_array[4];
        } else {
            $grbit = 0;
        }
        if (isset($col_array[5])) {
            $level = $col_array[5];
        } else {
            $level = 0;
        }
        $record = 125;
        // Record identifier
        $length = 12;
        // Number of bytes to follow
        $coldx *= 256;
        // Convert to units of 1/256 of a char
        $ixfe = $xfIndex;
        $reserved = 0;
        // Reserved
        $level = max(0, min($level, 7));
        $grbit |= $level << 8;
        $header = pack('vv', $record, $length);
        $data = pack('vvvvvv', $colFirst, $colLast, $coldx, $ixfe, $grbit, $reserved);
        $this->append($header . $data);
    }
    /**
     * Write BIFF record SELECTION.
     */
    private function writeSelection()
    {
        // look up the selected cell range
        $selectedCells = Coordinate::splitRange($this->phpSheet->getSelectedCells());
        $selectedCells = $selectedCells[0];
        if (count($selectedCells) == 2) {
            list($first, $last) = $selectedCells;
        } else {
            $first = $selectedCells[0];
            $last = $selectedCells[0];
        }
        list($colFirst, $rwFirst) = Coordinate::coordinateFromString($first);
        $colFirst = Coordinate::columnIndexFromString($colFirst) - 1;
        // base 0 column index
        --$rwFirst;
        // base 0 row index
        list($colLast, $rwLast) = Coordinate::coordinateFromString($last);
        $colLast = Coordinate::columnIndexFromString($colLast) - 1;
        // base 0 column index
        --$rwLast;
        // base 0 row index
        // make sure we are not out of bounds
        $colFirst = min($colFirst, 255);
        $colLast = min($colLast, 255);
        $rwFirst = min($rwFirst, 65535);
        $rwLast = min($rwLast, 65535);
        $record = 29;
        // Record identifier
        $length = 15;
        // Number of bytes to follow
        $pnn = $this->activePane;
        // Pane position
        $rwAct = $rwFirst;
        // Active row
        $colAct = $colFirst;
        // Active column
        $irefAct = 0;
        // Active cell ref
        $cref = 1;
        // Number of refs
        if (!isset($rwLast)) {
            $rwLast = $rwFirst;
        }
        if (!isset($colLast)) {
            $colLast = $colFirst;
        }
        // Swap last row/col for first row/col as necessary
        if ($rwFirst > $rwLast) {
            list($rwFirst, $rwLast) = array($rwLast, $rwFirst);
        }
        if ($colFirst > $colLast) {
            list($colFirst, $colLast) = array($colLast, $colFirst);
        }
        $header = pack('vv', $record, $length);
        $data = pack('CvvvvvvCC', $pnn, $rwAct, $colAct, $irefAct, $cref, $rwFirst, $rwLast, $colFirst, $colLast);
        $this->append($header . $data);
    }
    /**
     * Store the MERGEDCELLS records for all ranges of merged cells.
     */
    private function writeMergedCells()
    {
        $mergeCells = $this->phpSheet->getMergeCells();
        $countMergeCells = count($mergeCells);
        if ($countMergeCells == 0) {
            return;
        }
        // maximum allowed number of merged cells per record
        $maxCountMergeCellsPerRecord = 1027;
        // record identifier
        $record = 229;
        // counter for total number of merged cells treated so far by the writer
        $i = 0;
        // counter for number of merged cells written in record currently being written
        $j = 0;
        // initialize record data
        $recordData = '';
        // loop through the merged cells
        foreach ($mergeCells as $mergeCell) {
            ++$i;
            ++$j;
            // extract the row and column indexes
            $range = Coordinate::splitRange($mergeCell);
            list($first, $last) = $range[0];
            list($firstColumn, $firstRow) = Coordinate::coordinateFromString($first);
            list($lastColumn, $lastRow) = Coordinate::coordinateFromString($last);
            $recordData .= pack('vvvv', $firstRow - 1, $lastRow - 1, Coordinate::columnIndexFromString($firstColumn) - 1, Coordinate::columnIndexFromString($lastColumn) - 1);
            // flush record if we have reached limit for number of merged cells, or reached final merged cell
            if ($j == $maxCountMergeCellsPerRecord or $i == $countMergeCells) {
                $recordData = pack('v', $j) . $recordData;
                $length = strlen($recordData);
                $header = pack('vv', $record, $length);
                $this->append($header . $recordData);
                // initialize for next record, if any
                $recordData = '';
                $j = 0;
            }
        }
    }
    /**
     * Write SHEETLAYOUT record.
     */
    private function writeSheetLayout()
    {
        if (!$this->phpSheet->isTabColorSet()) {
            return;
        }
        $recordData = pack('vvVVVvv', 2146, 0, 0, 0, 20, $this->colors[$this->phpSheet->getTabColor()->getRGB()], 0);
        $length = strlen($recordData);
        $record = 2146;
        // Record identifier
        $header = pack('vv', $record, $length);
        $this->append($header . $recordData);
    }
    /**
     * Write SHEETPROTECTION.
     */
    private function writeSheetProtection()
    {
        // record identifier
        $record = 2151;
        // prepare options
        $options = (int) (!$this->phpSheet->getProtection()->getObjects()) | (int) (!$this->phpSheet->getProtection()->getScenarios()) << 1 | (int) (!$this->phpSheet->getProtection()->getFormatCells()) << 2 | (int) (!$this->phpSheet->getProtection()->getFormatColumns()) << 3 | (int) (!$this->phpSheet->getProtection()->getFormatRows()) << 4 | (int) (!$this->phpSheet->getProtection()->getInsertColumns()) << 5 | (int) (!$this->phpSheet->getProtection()->getInsertRows()) << 6 | (int) (!$this->phpSheet->getProtection()->getInsertHyperlinks()) << 7 | (int) (!$this->phpSheet->getProtection()->getDeleteColumns()) << 8 | (int) (!$this->phpSheet->getProtection()->getDeleteRows()) << 9 | (int) (!$this->phpSheet->getProtection()->getSelectLockedCells()) << 10 | (int) (!$this->phpSheet->getProtection()->getSort()) << 11 | (int) (!$this->phpSheet->getProtection()->getAutoFilter()) << 12 | (int) (!$this->phpSheet->getProtection()->getPivotTables()) << 13 | (int) (!$this->phpSheet->getProtection()->getSelectUnlockedCells()) << 14;
        // record data
        $recordData = pack('vVVCVVvv', 2151, 0, 0, 0, 16777728, 4294967295.0, $options, 0);
        $length = strlen($recordData);
        $header = pack('vv', $record, $length);
        $this->append($header . $recordData);
    }
    /**
     * Write BIFF record RANGEPROTECTION.
     *
     * Openoffice.org's Documentaion of the Microsoft Excel File Format uses term RANGEPROTECTION for these records
     * Microsoft Office Excel 97-2007 Binary File Format Specification uses term FEAT for these records
     */
    private function writeRangeProtection()
    {
        foreach ($this->phpSheet->getProtectedCells() as $range => $password) {
            // number of ranges, e.g. 'A1:B3 C20:D25'
            $cellRanges = explode(' ', $range);
            $cref = count($cellRanges);
            $recordData = pack('vvVVvCVvVv', 2152, 0, 0, 0, 2, 0, 0, $cref, 0, 0);
            foreach ($cellRanges as $cellRange) {
                $recordData .= $this->writeBIFF8CellRangeAddressFixed($cellRange);
            }
            // the rgbFeat structure
            $recordData .= pack('VV', 0, hexdec($password));
            $recordData .= StringHelper::UTF8toBIFF8UnicodeLong('p' . md5($recordData));
            $length = strlen($recordData);
            $record = 2152;
            // Record identifier
            $header = pack('vv', $record, $length);
            $this->append($header . $recordData);
        }
    }
    /**
     * Writes the Excel BIFF PANE record.
     * The panes can either be frozen or thawed (unfrozen).
     * Frozen panes are specified in terms of an integer number of rows and columns.
     * Thawed panes are specified in terms of Excel's units for rows and columns.
     */
    private function writePanes()
    {
        $panes = array();
        if ($this->phpSheet->getFreezePane()) {
            list($column, $row) = Coordinate::coordinateFromString($this->phpSheet->getFreezePane());
            $panes[0] = Coordinate::columnIndexFromString($column) - 1;
            $panes[1] = $row - 1;
            list($leftMostColumn, $topRow) = Coordinate::coordinateFromString($this->phpSheet->getTopLeftCell());
            //Coordinates are zero-based in xls files
            $panes[2] = $topRow - 1;
            $panes[3] = Coordinate::columnIndexFromString($leftMostColumn) - 1;
        } else {
            // thaw panes
            return;
        }
        $x = isset($panes[0]) ? $panes[0] : null;
        $y = isset($panes[1]) ? $panes[1] : null;
        $rwTop = isset($panes[2]) ? $panes[2] : null;
        $colLeft = isset($panes[3]) ? $panes[3] : null;
        if (count($panes) > 4) {
            // if Active pane was received
            $pnnAct = $panes[4];
        } else {
            $pnnAct = null;
        }
        $record = 65;
        // Record identifier
        $length = 10;
        // Number of bytes to follow
        // Code specific to frozen or thawed panes.
        if ($this->phpSheet->getFreezePane()) {
            // Set default values for $rwTop and $colLeft
            if (!isset($rwTop)) {
                $rwTop = $y;
            }
            if (!isset($colLeft)) {
                $colLeft = $x;
            }
        } else {
            // Set default values for $rwTop and $colLeft
            if (!isset($rwTop)) {
                $rwTop = 0;
            }
            if (!isset($colLeft)) {
                $colLeft = 0;
            }
            // Convert Excel's row and column units to the internal units.
            // The default row height is 12.75
            // The default column width is 8.43
            // The following slope and intersection values were interpolated.
            //
            $y = 20 * $y + 255;
            $x = 113.879 * $x + 390;
        }
        // Determine which pane should be active. There is also the undocumented
        // option to override this should it be necessary: may be removed later.
        //
        if (!isset($pnnAct)) {
            if ($x != 0 && $y != 0) {
                $pnnAct = 0;
            }
            if ($x != 0 && $y == 0) {
                $pnnAct = 1;
            }
            if ($x == 0 && $y != 0) {
                $pnnAct = 2;
            }
            if ($x == 0 && $y == 0) {
                $pnnAct = 3;
            }
        }
        $this->activePane = $pnnAct;
        // Used in writeSelection
        $header = pack('vv', $record, $length);
        $data = pack('vvvvv', $x, $y, $rwTop, $colLeft, $pnnAct);
        $this->append($header . $data);
    }
    /**
     * Store the page setup SETUP BIFF record.
     */
    private function writeSetup()
    {
        $record = 161;
        // Record identifier
        $length = 34;
        // Number of bytes to follow
        $iPaperSize = $this->phpSheet->getPageSetup()->getPaperSize();
        // Paper size
        $iScale = $this->phpSheet->getPageSetup()->getScale() ? $this->phpSheet->getPageSetup()->getScale() : 100;
        // Print scaling factor
        $iPageStart = 1;
        // Starting page number
        $iFitWidth = (int) $this->phpSheet->getPageSetup()->getFitToWidth();
        // Fit to number of pages wide
        $iFitHeight = (int) $this->phpSheet->getPageSetup()->getFitToHeight();
        // Fit to number of pages high
        $grbit = 0;
        // Option flags
        $iRes = 600;
        // Print resolution
        $iVRes = 600;
        // Vertical print resolution
        $numHdr = $this->phpSheet->getPageMargins()->getHeader();
        // Header Margin
        $numFtr = $this->phpSheet->getPageMargins()->getFooter();
        // Footer Margin
        $iCopies = 1;
        // Number of copies
        $fLeftToRight = 0;
        // Print over then down
        // Page orientation
        $fLandscape = $this->phpSheet->getPageSetup()->getOrientation() == PageSetup::ORIENTATION_LANDSCAPE ? 0 : 1;
        $fNoPls = 0;
        // Setup not read from printer
        $fNoColor = 0;
        // Print black and white
        $fDraft = 0;
        // Print draft quality
        $fNotes = 0;
        // Print notes
        $fNoOrient = 0;
        // Orientation not set
        $fUsePage = 0;
        // Use custom starting page
        $grbit = $fLeftToRight;
        $grbit |= $fLandscape << 1;
        $grbit |= $fNoPls << 2;
        $grbit |= $fNoColor << 3;
        $grbit |= $fDraft << 4;
        $grbit |= $fNotes << 5;
        $grbit |= $fNoOrient << 6;
        $grbit |= $fUsePage << 7;
        $numHdr = pack('d', $numHdr);
        $numFtr = pack('d', $numFtr);
        if (self::getByteOrder()) {
            // if it's Big Endian
            $numHdr = strrev($numHdr);
            $numFtr = strrev($numFtr);
        }
        $header = pack('vv', $record, $length);
        $data1 = pack('vvvvvvvv', $iPaperSize, $iScale, $iPageStart, $iFitWidth, $iFitHeight, $grbit, $iRes, $iVRes);
        $data2 = $numHdr . $numFtr;
        $data3 = pack('v', $iCopies);
        $this->append($header . $data1 . $data2 . $data3);
    }
    /**
     * Store the header caption BIFF record.
     */
    private function writeHeader()
    {
        $record = 20;
        // Record identifier
        /* removing for now
           // need to fix character count (multibyte!)
           if (strlen($this->phpSheet->getHeaderFooter()->getOddHeader()) <= 255) {
           $str      = $this->phpSheet->getHeaderFooter()->getOddHeader();       // header string
           } else {
           $str = '';
           }
           */
        $recordData = StringHelper::UTF8toBIFF8UnicodeLong($this->phpSheet->getHeaderFooter()->getOddHeader());
        $length = strlen($recordData);
        $header = pack('vv', $record, $length);
        $this->append($header . $recordData);
    }
    /**
     * Store the footer caption BIFF record.
     */
    private function writeFooter()
    {
        $record = 21;
        // Record identifier
        /* removing for now
           // need to fix character count (multibyte!)
           if (strlen($this->phpSheet->getHeaderFooter()->getOddFooter()) <= 255) {
           $str = $this->phpSheet->getHeaderFooter()->getOddFooter();
           } else {
           $str = '';
           }
           */
        $recordData = StringHelper::UTF8toBIFF8UnicodeLong($this->phpSheet->getHeaderFooter()->getOddFooter());
        $length = strlen($recordData);
        $header = pack('vv', $record, $length);
        $this->append($header . $recordData);
    }
    /**
     * Store the horizontal centering HCENTER BIFF record.
     */
    private function writeHcenter()
    {
        $record = 131;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $fHCenter = $this->phpSheet->getPageSetup()->getHorizontalCentered() ? 1 : 0;
        // Horizontal centering
        $header = pack('vv', $record, $length);
        $data = pack('v', $fHCenter);
        $this->append($header . $data);
    }
    /**
     * Store the vertical centering VCENTER BIFF record.
     */
    private function writeVcenter()
    {
        $record = 132;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $fVCenter = $this->phpSheet->getPageSetup()->getVerticalCentered() ? 1 : 0;
        // Horizontal centering
        $header = pack('vv', $record, $length);
        $data = pack('v', $fVCenter);
        $this->append($header . $data);
    }
    /**
     * Store the LEFTMARGIN BIFF record.
     */
    private function writeMarginLeft()
    {
        $record = 38;
        // Record identifier
        $length = 8;
        // Bytes to follow
        $margin = $this->phpSheet->getPageMargins()->getLeft();
        // Margin in inches
        $header = pack('vv', $record, $length);
        $data = pack('d', $margin);
        if (self::getByteOrder()) {
            // if it's Big Endian
            $data = strrev($data);
        }
        $this->append($header . $data);
    }
    /**
     * Store the RIGHTMARGIN BIFF record.
     */
    private function writeMarginRight()
    {
        $record = 39;
        // Record identifier
        $length = 8;
        // Bytes to follow
        $margin = $this->phpSheet->getPageMargins()->getRight();
        // Margin in inches
        $header = pack('vv', $record, $length);
        $data = pack('d', $margin);
        if (self::getByteOrder()) {
            // if it's Big Endian
            $data = strrev($data);
        }
        $this->append($header . $data);
    }
    /**
     * Store the TOPMARGIN BIFF record.
     */
    private function writeMarginTop()
    {
        $record = 40;
        // Record identifier
        $length = 8;
        // Bytes to follow
        $margin = $this->phpSheet->getPageMargins()->getTop();
        // Margin in inches
        $header = pack('vv', $record, $length);
        $data = pack('d', $margin);
        if (self::getByteOrder()) {
            // if it's Big Endian
            $data = strrev($data);
        }
        $this->append($header . $data);
    }
    /**
     * Store the BOTTOMMARGIN BIFF record.
     */
    private function writeMarginBottom()
    {
        $record = 41;
        // Record identifier
        $length = 8;
        // Bytes to follow
        $margin = $this->phpSheet->getPageMargins()->getBottom();
        // Margin in inches
        $header = pack('vv', $record, $length);
        $data = pack('d', $margin);
        if (self::getByteOrder()) {
            // if it's Big Endian
            $data = strrev($data);
        }
        $this->append($header . $data);
    }
    /**
     * Write the PRINTHEADERS BIFF record.
     */
    private function writePrintHeaders()
    {
        $record = 42;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $fPrintRwCol = $this->printHeaders;
        // Boolean flag
        $header = pack('vv', $record, $length);
        $data = pack('v', $fPrintRwCol);
        $this->append($header . $data);
    }
    /**
     * Write the PRINTGRIDLINES BIFF record. Must be used in conjunction with the
     * GRIDSET record.
     */
    private function writePrintGridlines()
    {
        $record = 43;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $fPrintGrid = $this->phpSheet->getPrintGridlines() ? 1 : 0;
        // Boolean flag
        $header = pack('vv', $record, $length);
        $data = pack('v', $fPrintGrid);
        $this->append($header . $data);
    }
    /**
     * Write the GRIDSET BIFF record. Must be used in conjunction with the
     * PRINTGRIDLINES record.
     */
    private function writeGridset()
    {
        $record = 130;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $fGridSet = !$this->phpSheet->getPrintGridlines();
        // Boolean flag
        $header = pack('vv', $record, $length);
        $data = pack('v', $fGridSet);
        $this->append($header . $data);
    }
    /**
     * Write the AUTOFILTERINFO BIFF record. This is used to configure the number of autofilter select used in the sheet.
     */
    private function writeAutoFilterInfo()
    {
        $record = 157;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $rangeBounds = Coordinate::rangeBoundaries($this->phpSheet->getAutoFilter()->getRange());
        $iNumFilters = 1 + $rangeBounds[1][0] - $rangeBounds[0][0];
        $header = pack('vv', $record, $length);
        $data = pack('v', $iNumFilters);
        $this->append($header . $data);
    }
    /**
     * Write the GUTS BIFF record. This is used to configure the gutter margins
     * where Excel outline symbols are displayed. The visibility of the gutters is
     * controlled by a flag in WSBOOL.
     *
     * @see writeWsbool()
     */
    private function writeGuts()
    {
        $record = 128;
        // Record identifier
        $length = 8;
        // Bytes to follow
        $dxRwGut = 0;
        // Size of row gutter
        $dxColGut = 0;
        // Size of col gutter
        // determine maximum row outline level
        $maxRowOutlineLevel = 0;
        foreach ($this->phpSheet->getRowDimensions() as $rowDimension) {
            $maxRowOutlineLevel = max($maxRowOutlineLevel, $rowDimension->getOutlineLevel());
        }
        $col_level = 0;
        // Calculate the maximum column outline level. The equivalent calculation
        // for the row outline level is carried out in writeRow().
        $colcount = count($this->columnInfo);
        for ($i = 0; $i < $colcount; ++$i) {
            $col_level = max($this->columnInfo[$i][5], $col_level);
        }
        // Set the limits for the outline levels (0 <= x <= 7).
        $col_level = max(0, min($col_level, 7));
        // The displayed level is one greater than the max outline levels
        if ($maxRowOutlineLevel) {
            ++$maxRowOutlineLevel;
        }
        if ($col_level) {
            ++$col_level;
        }
        $header = pack('vv', $record, $length);
        $data = pack('vvvv', $dxRwGut, $dxColGut, $maxRowOutlineLevel, $col_level);
        $this->append($header . $data);
    }
    /**
     * Write the WSBOOL BIFF record, mainly for fit-to-page. Used in conjunction
     * with the SETUP record.
     */
    private function writeWsbool()
    {
        $record = 129;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $grbit = 0;
        // The only option that is of interest is the flag for fit to page. So we
        // set all the options in one go.
        //
        // Set the option flags
        $grbit |= 1;
        // Auto page breaks visible
        if ($this->outlineStyle) {
            $grbit |= 32;
        }
        if ($this->phpSheet->getShowSummaryBelow()) {
            $grbit |= 64;
        }
        if ($this->phpSheet->getShowSummaryRight()) {
            $grbit |= 128;
        }
        if ($this->phpSheet->getPageSetup()->getFitToPage()) {
            $grbit |= 256;
        }
        if ($this->outlineOn) {
            $grbit |= 1024;
        }
        $header = pack('vv', $record, $length);
        $data = pack('v', $grbit);
        $this->append($header . $data);
    }
    /**
     * Write the HORIZONTALPAGEBREAKS and VERTICALPAGEBREAKS BIFF records.
     */
    private function writeBreaks()
    {
        // initialize
        $vbreaks = array();
        $hbreaks = array();
        foreach ($this->phpSheet->getBreaks() as $cell => $breakType) {
            // Fetch coordinates
            $coordinates = Coordinate::coordinateFromString($cell);
            // Decide what to do by the type of break
            switch ($breakType) {
                case \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_COLUMN:
                    // Add to list of vertical breaks
                    $vbreaks[] = Coordinate::columnIndexFromString($coordinates[0]) - 1;
                    break;
                case \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_ROW:
                    // Add to list of horizontal breaks
                    $hbreaks[] = $coordinates[1];
                    break;
                case \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_NONE:
                default:
                    // Nothing to do
                    break;
            }
        }
        //horizontal page breaks
        if (!empty($hbreaks)) {
            // Sort and filter array of page breaks
            sort($hbreaks, SORT_NUMERIC);
            if ($hbreaks[0] == 0) {
                // don't use first break if it's 0
                array_shift($hbreaks);
            }
            $record = 27;
            // Record identifier
            $cbrk = count($hbreaks);
            // Number of page breaks
            $length = 2 + 6 * $cbrk;
            // Bytes to follow
            $header = pack('vv', $record, $length);
            $data = pack('v', $cbrk);
            // Append each page break
            foreach ($hbreaks as $hbreak) {
                $data .= pack('vvv', $hbreak, 0, 255);
            }
            $this->append($header . $data);
        }
        // vertical page breaks
        if (!empty($vbreaks)) {
            // 1000 vertical pagebreaks appears to be an internal Excel 5 limit.
            // It is slightly higher in Excel 97/200, approx. 1026
            $vbreaks = array_slice($vbreaks, 0, 1000);
            // Sort and filter array of page breaks
            sort($vbreaks, SORT_NUMERIC);
            if ($vbreaks[0] == 0) {
                // don't use first break if it's 0
                array_shift($vbreaks);
            }
            $record = 26;
            // Record identifier
            $cbrk = count($vbreaks);
            // Number of page breaks
            $length = 2 + 6 * $cbrk;
            // Bytes to follow
            $header = pack('vv', $record, $length);
            $data = pack('v', $cbrk);
            // Append each page break
            foreach ($vbreaks as $vbreak) {
                $data .= pack('vvv', $vbreak, 0, 65535);
            }
            $this->append($header . $data);
        }
    }
    /**
     * Set the Biff PROTECT record to indicate that the worksheet is protected.
     */
    private function writeProtect()
    {
        // Exit unless sheet protection has been specified
        if (!$this->phpSheet->getProtection()->getSheet()) {
            return;
        }
        $record = 18;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $fLock = 1;
        // Worksheet is protected
        $header = pack('vv', $record, $length);
        $data = pack('v', $fLock);
        $this->append($header . $data);
    }
    /**
     * Write SCENPROTECT.
     */
    private function writeScenProtect()
    {
        // Exit if sheet protection is not active
        if (!$this->phpSheet->getProtection()->getSheet()) {
            return;
        }
        // Exit if scenarios are not protected
        if (!$this->phpSheet->getProtection()->getScenarios()) {
            return;
        }
        $record = 221;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $header = pack('vv', $record, $length);
        $data = pack('v', 1);
        $this->append($header . $data);
    }
    /**
     * Write OBJECTPROTECT.
     */
    private function writeObjectProtect()
    {
        // Exit if sheet protection is not active
        if (!$this->phpSheet->getProtection()->getSheet()) {
            return;
        }
        // Exit if objects are not protected
        if (!$this->phpSheet->getProtection()->getObjects()) {
            return;
        }
        $record = 99;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $header = pack('vv', $record, $length);
        $data = pack('v', 1);
        $this->append($header . $data);
    }
    /**
     * Write the worksheet PASSWORD record.
     */
    private function writePassword()
    {
        // Exit unless sheet protection and password have been specified
        if (!$this->phpSheet->getProtection()->getSheet() || !$this->phpSheet->getProtection()->getPassword()) {
            return;
        }
        $record = 19;
        // Record identifier
        $length = 2;
        // Bytes to follow
        $wPassword = hexdec($this->phpSheet->getProtection()->getPassword());
        // Encoded password
        $header = pack('vv', $record, $length);
        $data = pack('v', $wPassword);
        $this->append($header . $data);
    }
    /**
     * Insert a 24bit bitmap image in a worksheet.
     *
     * @param int $row The row we are going to insert the bitmap into
     * @param int $col The column we are going to insert the bitmap into
     * @param mixed $bitmap The bitmap filename or GD-image resource
     * @param int $x the horizontal position (offset) of the image inside the cell
     * @param int $y the vertical position (offset) of the image inside the cell
     * @param float $scale_x The horizontal scale
     * @param float $scale_y The vertical scale
     */
    public function insertBitmap($row, $col, $bitmap, $x = 0, $y = 0, $scale_x = 1, $scale_y = 1)
    {
        $bitmap_array = is_resource($bitmap) ? $this->processBitmapGd($bitmap) : $this->processBitmap($bitmap);
        list($width, $height, $size, $data) = $bitmap_array;
        // Scale the frame of the image.
        $width *= $scale_x;
        $height *= $scale_y;
        // Calculate the vertices of the image and write the OBJ record
        $this->positionImage($col, $row, $x, $y, $width, $height);
        // Write the IMDATA record to store the bitmap data
        $record = 127;
        $length = 8 + $size;
        $cf = 9;
        $env = 1;
        $lcb = $size;
        $header = pack('vvvvV', $record, $length, $cf, $env, $lcb);
        $this->append($header . $data);
    }
    /**
     * Calculate the vertices that define the position of the image as required by
     * the OBJ record.
     *
     *         +------------+------------+
     *         |     A      |      B     |
     *   +-----+------------+------------+
     *   |     |(x1,y1)     |            |
     *   |  1  |(A1)._______|______      |
     *   |     |    |              |     |
     *   |     |    |              |     |
     *   +-----+----|    BITMAP    |-----+
     *   |     |    |              |     |
     *   |  2  |    |______________.     |
     *   |     |            |        (B2)|
     *   |     |            |     (x2,y2)|
     *   +---- +------------+------------+
     *
     * Example of a bitmap that covers some of the area from cell A1 to cell B2.
     *
     * Based on the width and height of the bitmap we need to calculate 8 vars:
     *     $col_start, $row_start, $col_end, $row_end, $x1, $y1, $x2, $y2.
     * The width and height of the cells are also variable and have to be taken into
     * account.
     * The values of $col_start and $row_start are passed in from the calling
     * function. The values of $col_end and $row_end are calculated by subtracting
     * the width and height of the bitmap from the width and height of the
     * underlying cells.
     * The vertices are expressed as a percentage of the underlying cell width as
     * follows (rhs values are in pixels):
     *
     *       x1 = X / W *1024
     *       y1 = Y / H *256
     *       x2 = (X-1) / W *1024
     *       y2 = (Y-1) / H *256
     *
     *       Where:  X is distance from the left side of the underlying cell
     *               Y is distance from the top of the underlying cell
     *               W is the width of the cell
     *               H is the height of the cell
     * The SDK incorrectly states that the height should be expressed as a
     *        percentage of 1024.
     *
     * @param int $col_start Col containing upper left corner of object
     * @param int $row_start Row containing top left corner of object
     * @param int $x1 Distance to left side of object
     * @param int $y1 Distance to top of object
     * @param int $width Width of image frame
     * @param int $height Height of image frame
     */
    public function positionImage($col_start, $row_start, $x1, $y1, $width, $height)
    {
        // Initialise end cell to the same as the start cell
        $col_end = $col_start;
        // Col containing lower right corner of object
        $row_end = $row_start;
        // Row containing bottom right corner of object
        // Zero the specified offset if greater than the cell dimensions
        if ($x1 >= Xls::sizeCol($this->phpSheet, Coordinate::stringFromColumnIndex($col_start + 1))) {
            $x1 = 0;
        }
        if ($y1 >= Xls::sizeRow($this->phpSheet, $row_start + 1)) {
            $y1 = 0;
        }
        $width = $width + $x1 - 1;
        $height = $height + $y1 - 1;
        // Subtract the underlying cell widths to find the end cell of the image
        while ($width >= Xls::sizeCol($this->phpSheet, Coordinate::stringFromColumnIndex($col_end + 1))) {
            $width -= Xls::sizeCol($this->phpSheet, Coordinate::stringFromColumnIndex($col_end + 1));
            ++$col_end;
        }
        // Subtract the underlying cell heights to find the end cell of the image
        while ($height >= Xls::sizeRow($this->phpSheet, $row_end + 1)) {
            $height -= Xls::sizeRow($this->phpSheet, $row_end + 1);
            ++$row_end;
        }
        // Bitmap isn't allowed to start or finish in a hidden cell, i.e. a cell
        // with zero eight or width.
        //
        if (Xls::sizeCol($this->phpSheet, Coordinate::stringFromColumnIndex($col_start + 1)) == 0) {
            return;
        }
        if (Xls::sizeCol($this->phpSheet, Coordinate::stringFromColumnIndex($col_end + 1)) == 0) {
            return;
        }
        if (Xls::sizeRow($this->phpSheet, $row_start + 1) == 0) {
            return;
        }
        if (Xls::sizeRow($this->phpSheet, $row_end + 1) == 0) {
            return;
        }
        // Convert the pixel values to the percentage value expected by Excel
        $x1 = $x1 / Xls::sizeCol($this->phpSheet, Coordinate::stringFromColumnIndex($col_start + 1)) * 1024;
        $y1 = $y1 / Xls::sizeRow($this->phpSheet, $row_start + 1) * 256;
        $x2 = $width / Xls::sizeCol($this->phpSheet, Coordinate::stringFromColumnIndex($col_end + 1)) * 1024;
        // Distance to right side of object
        $y2 = $height / Xls::sizeRow($this->phpSheet, $row_end + 1) * 256;
        // Distance to bottom of object
        $this->writeObjPicture($col_start, $x1, $row_start, $y1, $col_end, $x2, $row_end, $y2);
    }
    /**
     * Store the OBJ record that precedes an IMDATA record. This could be generalise
     * to support other Excel objects.
     *
     * @param int $colL Column containing upper left corner of object
     * @param int $dxL Distance from left side of cell
     * @param int $rwT Row containing top left corner of object
     * @param int $dyT Distance from top of cell
     * @param int $colR Column containing lower right corner of object
     * @param int $dxR Distance from right of cell
     * @param int $rwB Row containing bottom right corner of object
     * @param int $dyB Distance from bottom of cell
     */
    private function writeObjPicture($colL, $dxL, $rwT, $dyT, $colR, $dxR, $rwB, $dyB)
    {
        $record = 93;
        // Record identifier
        $length = 60;
        // Bytes to follow
        $cObj = 1;
        // Count of objects in file (set to 1)
        $OT = 8;
        // Object type. 8 = Picture
        $id = 1;
        // Object ID
        $grbit = 1556;
        // Option flags
        $cbMacro = 0;
        // Length of FMLA structure
        $Reserved1 = 0;
        // Reserved
        $Reserved2 = 0;
        // Reserved
        $icvBack = 9;
        // Background colour
        $icvFore = 9;
        // Foreground colour
        $fls = 0;
        // Fill pattern
        $fAuto = 0;
        // Automatic fill
        $icv = 8;
        // Line colour
        $lns = 255;
        // Line style
        $lnw = 1;
        // Line weight
        $fAutoB = 0;
        // Automatic border
        $frs = 0;
        // Frame style
        $cf = 9;
        // Image format, 9 = bitmap
        $Reserved3 = 0;
        // Reserved
        $cbPictFmla = 0;
        // Length of FMLA structure
        $Reserved4 = 0;
        // Reserved
        $grbit2 = 1;
        // Option flags
        $Reserved5 = 0;
        // Reserved
        $header = pack('vv', $record, $length);
        $data = pack('V', $cObj);
        $data .= pack('v', $OT);
        $data .= pack('v', $id);
        $data .= pack('v', $grbit);
        $data .= pack('v', $colL);
        $data .= pack('v', $dxL);
        $data .= pack('v', $rwT);
        $data .= pack('v', $dyT);
        $data .= pack('v', $colR);
        $data .= pack('v', $dxR);
        $data .= pack('v', $rwB);
        $data .= pack('v', $dyB);
        $data .= pack('v', $cbMacro);
        $data .= pack('V', $Reserved1);
        $data .= pack('v', $Reserved2);
        $data .= pack('C', $icvBack);
        $data .= pack('C', $icvFore);
        $data .= pack('C', $fls);
        $data .= pack('C', $fAuto);
        $data .= pack('C', $icv);
        $data .= pack('C', $lns);
        $data .= pack('C', $lnw);
        $data .= pack('C', $fAutoB);
        $data .= pack('v', $frs);
        $data .= pack('V', $cf);
        $data .= pack('v', $Reserved3);
        $data .= pack('v', $cbPictFmla);
        $data .= pack('v', $Reserved4);
        $data .= pack('v', $grbit2);
        $data .= pack('V', $Reserved5);
        $this->append($header . $data);
    }
    /**
     * Convert a GD-image into the internal format.
     *
     * @param resource $image The image to process
     *
     * @return array Array with data and properties of the bitmap
     */
    public function processBitmapGd($image)
    {
        $width = imagesx($image);
        $height = imagesy($image);
        $data = pack('Vvvvv', 12, $width, $height, 1, 24);
        for ($j = $height; --$j;) {
            for ($i = 0; $i < $width; ++$i) {
                $color = imagecolorsforindex($image, imagecolorat($image, $i, $j));
                foreach (array('red', 'green', 'blue') as $key) {
                    $color[$key] = $color[$key] + round((255 - $color[$key]) * $color['alpha'] / 127);
                }
                $data .= chr($color['blue']) . chr($color['green']) . chr($color['red']);
            }
            if (3 * $width % 4) {
                $data .= str_repeat(' ', 4 - 3 * $width % 4);
            }
        }
        return array($width, $height, strlen($data), $data);
    }
    /**
     * Convert a 24 bit bitmap into the modified internal format used by Windows.
     * This is described in BITMAPCOREHEADER and BITMAPCOREINFO structures in the
     * MSDN library.
     *
     * @param string $bitmap The bitmap to process
     *
     * @return array Array with data and properties of the bitmap
     */
    public function processBitmap($bitmap)
    {
        // Open file.
        $bmp_fd = @fopen($bitmap, 'rb');
        if (!$bmp_fd) {
            throw new WriterException("Couldn't import {$bitmap}");
        }
        // Slurp the file into a string.
        $data = fread($bmp_fd, filesize($bitmap));
        // Check that the file is big enough to be a bitmap.
        if (strlen($data) <= 54) {
            throw new WriterException("{$bitmap} doesn't contain enough data.\n");
        }
        // The first 2 bytes are used to identify the bitmap.
        $identity = unpack('A2ident', $data);
        if ($identity['ident'] != 'BM') {
            throw new WriterException("{$bitmap} doesn't appear to be a valid bitmap image.\n");
        }
        // Remove bitmap data: ID.
        $data = substr($data, 2);
        // Read and remove the bitmap size. This is more reliable than reading
        // the data size at offset 0x22.
        //
        $size_array = unpack('Vsa', substr($data, 0, 4));
        $size = $size_array['sa'];
        $data = substr($data, 4);
        $size -= 54;
        // Subtract size of bitmap header.
        $size += 12;
        // Add size of BIFF header.
        // Remove bitmap data: reserved, offset, header length.
        $data = substr($data, 12);
        // Read and remove the bitmap width and height. Verify the sizes.
        $width_and_height = unpack('V2', substr($data, 0, 8));
        $width = $width_and_height[1];
        $height = $width_and_height[2];
        $data = substr($data, 8);
        if ($width > 65535) {
            throw new WriterException("{$bitmap}: largest image width supported is 65k.\n");
        }
        if ($height > 65535) {
            throw new WriterException("{$bitmap}: largest image height supported is 65k.\n");
        }
        // Read and remove the bitmap planes and bpp data. Verify them.
        $planes_and_bitcount = unpack('v2', substr($data, 0, 4));
        $data = substr($data, 4);
        if ($planes_and_bitcount[2] != 24) {
            // Bitcount
            throw new WriterException("{$bitmap} isn't a 24bit true color bitmap.\n");
        }
        if ($planes_and_bitcount[1] != 1) {
            throw new WriterException("{$bitmap}: only 1 plane supported in bitmap image.\n");
        }
        // Read and remove the bitmap compression. Verify compression.
        $compression = unpack('Vcomp', substr($data, 0, 4));
        $data = substr($data, 4);
        if ($compression['comp'] != 0) {
            throw new WriterException("{$bitmap}: compression not supported in bitmap image.\n");
        }
        // Remove bitmap data: data size, hres, vres, colours, imp. colours.
        $data = substr($data, 20);
        // Add the BITMAPCOREHEADER data
        $header = pack('Vvvvv', 12, $width, $height, 1, 24);
        $data = $header . $data;
        return array($width, $height, $size, $data);
    }
    /**
     * Store the window zoom factor. This should be a reduced fraction but for
     * simplicity we will store all fractions with a numerator of 100.
     */
    private function writeZoom()
    {
        // If scale is 100 we don't need to write a record
        if ($this->phpSheet->getSheetView()->getZoomScale() == 100) {
            return;
        }
        $record = 160;
        // Record identifier
        $length = 4;
        // Bytes to follow
        $header = pack('vv', $record, $length);
        $data = pack('vv', $this->phpSheet->getSheetView()->getZoomScale(), 100);
        $this->append($header . $data);
    }
    /**
     * Get Escher object.
     *
     * @return \PhpOffice\PhpSpreadsheet\Shared\Escher
     */
    public function getEscher()
    {
        return $this->escher;
    }
    /**
     * Set Escher object.
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\Escher $pValue
     */
    public function setEscher(\PhpOffice\PhpSpreadsheet\Shared\Escher $pValue = null)
    {
        $this->escher = $pValue;
    }
    /**
     * Write MSODRAWING record.
     */
    private function writeMsoDrawing()
    {
        // write the Escher stream if necessary
        if (isset($this->escher)) {
            $writer = new Escher($this->escher);
            $data = $writer->close();
            $spOffsets = $writer->getSpOffsets();
            $spTypes = $writer->getSpTypes();
            // write the neccesary MSODRAWING, OBJ records
            // split the Escher stream
            $spOffsets[0] = 0;
            $nm = count($spOffsets) - 1;
            // number of shapes excluding first shape
            for ($i = 1; $i <= $nm; ++$i) {
                // MSODRAWING record
                $record = 236;
                // Record identifier
                // chunk of Escher stream for one shape
                $dataChunk = substr($data, $spOffsets[$i - 1], $spOffsets[$i] - $spOffsets[$i - 1]);
                $length = strlen($dataChunk);
                $header = pack('vv', $record, $length);
                $this->append($header . $dataChunk);
                // OBJ record
                $record = 93;
                // record identifier
                $objData = '';
                // ftCmo
                if ($spTypes[$i] == 201) {
                    // Add ftCmo (common object data) subobject
                    $objData .= pack('vvvvvVVV', 21, 18, 20, $i, 8449, 0, 0, 0);
                    // Add ftSbs Scroll bar subobject
                    $objData .= pack('vv', 12, 20);
                    $objData .= pack('H*', '0000000000000000640001000A00000010000100');
                    // Add ftLbsData (List box data) subobject
                    $objData .= pack('vv', 19, 8174);
                    $objData .= pack('H*', '00000000010001030000020008005700');
                } else {
                    // Add ftCmo (common object data) subobject
                    $objData .= pack('vvvvvVVV', 21, 18, 8, $i, 24593, 0, 0, 0);
                }
                // ftEnd
                $objData .= pack('vv', 0, 0);
                $length = strlen($objData);
                $header = pack('vv', $record, $length);
                $this->append($header . $objData);
            }
        }
    }
    /**
     * Store the DATAVALIDATIONS and DATAVALIDATION records.
     */
    private function writeDataValidity()
    {
        // Datavalidation collection
        $dataValidationCollection = $this->phpSheet->getDataValidationCollection();
        // Write data validations?
        if (!empty($dataValidationCollection)) {
            // DATAVALIDATIONS record
            $record = 434;
            // Record identifier
            $length = 18;
            // Bytes to follow
            $grbit = 0;
            // Prompt box at cell, no cached validity data at DV records
            $horPos = 0;
            // Horizontal position of prompt box, if fixed position
            $verPos = 0;
            // Vertical position of prompt box, if fixed position
            $objId = 4294967295.0;
            // Object identifier of drop down arrow object, or -1 if not visible
            $header = pack('vv', $record, $length);
            $data = pack('vVVVV', $grbit, $horPos, $verPos, $objId, count($dataValidationCollection));
            $this->append($header . $data);
            // DATAVALIDATION records
            $record = 446;
            // Record identifier
            foreach ($dataValidationCollection as $cellCoordinate => $dataValidation) {
                // initialize record data
                $data = '';
                // options
                $options = 0;
                // data type
                $type = 0;
                switch ($dataValidation->getType()) {
                    case DataValidation::TYPE_NONE:
                        $type = 0;
                        break;
                    case DataValidation::TYPE_WHOLE:
                        $type = 1;
                        break;
                    case DataValidation::TYPE_DECIMAL:
                        $type = 2;
                        break;
                    case DataValidation::TYPE_LIST:
                        $type = 3;
                        break;
                    case DataValidation::TYPE_DATE:
                        $type = 4;
                        break;
                    case DataValidation::TYPE_TIME:
                        $type = 5;
                        break;
                    case DataValidation::TYPE_TEXTLENGTH:
                        $type = 6;
                        break;
                    case DataValidation::TYPE_CUSTOM:
                        $type = 7;
                        break;
                }
                $options |= $type << 0;
                // error style
                $errorStyle = 0;
                switch ($dataValidation->getErrorStyle()) {
                    case DataValidation::STYLE_STOP:
                        $errorStyle = 0;
                        break;
                    case DataValidation::STYLE_WARNING:
                        $errorStyle = 1;
                        break;
                    case DataValidation::STYLE_INFORMATION:
                        $errorStyle = 2;
                        break;
                }
                $options |= $errorStyle << 4;
                // explicit formula?
                if ($type == 3 && preg_match('/^\\".*\\"$/', $dataValidation->getFormula1())) {
                    $options |= 1 << 7;
                }
                // empty cells allowed
                $options |= $dataValidation->getAllowBlank() << 8;
                // show drop down
                $options |= !$dataValidation->getShowDropDown() << 9;
                // show input message
                $options |= $dataValidation->getShowInputMessage() << 18;
                // show error message
                $options |= $dataValidation->getShowErrorMessage() << 19;
                // condition operator
                $operator = 0;
                switch ($dataValidation->getOperator()) {
                    case DataValidation::OPERATOR_BETWEEN:
                        $operator = 0;
                        break;
                    case DataValidation::OPERATOR_NOTBETWEEN:
                        $operator = 1;
                        break;
                    case DataValidation::OPERATOR_EQUAL:
                        $operator = 2;
                        break;
                    case DataValidation::OPERATOR_NOTEQUAL:
                        $operator = 3;
                        break;
                    case DataValidation::OPERATOR_GREATERTHAN:
                        $operator = 4;
                        break;
                    case DataValidation::OPERATOR_LESSTHAN:
                        $operator = 5;
                        break;
                    case DataValidation::OPERATOR_GREATERTHANOREQUAL:
                        $operator = 6;
                        break;
                    case DataValidation::OPERATOR_LESSTHANOREQUAL:
                        $operator = 7;
                        break;
                }
                $options |= $operator << 20;
                $data = pack('V', $options);
                // prompt title
                $promptTitle = $dataValidation->getPromptTitle() !== '' ? $dataValidation->getPromptTitle() : chr(0);
                $data .= StringHelper::UTF8toBIFF8UnicodeLong($promptTitle);
                // error title
                $errorTitle = $dataValidation->getErrorTitle() !== '' ? $dataValidation->getErrorTitle() : chr(0);
                $data .= StringHelper::UTF8toBIFF8UnicodeLong($errorTitle);
                // prompt text
                $prompt = $dataValidation->getPrompt() !== '' ? $dataValidation->getPrompt() : chr(0);
                $data .= StringHelper::UTF8toBIFF8UnicodeLong($prompt);
                // error text
                $error = $dataValidation->getError() !== '' ? $dataValidation->getError() : chr(0);
                $data .= StringHelper::UTF8toBIFF8UnicodeLong($error);
                // formula 1
                try {
                    $formula1 = $dataValidation->getFormula1();
                    if ($type == 3) {
                        // list type
                        $formula1 = str_replace(',', chr(0), $formula1);
                    }
                    $this->parser->parse($formula1);
                    $formula1 = $this->parser->toReversePolish();
                    $sz1 = strlen($formula1);
                } catch (PhpSpreadsheetException $e) {
                    $sz1 = 0;
                    $formula1 = '';
                }
                $data .= pack('vv', $sz1, 0);
                $data .= $formula1;
                // formula 2
                try {
                    $formula2 = $dataValidation->getFormula2();
                    if ($formula2 === '') {
                        throw new WriterException('No formula2');
                    }
                    $this->parser->parse($formula2);
                    $formula2 = $this->parser->toReversePolish();
                    $sz2 = strlen($formula2);
                } catch (PhpSpreadsheetException $e) {
                    $sz2 = 0;
                    $formula2 = '';
                }
                $data .= pack('vv', $sz2, 0);
                $data .= $formula2;
                // cell range address list
                $data .= pack('v', 1);
                $data .= $this->writeBIFF8CellRangeAddressFixed($cellCoordinate);
                $length = strlen($data);
                $header = pack('vv', $record, $length);
                $this->append($header . $data);
            }
        }
    }
    /**
     * Map Error code.
     *
     * @param string $errorCode
     *
     * @return int
     */
    private static function mapErrorCode($errorCode)
    {
        switch ($errorCode) {
            case '#NULL!':
                return 0;
            case '#DIV/0!':
                return 7;
            case '#VALUE!':
                return 15;
            case '#REF!':
                return 23;
            case '#NAME?':
                return 29;
            case '#NUM!':
                return 36;
            case '#N/A':
                return 42;
        }
        return 0;
    }
    /**
     * Write PLV Record.
     */
    private function writePageLayoutView()
    {
        $record = 2187;
        // Record identifier
        $length = 16;
        // Bytes to follow
        $rt = 2187;
        // 2
        $grbitFrt = 0;
        // 2
        $reserved = 0;
        // 8
        $wScalvePLV = $this->phpSheet->getSheetView()->getZoomScale();
        // 2
        // The options flags that comprise $grbit
        if ($this->phpSheet->getSheetView()->getView() == SheetView::SHEETVIEW_PAGE_LAYOUT) {
            $fPageLayoutView = 1;
        } else {
            $fPageLayoutView = 0;
        }
        $fRulerVisible = 0;
        $fWhitespaceHidden = 0;
        $grbit = $fPageLayoutView;
        // 2
        $grbit |= $fRulerVisible << 1;
        $grbit |= $fWhitespaceHidden << 3;
        $header = pack('vv', $record, $length);
        $data = pack('vvVVvv', $rt, $grbitFrt, 0, 0, $wScalvePLV, $grbit);
        $this->append($header . $data);
    }
    /**
     * Write CFRule Record.
     *
     * @param Conditional $conditional
     */
    private function writeCFRule(Conditional $conditional)
    {
        $record = 433;
        // Record identifier
        // $type : Type of the CF
        // $operatorType : Comparison operator
        if ($conditional->getConditionType() == Conditional::CONDITION_EXPRESSION) {
            $type = 2;
            $operatorType = 0;
        } elseif ($conditional->getConditionType() == Conditional::CONDITION_CELLIS) {
            $type = 1;
            switch ($conditional->getOperatorType()) {
                case Conditional::OPERATOR_NONE:
                    $operatorType = 0;
                    break;
                case Conditional::OPERATOR_EQUAL:
                    $operatorType = 3;
                    break;
                case Conditional::OPERATOR_GREATERTHAN:
                    $operatorType = 5;
                    break;
                case Conditional::OPERATOR_GREATERTHANOREQUAL:
                    $operatorType = 7;
                    break;
                case Conditional::OPERATOR_LESSTHAN:
                    $operatorType = 6;
                    break;
                case Conditional::OPERATOR_LESSTHANOREQUAL:
                    $operatorType = 8;
                    break;
                case Conditional::OPERATOR_NOTEQUAL:
                    $operatorType = 4;
                    break;
                case Conditional::OPERATOR_BETWEEN:
                    $operatorType = 1;
                    break;
            }
        }
        // $szValue1 : size of the formula data for first value or formula
        // $szValue2 : size of the formula data for second value or formula
        $arrConditions = $conditional->getConditions();
        $numConditions = count($arrConditions);
        if ($numConditions == 1) {
            $szValue1 = $arrConditions[0] <= 65535 ? 3 : 0;
            $szValue2 = 0;
            $operand1 = pack('Cv', 30, $arrConditions[0]);
            $operand2 = null;
        } elseif ($numConditions == 2 && $conditional->getOperatorType() == Conditional::OPERATOR_BETWEEN) {
            $szValue1 = $arrConditions[0] <= 65535 ? 3 : 0;
            $szValue2 = $arrConditions[1] <= 65535 ? 3 : 0;
            $operand1 = pack('Cv', 30, $arrConditions[0]);
            $operand2 = pack('Cv', 30, $arrConditions[1]);
        } else {
            $szValue1 = 0;
            $szValue2 = 0;
            $operand1 = null;
            $operand2 = null;
        }
        // $flags : Option flags
        // Alignment
        $bAlignHz = $conditional->getStyle()->getAlignment()->getHorizontal() == null ? 1 : 0;
        $bAlignVt = $conditional->getStyle()->getAlignment()->getVertical() == null ? 1 : 0;
        $bAlignWrapTx = $conditional->getStyle()->getAlignment()->getWrapText() == false ? 1 : 0;
        $bTxRotation = $conditional->getStyle()->getAlignment()->getTextRotation() == null ? 1 : 0;
        $bIndent = $conditional->getStyle()->getAlignment()->getIndent() == 0 ? 1 : 0;
        $bShrinkToFit = $conditional->getStyle()->getAlignment()->getShrinkToFit() == false ? 1 : 0;
        if ($bAlignHz == 0 || $bAlignVt == 0 || $bAlignWrapTx == 0 || $bTxRotation == 0 || $bIndent == 0 || $bShrinkToFit == 0) {
            $bFormatAlign = 1;
        } else {
            $bFormatAlign = 0;
        }
        // Protection
        $bProtLocked = $conditional->getStyle()->getProtection()->getLocked() == null ? 1 : 0;
        $bProtHidden = $conditional->getStyle()->getProtection()->getHidden() == null ? 1 : 0;
        if ($bProtLocked == 0 || $bProtHidden == 0) {
            $bFormatProt = 1;
        } else {
            $bFormatProt = 0;
        }
        // Border
        $bBorderLeft = $conditional->getStyle()->getBorders()->getLeft()->getColor()->getARGB() == Color::COLOR_BLACK && $conditional->getStyle()->getBorders()->getLeft()->getBorderStyle() == Border::BORDER_NONE ? 1 : 0;
        $bBorderRight = $conditional->getStyle()->getBorders()->getRight()->getColor()->getARGB() == Color::COLOR_BLACK && $conditional->getStyle()->getBorders()->getRight()->getBorderStyle() == Border::BORDER_NONE ? 1 : 0;
        $bBorderTop = $conditional->getStyle()->getBorders()->getTop()->getColor()->getARGB() == Color::COLOR_BLACK && $conditional->getStyle()->getBorders()->getTop()->getBorderStyle() == Border::BORDER_NONE ? 1 : 0;
        $bBorderBottom = $conditional->getStyle()->getBorders()->getBottom()->getColor()->getARGB() == Color::COLOR_BLACK && $conditional->getStyle()->getBorders()->getBottom()->getBorderStyle() == Border::BORDER_NONE ? 1 : 0;
        if ($bBorderLeft == 0 || $bBorderRight == 0 || $bBorderTop == 0 || $bBorderBottom == 0) {
            $bFormatBorder = 1;
        } else {
            $bFormatBorder = 0;
        }
        // Pattern
        $bFillStyle = $conditional->getStyle()->getFill()->getFillType() == null ? 0 : 1;
        $bFillColor = $conditional->getStyle()->getFill()->getStartColor()->getARGB() == null ? 0 : 1;
        $bFillColorBg = $conditional->getStyle()->getFill()->getEndColor()->getARGB() == null ? 0 : 1;
        if ($bFillStyle == 0 || $bFillColor == 0 || $bFillColorBg == 0) {
            $bFormatFill = 1;
        } else {
            $bFormatFill = 0;
        }
        // Font
        if ($conditional->getStyle()->getFont()->getName() != null || $conditional->getStyle()->getFont()->getSize() != null || $conditional->getStyle()->getFont()->getBold() != null || $conditional->getStyle()->getFont()->getItalic() != null || $conditional->getStyle()->getFont()->getSuperscript() != null || $conditional->getStyle()->getFont()->getSubscript() != null || $conditional->getStyle()->getFont()->getUnderline() != null || $conditional->getStyle()->getFont()->getStrikethrough() != null || $conditional->getStyle()->getFont()->getColor()->getARGB() != null) {
            $bFormatFont = 1;
        } else {
            $bFormatFont = 0;
        }
        // Alignment
        $flags = 0;
        $flags |= 1 == $bAlignHz ? 1 : 0;
        $flags |= 1 == $bAlignVt ? 2 : 0;
        $flags |= 1 == $bAlignWrapTx ? 4 : 0;
        $flags |= 1 == $bTxRotation ? 8 : 0;
        // Justify last line flag
        $flags |= 1 == 1 ? 16 : 0;
        $flags |= 1 == $bIndent ? 32 : 0;
        $flags |= 1 == $bShrinkToFit ? 64 : 0;
        // Default
        $flags |= 1 == 1 ? 128 : 0;
        // Protection
        $flags |= 1 == $bProtLocked ? 256 : 0;
        $flags |= 1 == $bProtHidden ? 512 : 0;
        // Border
        $flags |= 1 == $bBorderLeft ? 1024 : 0;
        $flags |= 1 == $bBorderRight ? 2048 : 0;
        $flags |= 1 == $bBorderTop ? 4096 : 0;
        $flags |= 1 == $bBorderBottom ? 8192 : 0;
        $flags |= 1 == 1 ? 16384 : 0;
        // Top left to Bottom right border
        $flags |= 1 == 1 ? 32768 : 0;
        // Bottom left to Top right border
        // Pattern
        $flags |= 1 == $bFillStyle ? 65536 : 0;
        $flags |= 1 == $bFillColor ? 131072 : 0;
        $flags |= 1 == $bFillColorBg ? 262144 : 0;
        $flags |= 1 == 1 ? 3670016 : 0;
        // Font
        $flags |= 1 == $bFormatFont ? 67108864 : 0;
        // Alignment:
        $flags |= 1 == $bFormatAlign ? 134217728 : 0;
        // Border
        $flags |= 1 == $bFormatBorder ? 268435456 : 0;
        // Pattern
        $flags |= 1 == $bFormatFill ? 536870912 : 0;
        // Protection
        $flags |= 1 == $bFormatProt ? 1073741824 : 0;
        // Text direction
        $flags |= 1 == 0 ? 2147483648.0 : 0;
        // Data Blocks
        if ($bFormatFont == 1) {
            // Font Name
            if ($conditional->getStyle()->getFont()->getName() == null) {
                $dataBlockFont = pack('VVVVVVVV', 0, 0, 0, 0, 0, 0, 0, 0);
                $dataBlockFont .= pack('VVVVVVVV', 0, 0, 0, 0, 0, 0, 0, 0);
            } else {
                $dataBlockFont = StringHelper::UTF8toBIFF8UnicodeLong($conditional->getStyle()->getFont()->getName());
            }
            // Font Size
            if ($conditional->getStyle()->getFont()->getSize() == null) {
                $dataBlockFont .= pack('V', 20 * 11);
            } else {
                $dataBlockFont .= pack('V', 20 * $conditional->getStyle()->getFont()->getSize());
            }
            // Font Options
            $dataBlockFont .= pack('V', 0);
            // Font weight
            if ($conditional->getStyle()->getFont()->getBold() == true) {
                $dataBlockFont .= pack('v', 700);
            } else {
                $dataBlockFont .= pack('v', 400);
            }
            // Escapement type
            if ($conditional->getStyle()->getFont()->getSubscript() == true) {
                $dataBlockFont .= pack('v', 2);
                $fontEscapement = 0;
            } elseif ($conditional->getStyle()->getFont()->getSuperscript() == true) {
                $dataBlockFont .= pack('v', 1);
                $fontEscapement = 0;
            } else {
                $dataBlockFont .= pack('v', 0);
                $fontEscapement = 1;
            }
            // Underline type
            switch ($conditional->getStyle()->getFont()->getUnderline()) {
                case \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_NONE:
                    $dataBlockFont .= pack('C', 0);
                    $fontUnderline = 0;
                    break;
                case \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLE:
                    $dataBlockFont .= pack('C', 2);
                    $fontUnderline = 0;
                    break;
                case \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLEACCOUNTING:
                    $dataBlockFont .= pack('C', 34);
                    $fontUnderline = 0;
                    break;
                case \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLE:
                    $dataBlockFont .= pack('C', 1);
                    $fontUnderline = 0;
                    break;
                case \PhpOffice\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLEACCOUNTING:
                    $dataBlockFont .= pack('C', 33);
                    $fontUnderline = 0;
                    break;
                default:
                    $dataBlockFont .= pack('C', 0);
                    $fontUnderline = 1;
                    break;
            }
            // Not used (3)
            $dataBlockFont .= pack('vC', 0, 0);
            // Font color index
            switch ($conditional->getStyle()->getFont()->getColor()->getRGB()) {
                case '000000':
                    $colorIdx = 8;
                    break;
                case 'FFFFFF':
                    $colorIdx = 9;
                    break;
                case 'FF0000':
                    $colorIdx = 10;
                    break;
                case '00FF00':
                    $colorIdx = 11;
                    break;
                case '0000FF':
                    $colorIdx = 12;
                    break;
                case 'FFFF00':
                    $colorIdx = 13;
                    break;
                case 'FF00FF':
                    $colorIdx = 14;
                    break;
                case '00FFFF':
                    $colorIdx = 15;
                    break;
                case '800000':
                    $colorIdx = 16;
                    break;
                case '008000':
                    $colorIdx = 17;
                    break;
                case '000080':
                    $colorIdx = 18;
                    break;
                case '808000':
                    $colorIdx = 19;
                    break;
                case '800080':
                    $colorIdx = 20;
                    break;
                case '008080':
                    $colorIdx = 21;
                    break;
                case 'C0C0C0':
                    $colorIdx = 22;
                    break;
                case '808080':
                    $colorIdx = 23;
                    break;
                case '9999FF':
                    $colorIdx = 24;
                    break;
                case '993366':
                    $colorIdx = 25;
                    break;
                case 'FFFFCC':
                    $colorIdx = 26;
                    break;
                case 'CCFFFF':
                    $colorIdx = 27;
                    break;
                case '660066':
                    $colorIdx = 28;
                    break;
                case 'FF8080':
                    $colorIdx = 29;
                    break;
                case '0066CC':
                    $colorIdx = 30;
                    break;
                case 'CCCCFF':
                    $colorIdx = 31;
                    break;
                case '000080':
                    $colorIdx = 32;
                    break;
                case 'FF00FF':
                    $colorIdx = 33;
                    break;
                case 'FFFF00':
                    $colorIdx = 34;
                    break;
                case '00FFFF':
                    $colorIdx = 35;
                    break;
                case '800080':
                    $colorIdx = 36;
                    break;
                case '800000':
                    $colorIdx = 37;
                    break;
                case '008080':
                    $colorIdx = 38;
                    break;
                case '0000FF':
                    $colorIdx = 39;
                    break;
                case '00CCFF':
                    $colorIdx = 40;
                    break;
                case 'CCFFFF':
                    $colorIdx = 41;
                    break;
                case 'CCFFCC':
                    $colorIdx = 42;
                    break;
                case 'FFFF99':
                    $colorIdx = 43;
                    break;
                case '99CCFF':
                    $colorIdx = 44;
                    break;
                case 'FF99CC':
                    $colorIdx = 45;
                    break;
                case 'CC99FF':
                    $colorIdx = 46;
                    break;
                case 'FFCC99':
                    $colorIdx = 47;
                    break;
                case '3366FF':
                    $colorIdx = 48;
                    break;
                case '33CCCC':
                    $colorIdx = 49;
                    break;
                case '99CC00':
                    $colorIdx = 50;
                    break;
                case 'FFCC00':
                    $colorIdx = 51;
                    break;
                case 'FF9900':
                    $colorIdx = 52;
                    break;
                case 'FF6600':
                    $colorIdx = 53;
                    break;
                case '666699':
                    $colorIdx = 54;
                    break;
                case '969696':
                    $colorIdx = 55;
                    break;
                case '003366':
                    $colorIdx = 56;
                    break;
                case '339966':
                    $colorIdx = 57;
                    break;
                case '003300':
                    $colorIdx = 58;
                    break;
                case '333300':
                    $colorIdx = 59;
                    break;
                case '993300':
                    $colorIdx = 60;
                    break;
                case '993366':
                    $colorIdx = 61;
                    break;
                case '333399':
                    $colorIdx = 62;
                    break;
                case '333333':
                    $colorIdx = 63;
                    break;
                default:
                    $colorIdx = 0;
                    break;
            }
            $dataBlockFont .= pack('V', $colorIdx);
            // Not used (4)
            $dataBlockFont .= pack('V', 0);
            // Options flags for modified font attributes
            $optionsFlags = 0;
            $optionsFlagsBold = $conditional->getStyle()->getFont()->getBold() == null ? 1 : 0;
            $optionsFlags |= 1 == $optionsFlagsBold ? 2 : 0;
            $optionsFlags |= 1 == 1 ? 8 : 0;
            $optionsFlags |= 1 == 1 ? 16 : 0;
            $optionsFlags |= 1 == 0 ? 32 : 0;
            $optionsFlags |= 1 == 1 ? 128 : 0;
            $dataBlockFont .= pack('V', $optionsFlags);
            // Escapement type
            $dataBlockFont .= pack('V', $fontEscapement);
            // Underline type
            $dataBlockFont .= pack('V', $fontUnderline);
            // Always
            $dataBlockFont .= pack('V', 0);
            // Always
            $dataBlockFont .= pack('V', 0);
            // Not used (8)
            $dataBlockFont .= pack('VV', 0, 0);
            // Always
            $dataBlockFont .= pack('v', 1);
        }
        if ($bFormatAlign == 1) {
            $blockAlign = 0;
            // Alignment and text break
            switch ($conditional->getStyle()->getAlignment()->getHorizontal()) {
                case Alignment::HORIZONTAL_GENERAL:
                    $blockAlign = 0;
                    break;
                case Alignment::HORIZONTAL_LEFT:
                    $blockAlign = 1;
                    break;
                case Alignment::HORIZONTAL_RIGHT:
                    $blockAlign = 3;
                    break;
                case Alignment::HORIZONTAL_CENTER:
                    $blockAlign = 2;
                    break;
                case Alignment::HORIZONTAL_CENTER_CONTINUOUS:
                    $blockAlign = 6;
                    break;
                case Alignment::HORIZONTAL_JUSTIFY:
                    $blockAlign = 5;
                    break;
            }
            if ($conditional->getStyle()->getAlignment()->getWrapText() == true) {
                $blockAlign |= 1 << 3;
            } else {
                $blockAlign |= 0 << 3;
            }
            switch ($conditional->getStyle()->getAlignment()->getVertical()) {
                case Alignment::VERTICAL_BOTTOM:
                    $blockAlign = 2 << 4;
                    break;
                case Alignment::VERTICAL_TOP:
                    $blockAlign = 0 << 4;
                    break;
                case Alignment::VERTICAL_CENTER:
                    $blockAlign = 1 << 4;
                    break;
                case Alignment::VERTICAL_JUSTIFY:
                    $blockAlign = 3 << 4;
                    break;
            }
            $blockAlign |= 0 << 7;
            // Text rotation angle
            $blockRotation = $conditional->getStyle()->getAlignment()->getTextRotation();
            // Indentation
            $blockIndent = $conditional->getStyle()->getAlignment()->getIndent();
            if ($conditional->getStyle()->getAlignment()->getShrinkToFit() == true) {
                $blockIndent |= 1 << 4;
            } else {
                $blockIndent |= 0 << 4;
            }
            $blockIndent |= 0 << 6;
            // Relative indentation
            $blockIndentRelative = 255;
            $dataBlockAlign = pack('CCvvv', $blockAlign, $blockRotation, $blockIndent, $blockIndentRelative, 0);
        }
        if ($bFormatBorder == 1) {
            $blockLineStyle = 0;
            switch ($conditional->getStyle()->getBorders()->getLeft()->getBorderStyle()) {
                case Border::BORDER_NONE:
                    $blockLineStyle |= 0;
                    break;
                case Border::BORDER_THIN:
                    $blockLineStyle |= 1;
                    break;
                case Border::BORDER_MEDIUM:
                    $blockLineStyle |= 2;
                    break;
                case Border::BORDER_DASHED:
                    $blockLineStyle |= 3;
                    break;
                case Border::BORDER_DOTTED:
                    $blockLineStyle |= 4;
                    break;
                case Border::BORDER_THICK:
                    $blockLineStyle |= 5;
                    break;
                case Border::BORDER_DOUBLE:
                    $blockLineStyle |= 6;
                    break;
                case Border::BORDER_HAIR:
                    $blockLineStyle |= 7;
                    break;
                case Border::BORDER_MEDIUMDASHED:
                    $blockLineStyle |= 8;
                    break;
                case Border::BORDER_DASHDOT:
                    $blockLineStyle |= 9;
                    break;
                case Border::BORDER_MEDIUMDASHDOT:
                    $blockLineStyle |= 10;
                    break;
                case Border::BORDER_DASHDOTDOT:
                    $blockLineStyle |= 11;
                    break;
                case Border::BORDER_MEDIUMDASHDOTDOT:
                    $blockLineStyle |= 12;
                    break;
                case Border::BORDER_SLANTDASHDOT:
                    $blockLineStyle |= 13;
                    break;
            }
            switch ($conditional->getStyle()->getBorders()->getRight()->getBorderStyle()) {
                case Border::BORDER_NONE:
                    $blockLineStyle |= 0 << 4;
                    break;
                case Border::BORDER_THIN:
                    $blockLineStyle |= 1 << 4;
                    break;
                case Border::BORDER_MEDIUM:
                    $blockLineStyle |= 2 << 4;
                    break;
                case Border::BORDER_DASHED:
                    $blockLineStyle |= 3 << 4;
                    break;
                case Border::BORDER_DOTTED:
                    $blockLineStyle |= 4 << 4;
                    break;
                case Border::BORDER_THICK:
                    $blockLineStyle |= 5 << 4;
                    break;
                case Border::BORDER_DOUBLE:
                    $blockLineStyle |= 6 << 4;
                    break;
                case Border::BORDER_HAIR:
                    $blockLineStyle |= 7 << 4;
                    break;
                case Border::BORDER_MEDIUMDASHED:
                    $blockLineStyle |= 8 << 4;
                    break;
                case Border::BORDER_DASHDOT:
                    $blockLineStyle |= 9 << 4;
                    break;
                case Border::BORDER_MEDIUMDASHDOT:
                    $blockLineStyle |= 10 << 4;
                    break;
                case Border::BORDER_DASHDOTDOT:
                    $blockLineStyle |= 11 << 4;
                    break;
                case Border::BORDER_MEDIUMDASHDOTDOT:
                    $blockLineStyle |= 12 << 4;
                    break;
                case Border::BORDER_SLANTDASHDOT:
                    $blockLineStyle |= 13 << 4;
                    break;
            }
            switch ($conditional->getStyle()->getBorders()->getTop()->getBorderStyle()) {
                case Border::BORDER_NONE:
                    $blockLineStyle |= 0 << 8;
                    break;
                case Border::BORDER_THIN:
                    $blockLineStyle |= 1 << 8;
                    break;
                case Border::BORDER_MEDIUM:
                    $blockLineStyle |= 2 << 8;
                    break;
                case Border::BORDER_DASHED:
                    $blockLineStyle |= 3 << 8;
                    break;
                case Border::BORDER_DOTTED:
                    $blockLineStyle |= 4 << 8;
                    break;
                case Border::BORDER_THICK:
                    $blockLineStyle |= 5 << 8;
                    break;
                case Border::BORDER_DOUBLE:
                    $blockLineStyle |= 6 << 8;
                    break;
                case Border::BORDER_HAIR:
                    $blockLineStyle |= 7 << 8;
                    break;
                case Border::BORDER_MEDIUMDASHED:
                    $blockLineStyle |= 8 << 8;
                    break;
                case Border::BORDER_DASHDOT:
                    $blockLineStyle |= 9 << 8;
                    break;
                case Border::BORDER_MEDIUMDASHDOT:
                    $blockLineStyle |= 10 << 8;
                    break;
                case Border::BORDER_DASHDOTDOT:
                    $blockLineStyle |= 11 << 8;
                    break;
                case Border::BORDER_MEDIUMDASHDOTDOT:
                    $blockLineStyle |= 12 << 8;
                    break;
                case Border::BORDER_SLANTDASHDOT:
                    $blockLineStyle |= 13 << 8;
                    break;
            }
            switch ($conditional->getStyle()->getBorders()->getBottom()->getBorderStyle()) {
                case Border::BORDER_NONE:
                    $blockLineStyle |= 0 << 12;
                    break;
                case Border::BORDER_THIN:
                    $blockLineStyle |= 1 << 12;
                    break;
                case Border::BORDER_MEDIUM:
                    $blockLineStyle |= 2 << 12;
                    break;
                case Border::BORDER_DASHED:
                    $blockLineStyle |= 3 << 12;
                    break;
                case Border::BORDER_DOTTED:
                    $blockLineStyle |= 4 << 12;
                    break;
                case Border::BORDER_THICK:
                    $blockLineStyle |= 5 << 12;
                    break;
                case Border::BORDER_DOUBLE:
                    $blockLineStyle |= 6 << 12;
                    break;
                case Border::BORDER_HAIR:
                    $blockLineStyle |= 7 << 12;
                    break;
                case Border::BORDER_MEDIUMDASHED:
                    $blockLineStyle |= 8 << 12;
                    break;
                case Border::BORDER_DASHDOT:
                    $blockLineStyle |= 9 << 12;
                    break;
                case Border::BORDER_MEDIUMDASHDOT:
                    $blockLineStyle |= 10 << 12;
                    break;
                case Border::BORDER_DASHDOTDOT:
                    $blockLineStyle |= 11 << 12;
                    break;
                case Border::BORDER_MEDIUMDASHDOTDOT:
                    $blockLineStyle |= 12 << 12;
                    break;
                case Border::BORDER_SLANTDASHDOT:
                    $blockLineStyle |= 13 << 12;
                    break;
            }
            //@todo writeCFRule() => $blockLineStyle => Index Color for left line
            //@todo writeCFRule() => $blockLineStyle => Index Color for right line
            //@todo writeCFRule() => $blockLineStyle => Top-left to bottom-right on/off
            //@todo writeCFRule() => $blockLineStyle => Bottom-left to top-right on/off
            $blockColor = 0;
            //@todo writeCFRule() => $blockColor => Index Color for top line
            //@todo writeCFRule() => $blockColor => Index Color for bottom line
            //@todo writeCFRule() => $blockColor => Index Color for diagonal line
            switch ($conditional->getStyle()->getBorders()->getDiagonal()->getBorderStyle()) {
                case Border::BORDER_NONE:
                    $blockColor |= 0 << 21;
                    break;
                case Border::BORDER_THIN:
                    $blockColor |= 1 << 21;
                    break;
                case Border::BORDER_MEDIUM:
                    $blockColor |= 2 << 21;
                    break;
                case Border::BORDER_DASHED:
                    $blockColor |= 3 << 21;
                    break;
                case Border::BORDER_DOTTED:
                    $blockColor |= 4 << 21;
                    break;
                case Border::BORDER_THICK:
                    $blockColor |= 5 << 21;
                    break;
                case Border::BORDER_DOUBLE:
                    $blockColor |= 6 << 21;
                    break;
                case Border::BORDER_HAIR:
                    $blockColor |= 7 << 21;
                    break;
                case Border::BORDER_MEDIUMDASHED:
                    $blockColor |= 8 << 21;
                    break;
                case Border::BORDER_DASHDOT:
                    $blockColor |= 9 << 21;
                    break;
                case Border::BORDER_MEDIUMDASHDOT:
                    $blockColor |= 10 << 21;
                    break;
                case Border::BORDER_DASHDOTDOT:
                    $blockColor |= 11 << 21;
                    break;
                case Border::BORDER_MEDIUMDASHDOTDOT:
                    $blockColor |= 12 << 21;
                    break;
                case Border::BORDER_SLANTDASHDOT:
                    $blockColor |= 13 << 21;
                    break;
            }
            $dataBlockBorder = pack('vv', $blockLineStyle, $blockColor);
        }
        if ($bFormatFill == 1) {
            // Fill Patern Style
            $blockFillPatternStyle = 0;
            switch ($conditional->getStyle()->getFill()->getFillType()) {
                case Fill::FILL_NONE:
                    $blockFillPatternStyle = 0;
                    break;
                case Fill::FILL_SOLID:
                    $blockFillPatternStyle = 1;
                    break;
                case Fill::FILL_PATTERN_MEDIUMGRAY:
                    $blockFillPatternStyle = 2;
                    break;
                case Fill::FILL_PATTERN_DARKGRAY:
                    $blockFillPatternStyle = 3;
                    break;
                case Fill::FILL_PATTERN_LIGHTGRAY:
                    $blockFillPatternStyle = 4;
                    break;
                case Fill::FILL_PATTERN_DARKHORIZONTAL:
                    $blockFillPatternStyle = 5;
                    break;
                case Fill::FILL_PATTERN_DARKVERTICAL:
                    $blockFillPatternStyle = 6;
                    break;
                case Fill::FILL_PATTERN_DARKDOWN:
                    $blockFillPatternStyle = 7;
                    break;
                case Fill::FILL_PATTERN_DARKUP:
                    $blockFillPatternStyle = 8;
                    break;
                case Fill::FILL_PATTERN_DARKGRID:
                    $blockFillPatternStyle = 9;
                    break;
                case Fill::FILL_PATTERN_DARKTRELLIS:
                    $blockFillPatternStyle = 10;
                    break;
                case Fill::FILL_PATTERN_LIGHTHORIZONTAL:
                    $blockFillPatternStyle = 11;
                    break;
                case Fill::FILL_PATTERN_LIGHTVERTICAL:
                    $blockFillPatternStyle = 12;
                    break;
                case Fill::FILL_PATTERN_LIGHTDOWN:
                    $blockFillPatternStyle = 13;
                    break;
                case Fill::FILL_PATTERN_LIGHTUP:
                    $blockFillPatternStyle = 14;
                    break;
                case Fill::FILL_PATTERN_LIGHTGRID:
                    $blockFillPatternStyle = 15;
                    break;
                case Fill::FILL_PATTERN_LIGHTTRELLIS:
                    $blockFillPatternStyle = 16;
                    break;
                case Fill::FILL_PATTERN_GRAY125:
                    $blockFillPatternStyle = 17;
                    break;
                case Fill::FILL_PATTERN_GRAY0625:
                    $blockFillPatternStyle = 18;
                    break;
                case Fill::FILL_GRADIENT_LINEAR:
                    $blockFillPatternStyle = 0;
                    break;
                // does not exist in BIFF8
                case Fill::FILL_GRADIENT_PATH:
                    $blockFillPatternStyle = 0;
                    break;
                // does not exist in BIFF8
                default:
                    $blockFillPatternStyle = 0;
                    break;
            }
            // Color
            switch ($conditional->getStyle()->getFill()->getStartColor()->getRGB()) {
                case '000000':
                    $colorIdxBg = 8;
                    break;
                case 'FFFFFF':
                    $colorIdxBg = 9;
                    break;
                case 'FF0000':
                    $colorIdxBg = 10;
                    break;
                case '00FF00':
                    $colorIdxBg = 11;
                    break;
                case '0000FF':
                    $colorIdxBg = 12;
                    break;
                case 'FFFF00':
                    $colorIdxBg = 13;
                    break;
                case 'FF00FF':
                    $colorIdxBg = 14;
                    break;
                case '00FFFF':
                    $colorIdxBg = 15;
                    break;
                case '800000':
                    $colorIdxBg = 16;
                    break;
                case '008000':
                    $colorIdxBg = 17;
                    break;
                case '000080':
                    $colorIdxBg = 18;
                    break;
                case '808000':
                    $colorIdxBg = 19;
                    break;
                case '800080':
                    $colorIdxBg = 20;
                    break;
                case '008080':
                    $colorIdxBg = 21;
                    break;
                case 'C0C0C0':
                    $colorIdxBg = 22;
                    break;
                case '808080':
                    $colorIdxBg = 23;
                    break;
                case '9999FF':
                    $colorIdxBg = 24;
                    break;
                case '993366':
                    $colorIdxBg = 25;
                    break;
                case 'FFFFCC':
                    $colorIdxBg = 26;
                    break;
                case 'CCFFFF':
                    $colorIdxBg = 27;
                    break;
                case '660066':
                    $colorIdxBg = 28;
                    break;
                case 'FF8080':
                    $colorIdxBg = 29;
                    break;
                case '0066CC':
                    $colorIdxBg = 30;
                    break;
                case 'CCCCFF':
                    $colorIdxBg = 31;
                    break;
                case '000080':
                    $colorIdxBg = 32;
                    break;
                case 'FF00FF':
                    $colorIdxBg = 33;
                    break;
                case 'FFFF00':
                    $colorIdxBg = 34;
                    break;
                case '00FFFF':
                    $colorIdxBg = 35;
                    break;
                case '800080':
                    $colorIdxBg = 36;
                    break;
                case '800000':
                    $colorIdxBg = 37;
                    break;
                case '008080':
                    $colorIdxBg = 38;
                    break;
                case '0000FF':
                    $colorIdxBg = 39;
                    break;
                case '00CCFF':
                    $colorIdxBg = 40;
                    break;
                case 'CCFFFF':
                    $colorIdxBg = 41;
                    break;
                case 'CCFFCC':
                    $colorIdxBg = 42;
                    break;
                case 'FFFF99':
                    $colorIdxBg = 43;
                    break;
                case '99CCFF':
                    $colorIdxBg = 44;
                    break;
                case 'FF99CC':
                    $colorIdxBg = 45;
                    break;
                case 'CC99FF':
                    $colorIdxBg = 46;
                    break;
                case 'FFCC99':
                    $colorIdxBg = 47;
                    break;
                case '3366FF':
                    $colorIdxBg = 48;
                    break;
                case '33CCCC':
                    $colorIdxBg = 49;
                    break;
                case '99CC00':
                    $colorIdxBg = 50;
                    break;
                case 'FFCC00':
                    $colorIdxBg = 51;
                    break;
                case 'FF9900':
                    $colorIdxBg = 52;
                    break;
                case 'FF6600':
                    $colorIdxBg = 53;
                    break;
                case '666699':
                    $colorIdxBg = 54;
                    break;
                case '969696':
                    $colorIdxBg = 55;
                    break;
                case '003366':
                    $colorIdxBg = 56;
                    break;
                case '339966':
                    $colorIdxBg = 57;
                    break;
                case '003300':
                    $colorIdxBg = 58;
                    break;
                case '333300':
                    $colorIdxBg = 59;
                    break;
                case '993300':
                    $colorIdxBg = 60;
                    break;
                case '993366':
                    $colorIdxBg = 61;
                    break;
                case '333399':
                    $colorIdxBg = 62;
                    break;
                case '333333':
                    $colorIdxBg = 63;
                    break;
                default:
                    $colorIdxBg = 65;
                    break;
            }
            // Fg Color
            switch ($conditional->getStyle()->getFill()->getEndColor()->getRGB()) {
                case '000000':
                    $colorIdxFg = 8;
                    break;
                case 'FFFFFF':
                    $colorIdxFg = 9;
                    break;
                case 'FF0000':
                    $colorIdxFg = 10;
                    break;
                case '00FF00':
                    $colorIdxFg = 11;
                    break;
                case '0000FF':
                    $colorIdxFg = 12;
                    break;
                case 'FFFF00':
                    $colorIdxFg = 13;
                    break;
                case 'FF00FF':
                    $colorIdxFg = 14;
                    break;
                case '00FFFF':
                    $colorIdxFg = 15;
                    break;
                case '800000':
                    $colorIdxFg = 16;
                    break;
                case '008000':
                    $colorIdxFg = 17;
                    break;
                case '000080':
                    $colorIdxFg = 18;
                    break;
                case '808000':
                    $colorIdxFg = 19;
                    break;
                case '800080':
                    $colorIdxFg = 20;
                    break;
                case '008080':
                    $colorIdxFg = 21;
                    break;
                case 'C0C0C0':
                    $colorIdxFg = 22;
                    break;
                case '808080':
                    $colorIdxFg = 23;
                    break;
                case '9999FF':
                    $colorIdxFg = 24;
                    break;
                case '993366':
                    $colorIdxFg = 25;
                    break;
                case 'FFFFCC':
                    $colorIdxFg = 26;
                    break;
                case 'CCFFFF':
                    $colorIdxFg = 27;
                    break;
                case '660066':
                    $colorIdxFg = 28;
                    break;
                case 'FF8080':
                    $colorIdxFg = 29;
                    break;
                case '0066CC':
                    $colorIdxFg = 30;
                    break;
                case 'CCCCFF':
                    $colorIdxFg = 31;
                    break;
                case '000080':
                    $colorIdxFg = 32;
                    break;
                case 'FF00FF':
                    $colorIdxFg = 33;
                    break;
                case 'FFFF00':
                    $colorIdxFg = 34;
                    break;
                case '00FFFF':
                    $colorIdxFg = 35;
                    break;
                case '800080':
                    $colorIdxFg = 36;
                    break;
                case '800000':
                    $colorIdxFg = 37;
                    break;
                case '008080':
                    $colorIdxFg = 38;
                    break;
                case '0000FF':
                    $colorIdxFg = 39;
                    break;
                case '00CCFF':
                    $colorIdxFg = 40;
                    break;
                case 'CCFFFF':
                    $colorIdxFg = 41;
                    break;
                case 'CCFFCC':
                    $colorIdxFg = 42;
                    break;
                case 'FFFF99':
                    $colorIdxFg = 43;
                    break;
                case '99CCFF':
                    $colorIdxFg = 44;
                    break;
                case 'FF99CC':
                    $colorIdxFg = 45;
                    break;
                case 'CC99FF':
                    $colorIdxFg = 46;
                    break;
                case 'FFCC99':
                    $colorIdxFg = 47;
                    break;
                case '3366FF':
                    $colorIdxFg = 48;
                    break;
                case '33CCCC':
                    $colorIdxFg = 49;
                    break;
                case '99CC00':
                    $colorIdxFg = 50;
                    break;
                case 'FFCC00':
                    $colorIdxFg = 51;
                    break;
                case 'FF9900':
                    $colorIdxFg = 52;
                    break;
                case 'FF6600':
                    $colorIdxFg = 53;
                    break;
                case '666699':
                    $colorIdxFg = 54;
                    break;
                case '969696':
                    $colorIdxFg = 55;
                    break;
                case '003366':
                    $colorIdxFg = 56;
                    break;
                case '339966':
                    $colorIdxFg = 57;
                    break;
                case '003300':
                    $colorIdxFg = 58;
                    break;
                case '333300':
                    $colorIdxFg = 59;
                    break;
                case '993300':
                    $colorIdxFg = 60;
                    break;
                case '993366':
                    $colorIdxFg = 61;
                    break;
                case '333399':
                    $colorIdxFg = 62;
                    break;
                case '333333':
                    $colorIdxFg = 63;
                    break;
                default:
                    $colorIdxFg = 64;
                    break;
            }
            $dataBlockFill = pack('v', $blockFillPatternStyle);
            $dataBlockFill .= pack('v', $colorIdxFg | $colorIdxBg << 7);
        }
        if ($bFormatProt == 1) {
            $dataBlockProtection = 0;
            if ($conditional->getStyle()->getProtection()->getLocked() == Protection::PROTECTION_PROTECTED) {
                $dataBlockProtection = 1;
            }
            if ($conditional->getStyle()->getProtection()->getHidden() == Protection::PROTECTION_PROTECTED) {
                $dataBlockProtection = 1 << 1;
            }
        }
        $data = pack('CCvvVv', $type, $operatorType, $szValue1, $szValue2, $flags, 0);
        if ($bFormatFont == 1) {
            // Block Formatting : OK
            $data .= $dataBlockFont;
        }
        if ($bFormatAlign == 1) {
            $data .= $dataBlockAlign;
        }
        if ($bFormatBorder == 1) {
            $data .= $dataBlockBorder;
        }
        if ($bFormatFill == 1) {
            // Block Formatting : OK
            $data .= $dataBlockFill;
        }
        if ($bFormatProt == 1) {
            $data .= $dataBlockProtection;
        }
        if ($operand1 !== null) {
            $data .= $operand1;
        }
        if ($operand2 !== null) {
            $data .= $operand2;
        }
        $header = pack('vv', $record, strlen($data));
        $this->append($header . $data);
    }
    /**
     * Write CFHeader record.
     */
    private function writeCFHeader()
    {
        $record = 432;
        // Record identifier
        $length = 22;
        // Bytes to follow
        $numColumnMin = null;
        $numColumnMax = null;
        $numRowMin = null;
        $numRowMax = null;
        $arrConditional = array();
        foreach ($this->phpSheet->getConditionalStylesCollection() as $cellCoordinate => $conditionalStyles) {
            foreach ($conditionalStyles as $conditional) {
                if ($conditional->getConditionType() == Conditional::CONDITION_EXPRESSION || $conditional->getConditionType() == Conditional::CONDITION_CELLIS) {
                    if (!in_array($conditional->getHashCode(), $arrConditional)) {
                        $arrConditional[] = $conditional->getHashCode();
                    }
                    // Cells
                    $arrCoord = Coordinate::coordinateFromString($cellCoordinate);
                    if (!is_numeric($arrCoord[0])) {
                        $arrCoord[0] = Coordinate::columnIndexFromString($arrCoord[0]);
                    }
                    if ($numColumnMin === null || $numColumnMin > $arrCoord[0]) {
                        $numColumnMin = $arrCoord[0];
                    }
                    if ($numColumnMax === null || $numColumnMax < $arrCoord[0]) {
                        $numColumnMax = $arrCoord[0];
                    }
                    if ($numRowMin === null || $numRowMin > $arrCoord[1]) {
                        $numRowMin = $arrCoord[1];
                    }
                    if ($numRowMax === null || $numRowMax < $arrCoord[1]) {
                        $numRowMax = $arrCoord[1];
                    }
                }
            }
        }
        $needRedraw = 1;
        $cellRange = pack('vvvv', $numRowMin - 1, $numRowMax - 1, $numColumnMin - 1, $numColumnMax - 1);
        $header = pack('vv', $record, $length);
        $data = pack('vv', count($arrConditional), $needRedraw);
        $data .= $cellRange;
        $data .= pack('v', 1);
        $data .= $cellRange;
        $this->append($header . $data);
    }
}