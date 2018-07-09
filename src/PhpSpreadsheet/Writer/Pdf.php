<?php

namespace PhpOffice\PhpSpreadsheet\Writer;

use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
use PhpOffice\PhpSpreadsheet\Shared\File;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Writer\Exception as WriterException;
abstract class Pdf extends Html
{
    /**
     * Temporary storage directory.
     *
     * @var string
     */
    protected $tempDir = '';
    /**
     * Font.
     *
     * @var string
     */
    protected $font = 'freesans';
    /**
     * Orientation (Over-ride).
     *
     * @var string
     */
    protected $orientation;
    /**
     * Paper size (Over-ride).
     *
     * @var int
     */
    protected $paperSize;
    /**
     * Temporary storage for Save Array Return type.
     *
     * @var string
     */
    private $saveArrayReturnType;
    /**
     * Paper Sizes xRef List.
     *
     * @var array
     */
    protected static $paperSizes = array(PageSetup::PAPERSIZE_LETTER => 'LETTER', PageSetup::PAPERSIZE_LETTER_SMALL => 'LETTER', PageSetup::PAPERSIZE_TABLOID => array(792.0, 1224.0), PageSetup::PAPERSIZE_LEDGER => array(1224.0, 792.0), PageSetup::PAPERSIZE_LEGAL => 'LEGAL', PageSetup::PAPERSIZE_STATEMENT => array(396.0, 612.0), PageSetup::PAPERSIZE_EXECUTIVE => 'EXECUTIVE', PageSetup::PAPERSIZE_A3 => 'A3', PageSetup::PAPERSIZE_A4 => 'A4', PageSetup::PAPERSIZE_A4_SMALL => 'A4', PageSetup::PAPERSIZE_A5 => 'A5', PageSetup::PAPERSIZE_B4 => 'B4', PageSetup::PAPERSIZE_B5 => 'B5', PageSetup::PAPERSIZE_FOLIO => 'FOLIO', PageSetup::PAPERSIZE_QUARTO => array(609.45, 779.53), PageSetup::PAPERSIZE_STANDARD_1 => array(720.0, 1008.0), PageSetup::PAPERSIZE_STANDARD_2 => array(792.0, 1224.0), PageSetup::PAPERSIZE_NOTE => 'LETTER', PageSetup::PAPERSIZE_NO9_ENVELOPE => array(279.0, 639.0), PageSetup::PAPERSIZE_NO10_ENVELOPE => array(297.0, 684.0), PageSetup::PAPERSIZE_NO11_ENVELOPE => array(324.0, 747.0), PageSetup::PAPERSIZE_NO12_ENVELOPE => array(342.0, 792.0), PageSetup::PAPERSIZE_NO14_ENVELOPE => array(360.0, 828.0), PageSetup::PAPERSIZE_C => array(1224.0, 1584.0), PageSetup::PAPERSIZE_D => array(1584.0, 2448.0), PageSetup::PAPERSIZE_E => array(2448.0, 3168.0), PageSetup::PAPERSIZE_DL_ENVELOPE => array(311.81, 623.62), PageSetup::PAPERSIZE_C5_ENVELOPE => 'C5', PageSetup::PAPERSIZE_C3_ENVELOPE => 'C3', PageSetup::PAPERSIZE_C4_ENVELOPE => 'C4', PageSetup::PAPERSIZE_C6_ENVELOPE => 'C6', PageSetup::PAPERSIZE_C65_ENVELOPE => array(323.15, 649.13), PageSetup::PAPERSIZE_B4_ENVELOPE => 'B4', PageSetup::PAPERSIZE_B5_ENVELOPE => 'B5', PageSetup::PAPERSIZE_B6_ENVELOPE => array(498.9, 354.33), PageSetup::PAPERSIZE_ITALY_ENVELOPE => array(311.81, 651.97), PageSetup::PAPERSIZE_MONARCH_ENVELOPE => array(279.0, 540.0), PageSetup::PAPERSIZE_6_3_4_ENVELOPE => array(261.0, 468.0), PageSetup::PAPERSIZE_US_STANDARD_FANFOLD => array(1071.0, 792.0), PageSetup::PAPERSIZE_GERMAN_STANDARD_FANFOLD => array(612.0, 864.0), PageSetup::PAPERSIZE_GERMAN_LEGAL_FANFOLD => 'FOLIO', PageSetup::PAPERSIZE_ISO_B4 => 'B4', PageSetup::PAPERSIZE_JAPANESE_DOUBLE_POSTCARD => array(566.9299999999999, 419.53), PageSetup::PAPERSIZE_STANDARD_PAPER_1 => array(648.0, 792.0), PageSetup::PAPERSIZE_STANDARD_PAPER_2 => array(720.0, 792.0), PageSetup::PAPERSIZE_STANDARD_PAPER_3 => array(1080.0, 792.0), PageSetup::PAPERSIZE_INVITE_ENVELOPE => array(623.62, 623.62), PageSetup::PAPERSIZE_LETTER_EXTRA_PAPER => array(667.8, 864.0), PageSetup::PAPERSIZE_LEGAL_EXTRA_PAPER => array(667.8, 1080.0), PageSetup::PAPERSIZE_TABLOID_EXTRA_PAPER => array(841.6799999999999, 1296.0), PageSetup::PAPERSIZE_A4_EXTRA_PAPER => array(668.98, 912.76), PageSetup::PAPERSIZE_LETTER_TRANSVERSE_PAPER => array(595.8, 792.0), PageSetup::PAPERSIZE_A4_TRANSVERSE_PAPER => 'A4', PageSetup::PAPERSIZE_LETTER_EXTRA_TRANSVERSE_PAPER => array(667.8, 864.0), PageSetup::PAPERSIZE_SUPERA_SUPERA_A4_PAPER => array(643.46, 1009.13), PageSetup::PAPERSIZE_SUPERB_SUPERB_A3_PAPER => array(864.5700000000001, 1380.47), PageSetup::PAPERSIZE_LETTER_PLUS_PAPER => array(612.0, 913.6799999999999), PageSetup::PAPERSIZE_A4_PLUS_PAPER => array(595.28, 935.4299999999999), PageSetup::PAPERSIZE_A5_TRANSVERSE_PAPER => 'A5', PageSetup::PAPERSIZE_JIS_B5_TRANSVERSE_PAPER => array(515.91, 728.5), PageSetup::PAPERSIZE_A3_EXTRA_PAPER => array(912.76, 1261.42), PageSetup::PAPERSIZE_A5_EXTRA_PAPER => array(493.23, 666.14), PageSetup::PAPERSIZE_ISO_B5_EXTRA_PAPER => array(569.76, 782.36), PageSetup::PAPERSIZE_A2_PAPER => 'A2', PageSetup::PAPERSIZE_A3_TRANSVERSE_PAPER => 'A3', PageSetup::PAPERSIZE_A3_EXTRA_TRANSVERSE_PAPER => array(912.76, 1261.42));
    /**
     * Create a new PDF Writer instance.
     *
     * @param Spreadsheet $spreadsheet Spreadsheet object
     */
    public function __construct(Spreadsheet $spreadsheet)
    {
        parent::__construct($spreadsheet);
        $this->setUseInlineCss(true);
        $this->tempDir = File::sysGetTempDir();
    }
    /**
     * Get Font.
     *
     * @return string
     */
    public function getFont()
    {
        return $this->font;
    }
    /**
     * Set font. Examples:
     *      'arialunicid0-chinese-simplified'
     *      'arialunicid0-chinese-traditional'
     *      'arialunicid0-korean'
     *      'arialunicid0-japanese'.
     *
     * @param string $fontName
     *
     * @return Pdf
     */
    public function setFont($fontName)
    {
        $this->font = $fontName;
        return $this;
    }
    /**
     * Get Paper Size.
     *
     * @return int
     */
    public function getPaperSize()
    {
        return $this->paperSize;
    }
    /**
     * Set Paper Size.
     *
     * @param string $pValue Paper size see PageSetup::PAPERSIZE_*
     *
     * @return self
     */
    public function setPaperSize($pValue)
    {
        $this->paperSize = $pValue;
        return $this;
    }
    /**
     * Get Orientation.
     *
     * @return string
     */
    public function getOrientation()
    {
        return $this->orientation;
    }
    /**
     * Set Orientation.
     *
     * @param string $pValue Page orientation see PageSetup::ORIENTATION_*
     *
     * @return self
     */
    public function setOrientation($pValue)
    {
        $this->orientation = $pValue;
        return $this;
    }
    /**
     * Get temporary storage directory.
     *
     * @return string
     */
    public function getTempDir()
    {
        return $this->tempDir;
    }
    /**
     * Set temporary storage directory.
     *
     * @param string $pValue Temporary storage directory
     *
     * @throws WriterException when directory does not exist
     *
     * @return self
     */
    public function setTempDir($pValue)
    {
        if (is_dir($pValue)) {
            $this->tempDir = $pValue;
        } else {
            throw new WriterException("Directory does not exist: {$pValue}");
        }
        return $this;
    }
    /**
     * Save Spreadsheet to PDF file, pre-save.
     *
     * @param string $pFilename Name of the file to save as
     *
     * @throws WriterException
     *
     * @return resource
     */
    protected function prepareForSave($pFilename)
    {
        //  garbage collect
        $this->spreadsheet->garbageCollect();
        $this->saveArrayReturnType = Calculation::getArrayReturnType();
        Calculation::setArrayReturnType(Calculation::RETURN_ARRAY_AS_VALUE);
        //  Open file
        $fileHandle = fopen($pFilename, 'w');
        if ($fileHandle === false) {
            throw new WriterException("Could not open file {$pFilename} for writing.");
        }
        //  Set PDF
        $this->isPdf = true;
        //  Build CSS
        $this->buildCSS(true);
        return $fileHandle;
    }
    /**
     * Save PhpSpreadsheet to PDF file, post-save.
     *
     * @param resource $fileHandle
     */
    protected function restoreStateAfterSave($fileHandle)
    {
        //  Close file
        fclose($fileHandle);
        Calculation::setArrayReturnType($this->saveArrayReturnType);
    }
}