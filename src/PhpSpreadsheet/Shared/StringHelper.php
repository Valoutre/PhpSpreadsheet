<?php

namespace PhpOffice\PhpSpreadsheet\Shared;

use PhpOffice\PhpSpreadsheet\Calculation\Calculation;
class StringHelper
{
    /**    Constants                */
    /**    Regular Expressions        */
    //    Fraction
    const STRING_REGEXP_FRACTION = '(-?)(\\d+)\\s+(\\d+\\/\\d+)';
    /**
     * Control characters array.
     *
     * @var string[]
     */
    private static $controlCharacters = array();
    /**
     * SYLK Characters array.
     *
     * @var array
     */
    private static $SYLKCharacters = array();
    /**
     * Decimal separator.
     *
     * @var string
     */
    private static $decimalSeparator;
    /**
     * Thousands separator.
     *
     * @var string
     */
    private static $thousandsSeparator;
    /**
     * Currency code.
     *
     * @var string
     */
    private static $currencyCode;
    /**
     * Is iconv extension avalable?
     *
     * @var bool
     */
    private static $isIconvEnabled;
    /**
     * Build control characters array.
     */
    private static function buildControlCharacters()
    {
        for ($i = 0; $i <= 31; ++$i) {
            if ($i != 9 && $i != 10 && $i != 13) {
                $find = '_x' . sprintf('%04s', strtoupper(dechex($i))) . '_';
                $replace = chr($i);
                self::$controlCharacters[$find] = $replace;
            }
        }
    }
    /**
     * Build SYLK characters array.
     */
    private static function buildSYLKCharacters()
    {
        self::$SYLKCharacters = array(' 0' => chr(0), ' 1' => chr(1), ' 2' => chr(2), ' 3' => chr(3), ' 4' => chr(4), ' 5' => chr(5), ' 6' => chr(6), ' 7' => chr(7), ' 8' => chr(8), ' 9' => chr(9), ' :' => chr(10), ' ;' => chr(11), ' <' => chr(12), ' =' => chr(13), ' >' => chr(14), ' ?' => chr(15), '!0' => chr(16), '!1' => chr(17), '!2' => chr(18), '!3' => chr(19), '!4' => chr(20), '!5' => chr(21), '!6' => chr(22), '!7' => chr(23), '!8' => chr(24), '!9' => chr(25), '!:' => chr(26), '!;' => chr(27), '!<' => chr(28), '!=' => chr(29), '!>' => chr(30), '!?' => chr(31), '\'?' => chr(127), '(0' => '€', '(2' => '‚', '(3' => 'ƒ', '(4' => '„', '(5' => '…', '(6' => '†', '(7' => '‡', '(8' => 'ˆ', '(9' => '‰', '(:' => 'Š', '(;' => '‹', 'Nj' => 'Œ', '(>' => 'Ž', ')1' => '‘', ')2' => '’', ')3' => '“', ')4' => '”', ')5' => '•', ')6' => '–', ')7' => '—', ')8' => '˜', ')9' => '™', '):' => 'š', ');' => '›', 'Nz' => 'œ', ')>' => 'ž', ')?' => 'Ÿ', '*0' => ' ', 'N!' => '¡', 'N"' => '¢', 'N#' => '£', 'N(' => '¤', 'N%' => '¥', '*6' => '¦', 'N\'' => '§', 'NH ' => '¨', 'NS' => '©', 'Nc' => 'ª', 'N+' => '«', '*<' => '¬', '*=' => '­', 'NR' => '®', '*?' => '¯', 'N0' => '°', 'N1' => '±', 'N2' => '²', 'N3' => '³', 'NB ' => '´', 'N5' => 'µ', 'N6' => '¶', 'N7' => '·', '+8' => '¸', 'NQ' => '¹', 'Nk' => 'º', 'N;' => '»', 'N<' => '¼', 'N=' => '½', 'N>' => '¾', 'N?' => '¿', 'NAA' => 'À', 'NBA' => 'Á', 'NCA' => 'Â', 'NDA' => 'Ã', 'NHA' => 'Ä', 'NJA' => 'Å', 'Na' => 'Æ', 'NKC' => 'Ç', 'NAE' => 'È', 'NBE' => 'É', 'NCE' => 'Ê', 'NHE' => 'Ë', 'NAI' => 'Ì', 'NBI' => 'Í', 'NCI' => 'Î', 'NHI' => 'Ï', 'Nb' => 'Ð', 'NDN' => 'Ñ', 'NAO' => 'Ò', 'NBO' => 'Ó', 'NCO' => 'Ô', 'NDO' => 'Õ', 'NHO' => 'Ö', '-7' => '×', 'Ni' => 'Ø', 'NAU' => 'Ù', 'NBU' => 'Ú', 'NCU' => 'Û', 'NHU' => 'Ü', '-=' => 'Ý', 'Nl' => 'Þ', 'N{' => 'ß', 'NAa' => 'à', 'NBa' => 'á', 'NCa' => 'â', 'NDa' => 'ã', 'NHa' => 'ä', 'NJa' => 'å', 'Nq' => 'æ', 'NKc' => 'ç', 'NAe' => 'è', 'NBe' => 'é', 'NCe' => 'ê', 'NHe' => 'ë', 'NAi' => 'ì', 'NBi' => 'í', 'NCi' => 'î', 'NHi' => 'ï', 'Ns' => 'ð', 'NDn' => 'ñ', 'NAo' => 'ò', 'NBo' => 'ó', 'NCo' => 'ô', 'NDo' => 'õ', 'NHo' => 'ö', '/7' => '÷', 'Ny' => 'ø', 'NAu' => 'ù', 'NBu' => 'ú', 'NCu' => 'û', 'NHu' => 'ü', '/=' => 'ý', 'N|' => 'þ', 'NHy' => 'ÿ');
    }
    /**
     * Get whether iconv extension is available.
     *
     * @return bool
     */
    public static function getIsIconvEnabled()
    {
        if (isset(self::$isIconvEnabled)) {
            return self::$isIconvEnabled;
        }
        // Fail if iconv doesn't exist
        if (!function_exists('iconv')) {
            self::$isIconvEnabled = false;
            return false;
        }
        // Sometimes iconv is not working, and e.g. iconv('UTF-8', 'UTF-16LE', 'x') just returns false,
        if (!@iconv('UTF-8', 'UTF-16LE', 'x')) {
            self::$isIconvEnabled = false;
            return false;
        }
        // Sometimes iconv_substr('A', 0, 1, 'UTF-8') just returns false in PHP 5.2.0
        // we cannot use iconv in that case either (http://bugs.php.net/bug.php?id=37773)
        if (!@iconv_substr('A', 0, 1, 'UTF-8')) {
            self::$isIconvEnabled = false;
            return false;
        }
        // CUSTOM: IBM AIX iconv() does not work
        if (defined('PHP_OS') && @stristr(PHP_OS, 'AIX') && defined('ICONV_IMPL') && @strcasecmp(ICONV_IMPL, 'unknown') == 0 && defined('ICONV_VERSION') && @strcasecmp(ICONV_VERSION, 'unknown') == 0) {
            self::$isIconvEnabled = false;
            return false;
        }
        // If we reach here no problems were detected with iconv
        self::$isIconvEnabled = true;
        return true;
    }
    private static function buildCharacterSets()
    {
        if (empty(self::$controlCharacters)) {
            self::buildControlCharacters();
        }
        if (empty(self::$SYLKCharacters)) {
            self::buildSYLKCharacters();
        }
    }
    /**
     * Convert from OpenXML escaped control character to PHP control character.
     *
     * Excel 2007 team:
     * ----------------
     * That's correct, control characters are stored directly in the shared-strings table.
     * We do encode characters that cannot be represented in XML using the following escape sequence:
     * _xHHHH_ where H represents a hexadecimal character in the character's value...
     * So you could end up with something like _x0008_ in a string (either in a cell value (<v>)
     * element or in the shared string <t> element.
     *
     * @param string $value Value to unescape
     *
     * @return string
     */
    public static function controlCharacterOOXML2PHP($value)
    {
        self::buildCharacterSets();
        return str_replace(array_keys(self::$controlCharacters), array_values(self::$controlCharacters), $value);
    }
    /**
     * Convert from PHP control character to OpenXML escaped control character.
     *
     * Excel 2007 team:
     * ----------------
     * That's correct, control characters are stored directly in the shared-strings table.
     * We do encode characters that cannot be represented in XML using the following escape sequence:
     * _xHHHH_ where H represents a hexadecimal character in the character's value...
     * So you could end up with something like _x0008_ in a string (either in a cell value (<v>)
     * element or in the shared string <t> element.
     *
     * @param string $value Value to escape
     *
     * @return string
     */
    public static function controlCharacterPHP2OOXML($value)
    {
        self::buildCharacterSets();
        return str_replace(array_values(self::$controlCharacters), array_keys(self::$controlCharacters), $value);
    }
    /**
     * Try to sanitize UTF8, stripping invalid byte sequences. Not perfect. Does not surrogate characters.
     *
     * @param string $value
     *
     * @return string
     */
    public static function sanitizeUTF8($value)
    {
        if (self::getIsIconvEnabled()) {
            $value = @iconv('UTF-8', 'UTF-8', $value);
            return $value;
        }
        $value = mb_convert_encoding($value, 'UTF-8', 'UTF-8');
        return $value;
    }
    /**
     * Check if a string contains UTF8 data.
     *
     * @param string $value
     *
     * @return bool
     */
    public static function isUTF8($value)
    {
        return $value === '' || preg_match('/^./su', $value) === 1;
    }
    /**
     * Formats a numeric value as a string for output in various output writers forcing
     * point as decimal separator in case locale is other than English.
     *
     * @param mixed $value
     *
     * @return string
     */
    public static function formatNumber($value)
    {
        if (is_float($value)) {
            return str_replace(',', '.', $value);
        }
        return (string) $value;
    }
    /**
     * Converts a UTF-8 string into BIFF8 Unicode string data (8-bit string length)
     * Writes the string using uncompressed notation, no rich text, no Asian phonetics
     * If mbstring extension is not available, ASCII is assumed, and compressed notation is used
     * although this will give wrong results for non-ASCII strings
     * see OpenOffice.org's Documentation of the Microsoft Excel File Format, sect. 2.5.3.
     *
     * @param string $value UTF-8 encoded string
     * @param mixed[] $arrcRuns Details of rich text runs in $value
     *
     * @return string
     */
    public static function UTF8toBIFF8UnicodeShort($value, $arrcRuns = array())
    {
        // character count
        $ln = self::countCharacters($value, 'UTF-8');
        // option flags
        if (empty($arrcRuns)) {
            $data = pack('CC', $ln, 1);
            // characters
            $data .= self::convertEncoding($value, 'UTF-16LE', 'UTF-8');
        } else {
            $data = pack('vC', $ln, 9);
            $data .= pack('v', count($arrcRuns));
            // characters
            $data .= self::convertEncoding($value, 'UTF-16LE', 'UTF-8');
            foreach ($arrcRuns as $cRun) {
                $data .= pack('v', $cRun['strlen']);
                $data .= pack('v', $cRun['fontidx']);
            }
        }
        return $data;
    }
    /**
     * Converts a UTF-8 string into BIFF8 Unicode string data (16-bit string length)
     * Writes the string using uncompressed notation, no rich text, no Asian phonetics
     * If mbstring extension is not available, ASCII is assumed, and compressed notation is used
     * although this will give wrong results for non-ASCII strings
     * see OpenOffice.org's Documentation of the Microsoft Excel File Format, sect. 2.5.3.
     *
     * @param string $value UTF-8 encoded string
     *
     * @return string
     */
    public static function UTF8toBIFF8UnicodeLong($value)
    {
        // character count
        $ln = self::countCharacters($value, 'UTF-8');
        // characters
        $chars = self::convertEncoding($value, 'UTF-16LE', 'UTF-8');
        $data = pack('vC', $ln, 1) . $chars;
        return $data;
    }
    /**
     * Convert string from one encoding to another.
     *
     * @param string $value
     * @param string $to Encoding to convert to, e.g. 'UTF-8'
     * @param string $from Encoding to convert from, e.g. 'UTF-16LE'
     *
     * @return string
     */
    public static function convertEncoding($value, $to, $from)
    {
        if (self::getIsIconvEnabled()) {
            $result = iconv($from, $to . '//IGNORE//TRANSLIT', $value);
            if (false !== $result) {
                return $result;
            }
        }
        return mb_convert_encoding($value, $to, $from);
    }
    /**
     * Get character count.
     *
     * @param string $value
     * @param string $enc Encoding
     *
     * @return int Character count
     */
    public static function countCharacters($value, $enc = 'UTF-8')
    {
        return mb_strlen($value, $enc);
    }
    /**
     * Get a substring of a UTF-8 encoded string.
     *
     * @param string $pValue UTF-8 encoded string
     * @param int $pStart Start offset
     * @param int $pLength Maximum number of characters in substring
     *
     * @return string
     */
    public static function substring($pValue, $pStart, $pLength = 0)
    {
        return mb_substr($pValue, $pStart, $pLength, 'UTF-8');
    }
    /**
     * Convert a UTF-8 encoded string to upper case.
     *
     * @param string $pValue UTF-8 encoded string
     *
     * @return string
     */
    public static function strToUpper($pValue)
    {
        return mb_convert_case($pValue, MB_CASE_UPPER, 'UTF-8');
    }
    /**
     * Convert a UTF-8 encoded string to lower case.
     *
     * @param string $pValue UTF-8 encoded string
     *
     * @return string
     */
    public static function strToLower($pValue)
    {
        return mb_convert_case($pValue, MB_CASE_LOWER, 'UTF-8');
    }
    /**
     * Convert a UTF-8 encoded string to title/proper case
     * (uppercase every first character in each word, lower case all other characters).
     *
     * @param string $pValue UTF-8 encoded string
     *
     * @return string
     */
    public static function strToTitle($pValue)
    {
        return mb_convert_case($pValue, MB_CASE_TITLE, 'UTF-8');
    }
    public static function mbIsUpper($char)
    {
        return mb_strtolower($char, 'UTF-8') != $char;
    }
    public static function mbStrSplit($string)
    {
        // Split at all position not after the start: ^
        // and not before the end: $
        return preg_split('/(?<!^)(?!$)/u', $string);
    }
    /**
     * Reverse the case of a string, so that all uppercase characters become lowercase
     * and all lowercase characters become uppercase.
     *
     * @param string $pValue UTF-8 encoded string
     *
     * @return string
     */
    public static function strCaseReverse($pValue)
    {
        $characters = self::mbStrSplit($pValue);
        foreach ($characters as &$character) {
            if (self::mbIsUpper($character)) {
                $character = mb_strtolower($character, 'UTF-8');
            } else {
                $character = mb_strtoupper($character, 'UTF-8');
            }
        }
        return implode('', $characters);
    }
    /**
     * Identify whether a string contains a fractional numeric value,
     * and convert it to a numeric if it is.
     *
     * @param string &$operand string value to test
     *
     * @return bool
     */
    public static function convertToNumberIfFraction(&$operand)
    {
        if (preg_match('/^' . self::STRING_REGEXP_FRACTION . '$/i', $operand, $match)) {
            $sign = $match[1] == '-' ? '-' : '+';
            $fractionFormula = '=' . $sign . $match[2] . $sign . $match[3];
            $operand = Calculation::getInstance()->_calculateFormulaValue($fractionFormula);
            return true;
        }
        return false;
    }
    //    function convertToNumberIfFraction()
    /**
     * Get the decimal separator. If it has not yet been set explicitly, try to obtain number
     * formatting information from locale.
     *
     * @return string
     */
    public static function getDecimalSeparator()
    {
        if (!isset(self::$decimalSeparator)) {
            $localeconv = localeconv();
            self::$decimalSeparator = $localeconv['decimal_point'] != '' ? $localeconv['decimal_point'] : $localeconv['mon_decimal_point'];
            if (self::$decimalSeparator == '') {
                // Default to .
                self::$decimalSeparator = '.';
            }
        }
        return self::$decimalSeparator;
    }
    /**
     * Set the decimal separator. Only used by NumberFormat::toFormattedString()
     * to format output by \PhpOffice\PhpSpreadsheet\Writer\Html and \PhpOffice\PhpSpreadsheet\Writer\Pdf.
     *
     * @param string $pValue Character for decimal separator
     */
    public static function setDecimalSeparator($pValue)
    {
        self::$decimalSeparator = $pValue;
    }
    /**
     * Get the thousands separator. If it has not yet been set explicitly, try to obtain number
     * formatting information from locale.
     *
     * @return string
     */
    public static function getThousandsSeparator()
    {
        if (!isset(self::$thousandsSeparator)) {
            $localeconv = localeconv();
            self::$thousandsSeparator = $localeconv['thousands_sep'] != '' ? $localeconv['thousands_sep'] : $localeconv['mon_thousands_sep'];
            if (self::$thousandsSeparator == '') {
                // Default to .
                self::$thousandsSeparator = ',';
            }
        }
        return self::$thousandsSeparator;
    }
    /**
     * Set the thousands separator. Only used by NumberFormat::toFormattedString()
     * to format output by \PhpOffice\PhpSpreadsheet\Writer\Html and \PhpOffice\PhpSpreadsheet\Writer\Pdf.
     *
     * @param string $pValue Character for thousands separator
     */
    public static function setThousandsSeparator($pValue)
    {
        self::$thousandsSeparator = $pValue;
    }
    /**
     *    Get the currency code. If it has not yet been set explicitly, try to obtain the
     *        symbol information from locale.
     *
     * @return string
     */
    public static function getCurrencyCode()
    {
        if (!empty(self::$currencyCode)) {
            return self::$currencyCode;
        }
        self::$currencyCode = '$';
        $localeconv = localeconv();
        if (!empty($localeconv['currency_symbol'])) {
            self::$currencyCode = $localeconv['currency_symbol'];
            return self::$currencyCode;
        }
        if (!empty($localeconv['int_curr_symbol'])) {
            self::$currencyCode = $localeconv['int_curr_symbol'];
            return self::$currencyCode;
        }
        return self::$currencyCode;
    }
    /**
     * Set the currency code. Only used by NumberFormat::toFormattedString()
     *        to format output by \PhpOffice\PhpSpreadsheet\Writer\Html and \PhpOffice\PhpSpreadsheet\Writer\Pdf.
     *
     * @param string $pValue Character for currency code
     */
    public static function setCurrencyCode($pValue)
    {
        self::$currencyCode = $pValue;
    }
    /**
     * Convert SYLK encoded string to UTF-8.
     *
     * @param string $pValue
     *
     * @return string UTF-8 encoded string
     */
    public static function SYLKtoUTF8($pValue)
    {
        self::buildCharacterSets();
        // If there is no escape character in the string there is nothing to do
        if (strpos($pValue, '') === false) {
            return $pValue;
        }
        foreach (self::$SYLKCharacters as $k => $v) {
            $pValue = str_replace($k, $v, $pValue);
        }
        return $pValue;
    }
    /**
     * Retrieve any leading numeric part of a string, or return the full string if no leading numeric
     * (handles basic integer or float, but not exponent or non decimal).
     *
     * @param string $value
     *
     * @return mixed string or only the leading numeric part of the string
     */
    public static function testStringAsNumeric($value)
    {
        if (is_numeric($value)) {
            return $value;
        }
        $v = (double) $value;
        return is_numeric(substr($value, 0, strlen($v))) ? $v : $value;
    }
}