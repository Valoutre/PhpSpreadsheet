<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Xls\Color;

class BIFF8
{
    protected static $map = array(8 => '000000', 9 => 'FFFFFF', 10 => 'FF0000', 11 => '00FF00', 12 => '0000FF', 13 => 'FFFF00', 14 => 'FF00FF', 15 => '00FFFF', 16 => '800000', 17 => '008000', 18 => '000080', 19 => '808000', 20 => '800080', 21 => '008080', 22 => 'C0C0C0', 23 => '808080', 24 => '9999FF', 25 => '993366', 26 => 'FFFFCC', 27 => 'CCFFFF', 28 => '660066', 29 => 'FF8080', 30 => '0066CC', 31 => 'CCCCFF', 32 => '000080', 33 => 'FF00FF', 34 => 'FFFF00', 35 => '00FFFF', 36 => '800080', 37 => '800000', 38 => '008080', 39 => '0000FF', 40 => '00CCFF', 41 => 'CCFFFF', 42 => 'CCFFCC', 43 => 'FFFF99', 44 => '99CCFF', 45 => 'FF99CC', 46 => 'CC99FF', 47 => 'FFCC99', 48 => '3366FF', 49 => '33CCCC', 50 => '99CC00', 51 => 'FFCC00', 52 => 'FF9900', 53 => 'FF6600', 54 => '666699', 55 => '969696', 56 => '003366', 57 => '339966', 58 => '003300', 59 => '333300', 60 => '993300', 61 => '993366', 62 => '333399', 63 => '333333');
    /**
     * Map color array from BIFF8 built-in color index.
     *
     * @param int $color
     *
     * @return array
     */
    public static function lookup($color)
    {
        if (isset(self::$map[$color])) {
            return array('rgb' => self::$map[$color]);
        }
        return array('rgb' => '000000');
    }
}