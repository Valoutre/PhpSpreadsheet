<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Xls\Color;

class BIFF5
{
    protected static $map = array(8 => '000000', 9 => 'FFFFFF', 10 => 'FF0000', 11 => '00FF00', 12 => '0000FF', 13 => 'FFFF00', 14 => 'FF00FF', 15 => '00FFFF', 16 => '800000', 17 => '008000', 18 => '000080', 19 => '808000', 20 => '800080', 21 => '008080', 22 => 'C0C0C0', 23 => '808080', 24 => '8080FF', 25 => '802060', 26 => 'FFFFC0', 27 => 'A0E0F0', 28 => '600080', 29 => 'FF8080', 30 => '0080C0', 31 => 'C0C0FF', 32 => '000080', 33 => 'FF00FF', 34 => 'FFFF00', 35 => '00FFFF', 36 => '800080', 37 => '800000', 38 => '008080', 39 => '0000FF', 40 => '00CFFF', 41 => '69FFFF', 42 => 'E0FFE0', 43 => 'FFFF80', 44 => 'A6CAF0', 45 => 'DD9CB3', 46 => 'B38FEE', 47 => 'E3E3E3', 48 => '2A6FF9', 49 => '3FB8CD', 50 => '488436', 51 => '958C41', 52 => '8E5E42', 53 => 'A0627A', 54 => '624FAC', 55 => '969696', 56 => '1D2FBE', 57 => '286676', 58 => '004500', 59 => '453E01', 60 => '6A2813', 61 => '85396A', 62 => '4A3285', 63 => '424242');
    /**
     * Map color array from BIFF5 built-in color index.
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