<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Xls\Color;

class BuiltIn
{
    protected static $map = array(0 => '000000', 1 => 'FFFFFF', 2 => 'FF0000', 3 => '00FF00', 4 => '0000FF', 5 => 'FFFF00', 6 => 'FF00FF', 7 => '00FFFF', 64 => '000000', 65 => 'FFFFFF');
    /**
     * Map built-in color to RGB value.
     *
     * @param int $color Indexed color
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