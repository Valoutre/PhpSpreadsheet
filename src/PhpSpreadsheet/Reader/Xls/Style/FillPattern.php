<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Xls\Style;

use PhpOffice\PhpSpreadsheet\Style\Fill;
class FillPattern
{
    protected static $map = array(0 => Fill::FILL_NONE, 1 => Fill::FILL_SOLID, 2 => Fill::FILL_PATTERN_MEDIUMGRAY, 3 => Fill::FILL_PATTERN_DARKGRAY, 4 => Fill::FILL_PATTERN_LIGHTGRAY, 5 => Fill::FILL_PATTERN_DARKHORIZONTAL, 6 => Fill::FILL_PATTERN_DARKVERTICAL, 7 => Fill::FILL_PATTERN_DARKDOWN, 8 => Fill::FILL_PATTERN_DARKUP, 9 => Fill::FILL_PATTERN_DARKGRID, 10 => Fill::FILL_PATTERN_DARKTRELLIS, 11 => Fill::FILL_PATTERN_LIGHTHORIZONTAL, 12 => Fill::FILL_PATTERN_LIGHTVERTICAL, 13 => Fill::FILL_PATTERN_LIGHTDOWN, 14 => Fill::FILL_PATTERN_LIGHTUP, 15 => Fill::FILL_PATTERN_LIGHTGRID, 16 => Fill::FILL_PATTERN_LIGHTTRELLIS, 17 => Fill::FILL_PATTERN_GRAY125, 18 => Fill::FILL_PATTERN_GRAY0625);
    /**
     * Get fill pattern from index
     * OpenOffice documentation: 2.5.12.
     *
     * @param int $index
     *
     * @return string
     */
    public static function lookup($index)
    {
        if (isset(self::$map[$index])) {
            return self::$map[$index];
        }
        return Fill::FILL_NONE;
    }
}