<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Xls\Style;

use PhpOffice\PhpSpreadsheet\Style\Border as StyleBorder;
class Border
{
    protected static $map = array(0 => StyleBorder::BORDER_NONE, 1 => StyleBorder::BORDER_THIN, 2 => StyleBorder::BORDER_MEDIUM, 3 => StyleBorder::BORDER_DASHED, 4 => StyleBorder::BORDER_DOTTED, 5 => StyleBorder::BORDER_THICK, 6 => StyleBorder::BORDER_DOUBLE, 7 => StyleBorder::BORDER_HAIR, 8 => StyleBorder::BORDER_MEDIUMDASHED, 9 => StyleBorder::BORDER_DASHDOT, 10 => StyleBorder::BORDER_MEDIUMDASHDOT, 11 => StyleBorder::BORDER_DASHDOTDOT, 12 => StyleBorder::BORDER_MEDIUMDASHDOTDOT, 13 => StyleBorder::BORDER_SLANTDASHDOT);
    /**
     * Map border style
     * OpenOffice documentation: 2.5.11.
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
        return StyleBorder::BORDER_NONE;
    }
}