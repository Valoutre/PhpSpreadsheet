<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Xls;

class ErrorCode
{
    protected static $map = array(0 => '#NULL!', 7 => '#DIV/0!', 15 => '#VALUE!', 23 => '#REF!', 29 => '#NAME?', 36 => '#NUM!', 42 => '#N/A');
    /**
     * Map error code, e.g. '#N/A'.
     *
     * @param int $code
     *
     * @return bool|string
     */
    public static function lookup($code)
    {
        if (isset(self::$map[$code])) {
            return self::$map[$code];
        }
        return false;
    }
}