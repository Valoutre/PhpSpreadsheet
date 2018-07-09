<?php

namespace PhpOffice\PhpSpreadsheet\Reader\Xls;

class MD5
{
    // Context
    private $a;
    private $b;
    private $c;
    private $d;
    /**
     * MD5 stream constructor.
     */
    public function __construct()
    {
        $this->reset();
    }
    /**
     * Reset the MD5 stream context.
     */
    public function reset()
    {
        $this->a = 1732584193;
        $this->b = 0.0;
        $this->c = 0.0;
        $this->d = 271733878;
    }
    /**
     * Get MD5 stream context.
     *
     * @return string
     */
    public function getContext()
    {
        $s = '';
        foreach (array('a', 'b', 'c', 'd') as $i) {
            $v = $this->{$i};
            $s .= chr($v & 255);
            $s .= chr($v >> 8 & 255);
            $s .= chr($v >> 16 & 255);
            $s .= chr($v >> 24 & 255);
        }
        return $s;
    }
    /**
     * Add data to context.
     *
     * @param string $data Data to add
     */
    public function add($data)
    {
        $words = array_values(unpack('V16', $data));
        $A = $this->a;
        $B = $this->b;
        $C = $this->c;
        $D = $this->d;
        $F = array('self', 'f');
        $G = array('self', 'g');
        $H = array('self', 'h');
        $I = array('self', 'i');
        // ROUND 1
        self::step($F, $A, $B, $C, $D, $words[0], 7, 3614090360.0);
        self::step($F, $D, $A, $B, $C, $words[1], 12, 0.0);
        self::step($F, $C, $D, $A, $B, $words[2], 17, 606105819);
        self::step($F, $B, $C, $D, $A, $words[3], 22, 0.0);
        self::step($F, $A, $B, $C, $D, $words[4], 7, 4118548399.0);
        self::step($F, $D, $A, $B, $C, $words[5], 12, 1200080426);
        self::step($F, $C, $D, $A, $B, $words[6], 17, 2821735955.0);
        self::step($F, $B, $C, $D, $A, $words[7], 22, 4249261313.0);
        self::step($F, $A, $B, $C, $D, $words[8], 7, 1770035416);
        self::step($F, $D, $A, $B, $C, $words[9], 12, 2336552879.0);
        self::step($F, $C, $D, $A, $B, $words[10], 17, 4294925233.0);
        self::step($F, $B, $C, $D, $A, $words[11], 22, 0.0);
        self::step($F, $A, $B, $C, $D, $words[12], 7, 1804603682);
        self::step($F, $D, $A, $B, $C, $words[13], 12, 4254626195.0);
        self::step($F, $C, $D, $A, $B, $words[14], 17, 0.0);
        self::step($F, $B, $C, $D, $A, $words[15], 22, 1236535329);
        // ROUND 2
        self::step($G, $A, $B, $C, $D, $words[1], 5, 0.0);
        self::step($G, $D, $A, $B, $C, $words[6], 9, 3225465664.0);
        self::step($G, $C, $D, $A, $B, $words[11], 14, 643717713);
        self::step($G, $B, $C, $D, $A, $words[0], 20, 0.0);
        self::step($G, $A, $B, $C, $D, $words[5], 5, 3593408605.0);
        self::step($G, $D, $A, $B, $C, $words[10], 9, 38016083);
        self::step($G, $C, $D, $A, $B, $words[15], 14, 0.0);
        self::step($G, $B, $C, $D, $A, $words[4], 20, 0.0);
        self::step($G, $A, $B, $C, $D, $words[9], 5, 568446438);
        self::step($G, $D, $A, $B, $C, $words[14], 9, 3275163606.0);
        self::step($G, $C, $D, $A, $B, $words[3], 14, 4107603335.0);
        self::step($G, $B, $C, $D, $A, $words[8], 20, 1163531501);
        self::step($G, $A, $B, $C, $D, $words[13], 5, 0.0);
        self::step($G, $D, $A, $B, $C, $words[2], 9, 0.0);
        self::step($G, $C, $D, $A, $B, $words[7], 14, 1735328473);
        self::step($G, $B, $C, $D, $A, $words[12], 20, 2368359562.0);
        // ROUND 3
        self::step($H, $A, $B, $C, $D, $words[5], 4, 4294588738.0);
        self::step($H, $D, $A, $B, $C, $words[8], 11, 2272392833.0);
        self::step($H, $C, $D, $A, $B, $words[11], 16, 1839030562);
        self::step($H, $B, $C, $D, $A, $words[14], 23, 0.0);
        self::step($H, $A, $B, $C, $D, $words[1], 4, 0.0);
        self::step($H, $D, $A, $B, $C, $words[4], 11, 1272893353);
        self::step($H, $C, $D, $A, $B, $words[7], 16, 4139469664.0);
        self::step($H, $B, $C, $D, $A, $words[10], 23, 0.0);
        self::step($H, $A, $B, $C, $D, $words[13], 4, 681279174);
        self::step($H, $D, $A, $B, $C, $words[0], 11, 0.0);
        self::step($H, $C, $D, $A, $B, $words[3], 16, 0.0);
        self::step($H, $B, $C, $D, $A, $words[6], 23, 76029189);
        self::step($H, $A, $B, $C, $D, $words[9], 4, 3654602809.0);
        self::step($H, $D, $A, $B, $C, $words[12], 11, 0.0);
        self::step($H, $C, $D, $A, $B, $words[15], 16, 530742520);
        self::step($H, $B, $C, $D, $A, $words[2], 23, 3299628645.0);
        // ROUND 4
        self::step($I, $A, $B, $C, $D, $words[0], 6, 4096336452.0);
        self::step($I, $D, $A, $B, $C, $words[7], 10, 1126891415);
        self::step($I, $C, $D, $A, $B, $words[14], 15, 2878612391.0);
        self::step($I, $B, $C, $D, $A, $words[5], 21, 4237533241.0);
        self::step($I, $A, $B, $C, $D, $words[12], 6, 1700485571);
        self::step($I, $D, $A, $B, $C, $words[3], 10, 2399980690.0);
        self::step($I, $C, $D, $A, $B, $words[10], 15, 0.0);
        self::step($I, $B, $C, $D, $A, $words[1], 21, 2240044497.0);
        self::step($I, $A, $B, $C, $D, $words[8], 6, 1873313359);
        self::step($I, $D, $A, $B, $C, $words[15], 10, 0.0);
        self::step($I, $C, $D, $A, $B, $words[6], 15, 2734768916.0);
        self::step($I, $B, $C, $D, $A, $words[13], 21, 1309151649);
        self::step($I, $A, $B, $C, $D, $words[4], 6, 0.0);
        self::step($I, $D, $A, $B, $C, $words[11], 10, 3174756917.0);
        self::step($I, $C, $D, $A, $B, $words[2], 15, 718787259);
        self::step($I, $B, $C, $D, $A, $words[9], 21, 0.0);
        $this->a = $this->a + $A & 4294967295.0;
        $this->b = $this->b + $B & 4294967295.0;
        $this->c = $this->c + $C & 4294967295.0;
        $this->d = $this->d + $D & 4294967295.0;
    }
    private static function f($X, $Y, $Z)
    {
        return $X & $Y | ~$X & $Z;
    }
    private static function g($X, $Y, $Z)
    {
        return $X & $Z | $Y & ~$Z;
    }
    private static function h($X, $Y, $Z)
    {
        return $X ^ $Y ^ $Z;
    }
    private static function i($X, $Y, $Z)
    {
        return $Y ^ ($X | ~$Z);
    }
    private static function step($func, &$A, $B, $C, $D, $M, $s, $t)
    {
        $A = $A + call_user_func($func, $B, $C, $D) + $M + $t & 4294967295.0;
        $A = self::rotate($A, $s);
        $A = $B + $A & 4294967295.0;
    }
    private static function rotate($decimal, $bits)
    {
        $binary = str_pad(decbin($decimal), 32, '0', STR_PAD_LEFT);
        return bindec(substr($binary, $bits) . substr($binary, 0, $bits));
    }
}