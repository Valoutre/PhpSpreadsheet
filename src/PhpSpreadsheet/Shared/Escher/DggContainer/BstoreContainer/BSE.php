<?php

namespace PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer\BstoreContainer;

class BSE
{
    const BLIPTYPE_ERROR = 0;
    const BLIPTYPE_UNKNOWN = 1;
    const BLIPTYPE_EMF = 2;
    const BLIPTYPE_WMF = 3;
    const BLIPTYPE_PICT = 4;
    const BLIPTYPE_JPEG = 5;
    const BLIPTYPE_PNG = 6;
    const BLIPTYPE_DIB = 7;
    const BLIPTYPE_TIFF = 17;
    const BLIPTYPE_CMYKJPEG = 18;
    /**
     * The parent BLIP Store Entry Container.
     *
     * @var \PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer\BstoreContainer
     */
    private $parent;
    /**
     * The BLIP (Big Large Image or Picture).
     *
     * @var BSE\Blip
     */
    private $blip;
    /**
     * The BLIP type.
     *
     * @var int
     */
    private $blipType;
    /**
     * Set parent BLIP Store Entry Container.
     *
     * @param \PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer\BstoreContainer $parent
     */
    public function setParent($parent)
    {
        $this->parent = $parent;
    }
    /**
     * Get the BLIP.
     *
     * @return BSE\Blip
     */
    public function getBlip()
    {
        return $this->blip;
    }
    /**
     * Set the BLIP.
     *
     * @param BSE\Blip $blip
     */
    public function setBlip($blip)
    {
        $this->blip = $blip;
        $blip->setParent($this);
    }
    /**
     * Get the BLIP type.
     *
     * @return int
     */
    public function getBlipType()
    {
        return $this->blipType;
    }
    /**
     * Set the BLIP type.
     *
     * @param int $blipType
     */
    public function setBlipType($blipType)
    {
        $this->blipType = $blipType;
    }
}