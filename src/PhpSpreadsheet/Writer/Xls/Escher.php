<?php

namespace PhpOffice\PhpSpreadsheet\Writer\Xls;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer;
use PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer\SpgrContainer;
use PhpOffice\PhpSpreadsheet\Shared\Escher\DgContainer\SpgrContainer\SpContainer;
use PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer;
use PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer\BstoreContainer;
use PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer\BstoreContainer\BSE;
use PhpOffice\PhpSpreadsheet\Shared\Escher\DggContainer\BstoreContainer\BSE\Blip;
class Escher
{
    /**
     * The object we are writing.
     */
    private $object;
    /**
     * The written binary data.
     */
    private $data;
    /**
     * Shape offsets. Positions in binary stream where a new shape record begins.
     *
     * @var array
     */
    private $spOffsets;
    /**
     * Shape types.
     *
     * @var array
     */
    private $spTypes;
    /**
     * Constructor.
     *
     * @param mixed $object
     */
    public function __construct($object)
    {
        $this->object = $object;
    }
    /**
     * Process the object to be written.
     *
     * @return string
     */
    public function close()
    {
        // initialize
        $this->data = '';
        switch (get_class($this->object)) {
            case 'Escher':
                if ($dggContainer = $this->object->getDggContainer()) {
                    $writer = new self($dggContainer);
                    $this->data = $writer->close();
                } elseif ($dgContainer = $this->object->getDgContainer()) {
                    $writer = new self($dgContainer);
                    $this->data = $writer->close();
                    $this->spOffsets = $writer->getSpOffsets();
                    $this->spTypes = $writer->getSpTypes();
                }
                break;
            case 'DggContainer':
                // this is a container record
                // initialize
                $innerData = '';
                // write the dgg
                $recVer = 0;
                $recInstance = 0;
                $recType = 61446;
                $recVerInstance = $recVer;
                $recVerInstance |= $recInstance << 4;
                // dgg data
                $dggData = pack('VVVV', $this->object->getSpIdMax(), $this->object->getCDgSaved() + 1, $this->object->getCSpSaved(), $this->object->getCDgSaved());
                // add file identifier clusters (one per drawing)
                $IDCLs = $this->object->getIDCLs();
                foreach ($IDCLs as $dgId => $maxReducedSpId) {
                    $dggData .= pack('VV', $dgId, $maxReducedSpId + 1);
                }
                $header = pack('vvV', $recVerInstance, $recType, strlen($dggData));
                $innerData .= $header . $dggData;
                // write the bstoreContainer
                if ($bstoreContainer = $this->object->getBstoreContainer()) {
                    $writer = new self($bstoreContainer);
                    $innerData .= $writer->close();
                }
                // write the record
                $recVer = 15;
                $recInstance = 0;
                $recType = 61440;
                $length = strlen($innerData);
                $recVerInstance = $recVer;
                $recVerInstance |= $recInstance << 4;
                $header = pack('vvV', $recVerInstance, $recType, $length);
                $this->data = $header . $innerData;
                break;
            case 'BstoreContainer':
                // this is a container record
                // initialize
                $innerData = '';
                // treat the inner data
                if ($BSECollection = $this->object->getBSECollection()) {
                    foreach ($BSECollection as $BSE) {
                        $writer = new self($BSE);
                        $innerData .= $writer->close();
                    }
                }
                // write the record
                $recVer = 15;
                $recInstance = count($this->object->getBSECollection());
                $recType = 61441;
                $length = strlen($innerData);
                $recVerInstance = $recVer;
                $recVerInstance |= $recInstance << 4;
                $header = pack('vvV', $recVerInstance, $recType, $length);
                $this->data = $header . $innerData;
                break;
            case 'BSE':
                // this is a semi-container record
                // initialize
                $innerData = '';
                // here we treat the inner data
                if ($blip = $this->object->getBlip()) {
                    $writer = new self($blip);
                    $innerData .= $writer->close();
                }
                // initialize
                $data = '';
                $btWin32 = $this->object->getBlipType();
                $btMacOS = $this->object->getBlipType();
                $data .= pack('CC', $btWin32, $btMacOS);
                $rgbUid = pack('VVVV', 0, 0, 0, 0);
                // todo
                $data .= $rgbUid;
                $tag = 0;
                $size = strlen($innerData);
                $cRef = 1;
                $foDelay = 0;
                //todo
                $unused1 = 0;
                $cbName = 0;
                $unused2 = 0;
                $unused3 = 0;
                $data .= pack('vVVVCCCC', $tag, $size, $cRef, $foDelay, $unused1, $cbName, $unused2, $unused3);
                $data .= $innerData;
                // write the record
                $recVer = 2;
                $recInstance = $this->object->getBlipType();
                $recType = 61447;
                $length = strlen($data);
                $recVerInstance = $recVer;
                $recVerInstance |= $recInstance << 4;
                $header = pack('vvV', $recVerInstance, $recType, $length);
                $this->data = $header;
                $this->data .= $data;
                break;
            case 'Blip':
                // this is an atom record
                // write the record
                switch ($this->object->getParent()->getBlipType()) {
                    case BSE::BLIPTYPE_JPEG:
                        // initialize
                        $innerData = '';
                        $rgbUid1 = pack('VVVV', 0, 0, 0, 0);
                        // todo
                        $innerData .= $rgbUid1;
                        $tag = 255;
                        // todo
                        $innerData .= pack('C', $tag);
                        $innerData .= $this->object->getData();
                        $recVer = 0;
                        $recInstance = 1130;
                        $recType = 61469;
                        $length = strlen($innerData);
                        $recVerInstance = $recVer;
                        $recVerInstance |= $recInstance << 4;
                        $header = pack('vvV', $recVerInstance, $recType, $length);
                        $this->data = $header;
                        $this->data .= $innerData;
                        break;
                    case BSE::BLIPTYPE_PNG:
                        // initialize
                        $innerData = '';
                        $rgbUid1 = pack('VVVV', 0, 0, 0, 0);
                        // todo
                        $innerData .= $rgbUid1;
                        $tag = 255;
                        // todo
                        $innerData .= pack('C', $tag);
                        $innerData .= $this->object->getData();
                        $recVer = 0;
                        $recInstance = 1760;
                        $recType = 61470;
                        $length = strlen($innerData);
                        $recVerInstance = $recVer;
                        $recVerInstance |= $recInstance << 4;
                        $header = pack('vvV', $recVerInstance, $recType, $length);
                        $this->data = $header;
                        $this->data .= $innerData;
                        break;
                }
                break;
            case 'DgContainer':
                // this is a container record
                // initialize
                $innerData = '';
                // write the dg
                $recVer = 0;
                $recInstance = $this->object->getDgId();
                $recType = 61448;
                $length = 8;
                $recVerInstance = $recVer;
                $recVerInstance |= $recInstance << 4;
                $header = pack('vvV', $recVerInstance, $recType, $length);
                // number of shapes in this drawing (including group shape)
                $countShapes = count($this->object->getSpgrContainer()->getChildren());
                $innerData .= $header . pack('VV', $countShapes, $this->object->getLastSpId());
                // write the spgrContainer
                if ($spgrContainer = $this->object->getSpgrContainer()) {
                    $writer = new self($spgrContainer);
                    $innerData .= $writer->close();
                    // get the shape offsets relative to the spgrContainer record
                    $spOffsets = $writer->getSpOffsets();
                    $spTypes = $writer->getSpTypes();
                    // save the shape offsets relative to dgContainer
                    foreach ($spOffsets as &$spOffset) {
                        $spOffset += 24;
                    }
                    $this->spOffsets = $spOffsets;
                    $this->spTypes = $spTypes;
                }
                // write the record
                $recVer = 15;
                $recInstance = 0;
                $recType = 61442;
                $length = strlen($innerData);
                $recVerInstance = $recVer;
                $recVerInstance |= $recInstance << 4;
                $header = pack('vvV', $recVerInstance, $recType, $length);
                $this->data = $header . $innerData;
                break;
            case 'SpgrContainer':
                // this is a container record
                // initialize
                $innerData = '';
                // initialize spape offsets
                $totalSize = 8;
                $spOffsets = array();
                $spTypes = array();
                // treat the inner data
                foreach ($this->object->getChildren() as $spContainer) {
                    $writer = new self($spContainer);
                    $spData = $writer->close();
                    $innerData .= $spData;
                    // save the shape offsets (where new shape records begin)
                    $totalSize += strlen($spData);
                    $spOffsets[] = $totalSize;
                    $spTypes = array_merge($spTypes, $writer->getSpTypes());
                }
                // write the record
                $recVer = 15;
                $recInstance = 0;
                $recType = 61443;
                $length = strlen($innerData);
                $recVerInstance = $recVer;
                $recVerInstance |= $recInstance << 4;
                $header = pack('vvV', $recVerInstance, $recType, $length);
                $this->data = $header . $innerData;
                $this->spOffsets = $spOffsets;
                $this->spTypes = $spTypes;
                break;
            case 'SpContainer':
                // initialize
                $data = '';
                // build the data
                // write group shape record, if necessary?
                if ($this->object->getSpgr()) {
                    $recVer = 1;
                    $recInstance = 0;
                    $recType = 61449;
                    $length = 16;
                    $recVerInstance = $recVer;
                    $recVerInstance |= $recInstance << 4;
                    $header = pack('vvV', $recVerInstance, $recType, $length);
                    $data .= $header . pack('VVVV', 0, 0, 0, 0);
                }
                $this->spTypes[] = $this->object->getSpType();
                // write the shape record
                $recVer = 2;
                $recInstance = $this->object->getSpType();
                // shape type
                $recType = 61450;
                $length = 8;
                $recVerInstance = $recVer;
                $recVerInstance |= $recInstance << 4;
                $header = pack('vvV', $recVerInstance, $recType, $length);
                $data .= $header . pack('VV', $this->object->getSpId(), $this->object->getSpgr() ? 5 : 2560);
                // the options
                if ($this->object->getOPTCollection()) {
                    $optData = '';
                    $recVer = 3;
                    $recInstance = count($this->object->getOPTCollection());
                    $recType = 61451;
                    foreach ($this->object->getOPTCollection() as $property => $value) {
                        $optData .= pack('vV', $property, $value);
                    }
                    $length = strlen($optData);
                    $recVerInstance = $recVer;
                    $recVerInstance |= $recInstance << 4;
                    $header = pack('vvV', $recVerInstance, $recType, $length);
                    $data .= $header . $optData;
                }
                // the client anchor
                if ($this->object->getStartCoordinates()) {
                    $clientAnchorData = '';
                    $recVer = 0;
                    $recInstance = 0;
                    $recType = 61456;
                    // start coordinates
                    list($column, $row) = Coordinate::coordinateFromString($this->object->getStartCoordinates());
                    $c1 = Coordinate::columnIndexFromString($column) - 1;
                    $r1 = $row - 1;
                    // start offsetX
                    $startOffsetX = $this->object->getStartOffsetX();
                    // start offsetY
                    $startOffsetY = $this->object->getStartOffsetY();
                    // end coordinates
                    list($column, $row) = Coordinate::coordinateFromString($this->object->getEndCoordinates());
                    $c2 = Coordinate::columnIndexFromString($column) - 1;
                    $r2 = $row - 1;
                    // end offsetX
                    $endOffsetX = $this->object->getEndOffsetX();
                    // end offsetY
                    $endOffsetY = $this->object->getEndOffsetY();
                    $clientAnchorData = pack('vvvvvvvvv', $this->object->getSpFlag(), $c1, $startOffsetX, $r1, $startOffsetY, $c2, $endOffsetX, $r2, $endOffsetY);
                    $length = strlen($clientAnchorData);
                    $recVerInstance = $recVer;
                    $recVerInstance |= $recInstance << 4;
                    $header = pack('vvV', $recVerInstance, $recType, $length);
                    $data .= $header . $clientAnchorData;
                }
                // the client data, just empty for now
                if (!$this->object->getSpgr()) {
                    $clientDataData = '';
                    $recVer = 0;
                    $recInstance = 0;
                    $recType = 61457;
                    $length = strlen($clientDataData);
                    $recVerInstance = $recVer;
                    $recVerInstance |= $recInstance << 4;
                    $header = pack('vvV', $recVerInstance, $recType, $length);
                    $data .= $header . $clientDataData;
                }
                // write the record
                $recVer = 15;
                $recInstance = 0;
                $recType = 61444;
                $length = strlen($data);
                $recVerInstance = $recVer;
                $recVerInstance |= $recInstance << 4;
                $header = pack('vvV', $recVerInstance, $recType, $length);
                $this->data = $header . $data;
                break;
        }
        return $this->data;
    }
    /**
     * Gets the shape offsets.
     *
     * @return array
     */
    public function getSpOffsets()
    {
        return $this->spOffsets;
    }
    /**
     * Gets the shape types.
     *
     * @return array
     */
    public function getSpTypes()
    {
        return $this->spTypes;
    }
}