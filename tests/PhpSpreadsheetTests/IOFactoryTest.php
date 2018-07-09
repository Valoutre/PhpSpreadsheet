<?php

namespace PhpOffice\PhpSpreadsheetTests;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer;
use PHPUnit\Framework\TestCase;

class IOFactoryTest extends TestCase
{
    /**
     * @dataProvider providerCreateWriter
     *
     * @param string $name
     * @param string $expected
     */
    public function testCreateWriter($name, $expected)
    {
        $spreadsheet = new Spreadsheet();
        $actual = IOFactory::createWriter($spreadsheet, $name);
        self::assertInstanceOf($expected, $actual);
    }

    public function providerCreateWriter()
    {
        return [
            ['Xls', 'Xls'],
            ['Xlsx', 'Xlsx'],
            ['Ods', 'Ods'],
            ['Csv', 'Csv'],
            ['Html', 'Html'],
            ['Mpdf', 'Mpdf'],
            ['Tcpdf', 'Tcpdf'],
            ['Dompdf', 'Dompdf'],
        ];
    }

    public function testRegisterWriter()
    {
        IOFactory::registerWriter('Pdf', 'Mpdf');
        $spreadsheet = new Spreadsheet();
        $actual = IOFactory::createWriter($spreadsheet, 'Pdf');
        self::assertInstanceOf('Mpdf', $actual);
    }

    /**
     * @dataProvider providerCreateReader
     *
     * @param string $name
     * @param string $expected
     */
    public function testCreateReader($name, $expected)
    {
        $actual = IOFactory::createReader($name);
        self::assertInstanceOf($expected, $actual);
    }

    public function providerCreateReader()
    {
        return [
            ['Xls', 'Xls'],
            ['Xlsx', 'Xlsx'],
            ['Xml', 'Xml'],
            ['Ods', 'Ods'],
            ['Gnumeric', 'Gnumeric'],
            ['Csv', 'Csv'],
            ['Slk', 'Slk'],
            ['Html', 'Html'],
        ];
    }

    public function testRegisterReader()
    {
        IOFactory::registerReader('Custom', 'Html');
        $actual = IOFactory::createReader('Custom');
        self::assertInstanceOf('Html', $actual);
    }

    /**
     * @dataProvider providerIdentify
     *
     * @param string $file
     * @param string $expectedName
     * @param string $expectedClass
     */
    public function testIdentify($file, $expectedName, $expectedClass)
    {
        $actual = IOFactory::identify($file);
        self::assertSame($expectedName, $actual);
    }

    /**
     * @dataProvider providerIdentify
     *
     * @param string $file
     * @param string $expectedName
     * @param string $expectedClass
     */
    public function testCreateReaderForFile($file, $expectedName, $expectedClass)
    {
        $actual = IOFactory::createReaderForFile($file);
        self::assertInstanceOf($expectedClass, $actual);
    }

    public function providerIdentify()
    {
        return [
            ['../samples/templates/26template.xlsx', 'Xlsx', 'Xlsx'],
            ['../samples/templates/GnumericTest.gnumeric', 'Gnumeric', 'Gnumeric'],
            ['../samples/templates/30template.xls', 'Xls', 'Xls'],
            ['../samples/templates/OOCalcTest.ods', 'Ods', 'Ods'],
            ['../samples/templates/SylkTest.slk', 'Slk', 'Slk'],
            ['../samples/templates/Excel2003XMLTest.xml', 'Xml', 'Xml'],
            ['../samples/templates/46readHtml.html', 'Html', 'Html'],
        ];
    }

    public function testIdentifyNonExistingFileThrowException()
    {
        $this->expectException('InvalidArgumentException');

        IOFactory::identify('/non/existing/file');
    }

    public function testIdentifyExistingDirectoryThrowExceptions()
    {
        $this->expectException('InvalidArgumentException');

        IOFactory::identify('.');
    }

    public function testRegisterInvalidWriter()
    {
        $this->expectException('Exception');

        IOFactory::registerWriter('foo', 'bar');
    }

    public function testRegisterInvalidReader()
    {
        $this->expectException('Exception');

        IOFactory::registerReader('foo', 'bar');
    }
}
