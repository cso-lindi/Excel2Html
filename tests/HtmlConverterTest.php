<?php
use CSO\Excel2Html\Exceptions\SheetNotFoudException;
use PHPUnit\Framework\TestCase;
use CSO\Excel2Html\HtmlConverter;
use CSO\Excel2Html\StyleOptions;

final class HtmlConverterTest extends TestCase{
    //outputed html stays the same
    public function testHtmlProportionalFixedIsSame(): void{
        $conv = HtmlConverter::fromFilepath(
            'tests/assets/test.xlsx', 
            styleOption: StyleOptions::COLUMN_WIDTH_PROPORTIONAL | StyleOptions::WITH_COLUMN_WIDTH | StyleOptions::TABLE_SIZE_FIXED, 
            worksheetName:'TestTable', 
            columns:['A', 'B', 'C', 'D', 'E', 'F'],
            scale: 1.1);
        $res = $conv->getHtml();
        $expected = file_get_contents('tests/assets/results/testFixProp.html');
        $this->assertSame($expected, $res);
    }
    public function testHtmlWidthFixedIsSame(): void{
        $conv = HtmlConverter::fromFilepath(
            'tests/assets/test.xlsx', 
            styleOption: StyleOptions::WITH_COLUMN_WIDTH | StyleOptions::TABLE_SIZE_FIXED, 
            worksheetName:'TestTable', 
            columns:['A', 'B', 'C', 'D', 'E', 'F'],
            scale: 1.1);
        $res = $conv->getHtml();
        $expected = file_get_contents('tests/assets/results/testFixWidth.html');
        $this->assertSame($expected, $res);
    }
    public function testHtmlFixedIsSame(): void{
        $conv = HtmlConverter::fromFilepath(
            'tests/assets/test.xlsx', 
            styleOption: StyleOptions::TABLE_SIZE_FIXED, 
            worksheetName:'TestTable', 
            columns:['A', 'B', 'C', 'D', 'E', 'F'],
            scale: 1.1);
        $res = $conv->getHtml();
        $expected = file_get_contents('tests/assets/results/testFix.html');
        $this->assertSame($expected, $res);
    }
    public function testHtmlWidthIsSame(): void{
        $conv = HtmlConverter::fromFilepath(
            'tests/assets/test.xlsx', 
            styleOption: StyleOptions::WITH_COLUMN_WIDTH, 
            worksheetName:'TestTable', 
            columns:['A', 'B', 'C', 'D', 'E', 'F'],
            scale: 1.1);
        $res = $conv->getHtml();
        $expected = file_get_contents('tests/assets/results/testWidth.html');
        $this->assertSame($expected, $res);
    }
    public function testHtmlPropIsSame(): void{
        $conv = HtmlConverter::fromFilepath(
            'tests/assets/test.xlsx', 
            styleOption: StyleOptions::WITH_COLUMN_WIDTH | StyleOptions::COLUMN_WIDTH_PROPORTIONAL, 
            worksheetName:'TestTable', 
            columns:['A', 'B', 'C', 'D', 'E', 'F'],
            scale: 1.1);
        $res = $conv->getHtml();
        $expected = file_get_contents('tests/assets/results/testProp.html');
        $this->assertSame($expected, $res);
    }
    public function testHtmlPropIsSameWithRange(): void{
        $conv = HtmlConverter::fromFilepath(
            'tests/assets/test.xlsx', 
            styleOption: StyleOptions::WITH_COLUMN_WIDTH | StyleOptions::COLUMN_WIDTH_PROPORTIONAL, 
            worksheetName:'TestTable', 
            columns:['A', 'B-E', 'F'],
            scale: 1.1);
        $res = $conv->getHtml();
        $expected = file_get_contents('tests/assets/results/testProp.html');
        $this->assertSame($expected, $res);
    }
    public function testHtmlPropIsSameWithoutColumns(): void{
        $conv = HtmlConverter::fromFilepath(
            'tests/assets/test.xlsx', 
            styleOption: StyleOptions::WITH_COLUMN_WIDTH | StyleOptions::COLUMN_WIDTH_PROPORTIONAL, 
            worksheetName:'TestTable', 
            scale: 1.1);
        $res = $conv->getHtml();
        $expected = file_get_contents('tests/assets/results/testProp.html');
        $this->assertSame($expected, $res);
    }

    //Exceptions
    public function testCannotReadFromUnknownSheet(): void{
        $this->expectException(SheetNotFoudException::class);
        HtmlConverter::fromFilepath(
            'tests/assets/test.xlsx', 
            styleOption: StyleOptions::WITH_COLUMN_WIDTH | StyleOptions::COLUMN_WIDTH_PROPORTIONAL, 
            worksheetName:'gibberish', 
            columns:['A', 'B', 'C', 'D', 'E', 'F'],
            scale: 1.1);
    }
}