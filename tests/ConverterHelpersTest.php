<?php
use CSO\Excel2Html\ConverterHelpers;
use CSO\Excel2Html\Exceptions\MalformattedRangeStringException;
use PHPUnit\Framework\TestCase;

final class ConverterHelpersTest extends TestCase{
    public function testColumnRangeWorkingAsExpected(): void{
        $columns = ConverterHelpers::RangeToColumnArray("Y-AC");
        $res = implode("|", $columns);
        $this->assertSame('Y|Z|AA|AB|AC', $res);
    }
    public function testColumnRangeErrorOnWrongFormat(): void{
        $this->expectException(MalformattedRangeStringException::class);
        $columns = ConverterHelpers::RangeToColumnArray("-AC");
    }
    public function testColumnRangeErrorOnTwoLetterColumnFirst(): void{
        $this->expectException(MalformattedRangeStringException::class);
        $columns = ConverterHelpers::RangeToColumnArray("AC-Y");
    }
    public function testColumnRangeErrorOnHigherColumnFirst(): void{
        $this->expectException(MalformattedRangeStringException::class);
        $columns = ConverterHelpers::RangeToColumnArray("Z-B");
    }
}