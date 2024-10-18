<?php
use CSO\Excel2Html\ConverterHelpers;
use PHPUnit\Framework\TestCase;

final class ConverterHelpersTest extends TestCase{
    public function testColumnRangeWorkingAsExpected(): void{
        $columns = ConverterHelpers::RangeToColumnArray("Y-AC");
        $res = implode("|", $columns);
        $this->assertSame('Y|Z|AA|AB|AC', $res);
    }
}