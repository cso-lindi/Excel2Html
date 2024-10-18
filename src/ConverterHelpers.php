<?php
namespace CSO\Excel2Html;
use CSO\Excel2Html\Exceptions\MalformattedRangeStringException;

class ConverterHelpers{
    /**
     * Converts a range string to a columns array
     * @param string $range range string like 'A-C'
     * @return string[] array of all columns in the range like ['A','B','C']
     */
    public static function RangeToColumnArray(string $range): array{
        $rangeArray = explode("-",$range);
        if(count($rangeArray) !== 2){
            throw new MalformattedRangeStringException();
        }
        
        $columns = [];
        for($column = $rangeArray[0]; $column !== $rangeArray[1]; ++$column) {
            array_push($columns, $column);
        }
        array_push($columns, $column);
        return $columns;
    }
}