<?php
use CSO\Excel2Html\HtmlConverter;
use CSO\Excel2Html\StyleOptions;

require_once 'vendor/autoload.php';

$conv = HtmlConverter::fromFilepath(
    'tests/assets/test.xlsx', 
    styleOption: StyleOptions::TABLE_SIZE_FIXED | StyleOptions::WITH_COLUMN_WIDTH | StyleOptions::COLUMN_WIDTH_PROPORTIONAL, 
    worksheetName:'TestTable', 
    columns:['A', 'B', 'C', 'D', 'E', 'F'],
    scale: 1.1);

echo $conv->getHtml();