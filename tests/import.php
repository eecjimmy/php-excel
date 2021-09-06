<?php

use eecjimmy\Excel\ImportTrait;

require __DIR__ . '/../vendor/autoload.php';
class TestImport
{
    use ImportTrait;
    public function getExcelFile(): string
    {
        return './test.xlsx';
    }
    public function getColumnNum(): int
    {
        return 3;
    }

    public function saveRow(array $row, int $index): bool
    {
        echo sprintf("第%d行数据: %s\n", $index, json_encode($row, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES));
        return true;
    }

    public function getTemplateURL(): string
    {
        return "/text.xlsx";
    }
}

$m = new TestImport;
$m->import();
