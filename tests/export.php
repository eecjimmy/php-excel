<?php

/**
 * 1. `php -S 0.0.0.0:1234 -t ./`
 * 2. 访问http://127.0.0.1:1234/export.php
 */
require __DIR__ . '/../vendor/autoload.php';

use eecjimmy\Excel\ExportTrait;

class TestExport
{
    use ExportTrait;

    /**
     * @inheritDoc
     */
    public function getExcelFilename(): string
    {
        return "测试excel导出数据";
    }

    public function getExcelColumns(): array
    {
        return [
            "第一列", "第二列", "第三列"
        ];
    }

    public function getExcelData(): array
    {
        return [
            ['1-1', '1-2', '1-3'],
            ['2-1', '2-2', '2-3'],
            ['3-1', '3-2', '3-3'],
        ];
    }
}

$export  = new TestExport;
$export->export();
