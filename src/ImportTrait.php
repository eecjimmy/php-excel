<?php

namespace eecjimmy\Excel;

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

/**
 * trait ImportTrait
 * @package eecjimmy\Excel
 */
trait ImportTrait
{
    /**
     * 需要导入的excel文件
     * 需要在服务器上面可以访问
     * @return string
     */
    abstract public function getExcelFile(): string;

    /**
     * excel表的列数量
     * @return int
     */
    abstract public function getColumnNum(): int;

    /**
     * 导入单行数据
     * @param array $row
     * @param int $index
     * @return bool
     */
    abstract public function saveRow(array $row, int $index): bool;

    /**
     * 导入操作
     * @return int 导入成功的数量
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function import(): int
    {
        $excel = $this->getExcelFile();
        $reader = new Xlsx();
        $sheet = $reader->load($excel)->setActiveSheetIndex(0);
        $r = $sheet->getHighestDataRow('A');
        $maxCol = $this->getColumnNum();
        // 从第二行开始, 第一行作为标题
        for ($row = 2; $row <= $r; $row++) {
            $rowData = [];
            for ($col = 0; $col < $maxCol; $col++) {
                $coordinate = Coordinate::stringFromColumnIndex($col + 1) . $row; // cell name, eg. A1
                $cell = $sheet->getCell($coordinate);
                $v = $cell->getFormattedValue();
                $rowData[] = $v;
            }

            if (empty($rowData)) break;

            if (!$this->saveRow($rowData, $row)) {
                return false;
            }
        }

        return true;
    }

    /**
     * 返回导入模板的下载链接
     * @return string
     */
    abstract public function getTemplateURL(): string;
}
