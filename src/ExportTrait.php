<?php

namespace eecjimmy\Excel;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\IOFactory;

/**
 * Excel export trait
 * @package eecjimmy\Excel
 */
trait ExportTrait
{
    /**
     * 获取导出的文件名称
     * 不包括.xlsx
     *
     * @return string
     */
    abstract public function getExcelFilename(): string;

    /**
     * 获取导出的列标题
     *
     * @return array
     */
    abstract public function getExcelColumns(): array;


    /**
     * 获取导出的数据
     *
     * @return array 一个二维数组, 二维数组的元素个数与`getExcelColumns`一致
     */
    abstract public function getExcelData(): array;

    /**
     * 导出数据
     */
    public function export()
    {
        $sp = new Spreadsheet();
        $sheet = $sp->getActiveSheet();

        $columnTitles = $this->getExcelColumns();
        $maxCol = count($columnTitles);
        $sheet->mergeCells(sprintf("A1:%s1", strtoupper(chr($maxCol + 64))));
        $sheet->setCellValueByColumnAndRow(1, 1, "【{$this->getExcelFilename()}】");
        $sheet->getStyle("A1")->getAlignment()->setVertical(Alignment::HORIZONTAL_CENTER)->setHorizontal(Alignment::VERTICAL_CENTER);
        $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(18);
        $sheet->getRowDimension(1)->setRowHeight(50);
        foreach ($columnTitles as $key => $text) {
            $sheet->setCellValueByColumnAndRow($key + 1, 2, $text);
        }

        $current = 3;
        foreach ($this->getExcelData() as $row) {
            foreach ($row as $key => $value) {
                $c = $sheet->getCellByColumnAndRow($key + 1, $current)->getCoordinate();
                if (is_array($value)) {
                    $v = $value;
                    $value = $v['value'] ?? '';
                    $color = $v['color'] ?? '';
                    if ($color) {
                        $sheet->getStyle($c)->getFont()->getColor()->setRGB($color);
                    }
                }
                $sheet->setCellValue($c, $value);
            }
            $current++;
        }

        $writer = IOFactory::createWriter($sp, "Xlsx");

        header("Pragma: attachment");
        header("Accept-Ranges: bytes");
        header("Expires: 0");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
        header("Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-Disposition: attachment; filename=\"{$this->getExcelFilename()}.xlsx\"; filename*=utf-8''{$this->getExcelFilename()}.xlsx");
        ob_start();
        $writer->save('php://output');
        return ob_get_clean();
    }
}
