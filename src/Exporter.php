<?php


namespace eecjimmy\Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

/**
 * 导出数据到excel
 * @package eecjimmy\Excel
 */
abstract class Exporter
{
    /**
     * 导出的文件名，不包含.xlsx
     * @return string
     */
    abstract public function getFilename(): string;

    /**
     * 导出的文件列名
     * @return array
     */
    abstract public function getColumnTitles(): array;

    /**
     * 导出的数据二维数组
     * @return array
     */
    abstract public function getData(): array;

    /**
     * 导出excel表并下载
     * @return false|string
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function export()
    {
        $sp = new Spreadsheet();
        $sheet = $sp->getActiveSheet();

        $columnTitles = $this->getColumnTitles();
        $maxCol = count($columnTitles);
        $sheet->mergeCells(sprintf("A1:%s1", strtoupper(chr($maxCol + 64))));
        $sheet->setCellValueByColumnAndRow(1, 1, "【{$this->getFilename()}】");
        $sheet->getStyle("A1")->getAlignment()->setVertical(Alignment::HORIZONTAL_CENTER)->setHorizontal(Alignment::VERTICAL_CENTER);
        $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(18);
        $sheet->getRowDimension(1)->setRowHeight(50);
        foreach ($columnTitles as $key => $text) {
            $sheet->setCellValueByColumnAndRow($key + 1, 2, $text);
        }

        $current = 3;
        foreach ($this->getData() as $row) {
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
        header("Content-Disposition: must-revalidate, post-check=0, pre-check=0");
        header("Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        ob_start();
        $writer->save('php://output');
        return ob_get_clean();
    }
}