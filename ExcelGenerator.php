<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExcelGenerator
{
    public function generate($fio, $data)
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $sheet->setCellValue('B2', 'Группа:');
        $sheet->setCellValue('D2', 'ФИО:');
        $sheet->setCellValue('B3', 'Предмет');
        $sheet->setCellValue('C3', 'Оценка');
        $sheet->setCellValue('D3', 'Преподаватель');
        $sheet->setCellValue('E3', 'Дата');

        $sheet->setCellValue('E2', $fio);
        $sheet->setCellValue('C2', $data[0]['group_name'] ?? '');

        $rowCount = 4;
        foreach ($data as $row) {
            $sheet->setCellValue('B' . $rowCount, $row['subject_name'] ?? '');
            $sheet->setCellValue('C' . $rowCount, $row['mark'] ?? '');
            $sheet->setCellValue('D' . $rowCount, $row['teacher_name'] ?? '');
            $sheet->setCellValue('E' . $rowCount, $row['data'] ?? '');
            $rowCount++;
        }

        $fileName = sys_get_temp_dir() . '/' . $fio . '.xlsx';
        $writer = new Xlsx($spreadsheet);
        $writer->save($fileName);

        return $fileName;
    }
}