<?php
session_start();
include('connection.php');
include('log-write.php');
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

// Проверяем, авторизован ли пользователь и выбран ли студент
if (!isset($_SESSION['user']) || !isset($_SESSION['student'])) {
    $_SESSION['message'] = "Доступ запрещен или студент не выбран.";
    header("Location: ../html/record-book.php");
    exit;
}

$styleForHeader = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THICK,
        ],
    ],
    'font' => [
        'bold' => true
    ],
];
$styleForData = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => Border::BORDER_THIN,
        ],
    ],
];
$StyleForMark = [
    'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER, // Горизонтальное выравнивание по центру
    ],
];

$fio = $_SESSION['student'];

// Кэширование данных (если данные не меняются часто)
$cacheFile = __DIR__ . '/cache/generate-excel_' . md5($fio) . '.json';
$cacheTime = 3600; // 1 час

if (file_exists($cacheFile) && (time() - filemtime($cacheFile) < $cacheTime)) {
    // Используем данные из кэша
    $data = json_decode(file_get_contents($cacheFile), true);
} else {
    // Объединенный SQL-запрос для получения всех данных за один раз
    $sql = "
        SELECT DISTINCT
            s.id_student,
            sg.naimenovanie AS group_name,
            sub.naimenovanie AS subject_name,
            e.mark,
            t.fio AS teacher_name,
            e.data
        FROM student s
        LEFT JOIN appointment_student aps ON s.id_student = aps.id_student
        LEFT JOIN studgroup sg ON aps.id_group = sg.id_group
        LEFT JOIN exam e ON s.id_student = e.id_student
        LEFT JOIN subject sub ON e.id_subject = sub.id_subject
        LEFT JOIN appointment_teacher apt ON e.id_subject = apt.id_subject
        LEFT JOIN teacher t ON apt.id_teacher = t.id_teacher
        WHERE s.fio = ?
        ORDER BY e.data
    ";

    $stmt = $connect->prepare($sql);
    $stmt->bind_param("s", $fio);
    $stmt->execute();
    $result = $stmt->get_result();
    $data = $result->fetch_all(MYSQLI_ASSOC);

    // Сохраняем данные в кэш
    if (!is_dir(__DIR__ . '/cache')) {
        mkdir(__DIR__ . '/cache', 0777, true);
    }
    file_put_contents($cacheFile, json_encode($data));
}

// Если данные не найдены
if (empty($data)) {
    writeToLog("Пользователь " . $_SESSION['user'] . " попытался сформировать зачетную книжку для студента с ФИО: $fio, но данные не найдены", 'ERROR');
    $_SESSION['message'] = "Данные для студента не найдены.";
    header("Location: ../html/record-book.php");
    exit;
}

// Создаем Excel-документ
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Заголовки
$sheet->setCellValue('B2', 'Группа:');
$sheet->setCellValue('D2', 'ФИО:');
$sheet->setCellValue('B3', 'Предмет');
$sheet->setCellValue('C3', 'Оценка');
$sheet->setCellValue('D3', 'Преподаватель');
$sheet->setCellValue('E3', 'Дата');
$sheet->getStyle('B2:E2')->applyFromArray($styleForHeader);
$sheet->getStyle('B3:E3')->applyFromArray($styleForHeader);

// Заполняем данные
$sheet->setCellValue('E2', $fio);
$sheet->setCellValue('C2', $data[0]['group_name']); // Группа из первого элемента данных

// Заполняем таблицу данными
$rowCount = 4;
foreach ($data as $row) {
    $sheet->setCellValue('B' . $rowCount, $row['subject_name']);
    $sheet->setCellValue('C' . $rowCount, $row['mark']);
    $sheet->setCellValue('D' . $rowCount, $row['teacher_name']);
    $sheet->setCellValue('E' . $rowCount, $row['data']);
    $sheet->getStyle('B' .$rowCount. ':E' .$rowCount)->applyFromArray($styleForData);
    $sheet->getStyle('C' .$rowCount)->applyFromArray($StyleForMark);
    $rowCount++;
}


    $sheet->getColumnDimension("A")->setAutoSize(true);
    $sheet->getColumnDimension("B")->setAutoSize(true);
    $sheet->getColumnDimension("C")->setAutoSize(true);
    $sheet->getColumnDimension("D")->setAutoSize(true);
    $sheet->getColumnDimension("E")->setAutoSize(true);




// Логирование успешного создания документа
writeToLog("Пользователь " . $_SESSION['user'] . " успешно сформировал зачетную книжку для студента с ФИО: $fio в формате Excel", 'INFO');

// Сохраняем документ
$fileName = $fio . '.xlsx';
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . $fileName . '"');
header('Cache-Control: max-age=0');

$writer = new Xlsx($spreadsheet);
$writer->save('php://output');
exit;