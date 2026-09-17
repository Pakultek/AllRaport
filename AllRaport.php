<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        table {
            width: 100%;
            border-collapse: collapse;
        }
        th, td {
            border: 1px solid #000;
            padding: 8px;
            text-align: center;
        }
        th {
            background-color: #f2f2f2;
        }
    </style>
</head>
<body>

<?php
require 'Uchet/vendor/autoload.php'; // Подключение PhpSpreadsheet

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$data_t = date("Y-m-d");
// Подключение к базе данных (замените на свои данные)
$servername = "local";
$username = "user";
$password = "pass";
$dbname = "db_name";

// Создание подключения
$conn = new mysqli($servername, $username, $password, $dbname);

// Проверка подключения
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

// Получаем данные из сессии (предполагается, что сессия уже запущена)
session_start();
if (!isset($_SESSION['N_jurn_edit'])) {
    die("Сессия не содержит данных о журнале.");
}
$N_jurnala = $_SESSION['N_jurn_edit'];

// Запрос для получения данных из таблицы Jurnal_raportov
$stmt = $conn->prepare("SELECT * FROM Jurnal_raportov WHERE N_jurnala = ?");
$stmt->bind_param("i", $N_jurnala);
$stmt->execute();
$result = $stmt->get_result();

$Smena = $data_t = $N_jurn = $id_zad = $id_smeni = null;
if ($data = $result->fetch_assoc()) {
    $Smena = $data['Smena'];
    $data_t = $data['Data'];
    $N_jurn = $data['N_jurnala'];
    $id_zad = $data['id_zadaniya'];
    $id_smeni = $data['id_smeni'];
}

// Инициализация массивов
$Art = [];
$Kolich = [];
$Sekcion = [];
$Model = [];
$Kol_rad = [];
$Kol_sek = [];
$vsego_rad = 0;
$vsego_sek = 0;

// Запрос для получения данных из таблиц Plan_zadaniya, Radiatori_Zakazi, Connection
for ($j = 1; $j <= 2; $j++) {
    $stmt = $conn->prepare("SELECT * FROM Plan_zadaniya, Radiatori_Zakazi, Connection 
                            WHERE Radiatori_Zakazi.N_pp = Plan_zadaniya.id_poz_zakaza 
                            AND Connection.N_connect = Radiatori_Zakazi.id_connect 
                            AND Plan_zadaniya.id_jurnala = ? AND N_svarki = ?");
    $stmt->bind_param("ii", $id_zad, $j);
    $stmt->execute();
    $qr_resul = $stmt->get_result();

    while ($data = $qr_resul->fetch_assoc()) {
        if ($data["kol_fakt"] != 0) {
            $artikul = mb_substr($data["Artikul"], 0, 8);
            $artikul = mb_substr($artikul, 4, 4);

            if ($data["Name_connect"] != '') {
                $artikul .= '-' . $data["Name_connect"];
            }

            if (!isset($Art[$j])) {
                $Art[$j] = [];
            }
            if (!in_array($artikul, $Art[$j])) {
                $Art[$j][] = $artikul;
                $Model[$j][$artikul] = $artikul;
            }

            if (!in_array($data["Sekcion"], $Sekcion)) {
                $Sekcion[] = $data["Sekcion"];
            }

            if (!isset($Kolich[$j][$artikul])) {
                $Kolich[$j][$artikul] = [];
            }
            if (isset($Kolich[$j][$artikul][$data["Sekcion"]])) {
                $Kolich[$j][$artikul][$data["Sekcion"]] += $data["kol_fakt"];
            } else {
                $Kolich[$j][$artikul][$data["Sekcion"]] = $data["kol_fakt"];
            }

            if (isset($Kol_rad[$j][$artikul])) {
                $Kol_rad[$j][$artikul] += $data["kol_fakt"];
            } else {
                $Kol_rad[$j][$artikul] = $data["kol_fakt"];
            }

            if (isset($Kol_sek[$j][$artikul])) {
                $Kol_sek[$j][$artikul] += ($data["kol_fakt"] * $data["Sekcion"]);
            } else {
                $Kol_sek[$j][$artikul] = $data["kol_fakt"] * $data["Sekcion"];
            }
        }
    }
}

// Сортируем массив секционности по возрастанию
sort($Sekcion);

// Вывод данных в таблицу
echo "<table>";
echo "<thead>
        <tr>
            <th rowspan='2'>Наименование операции</th>
            <th rowspan='2'>Модель</th>
            <th colspan='2'>Количество</th>
            <th rowspan='2'>% брака</th>
            <th colspan='" . (count($Sekcion)) . "'>Секционность</th>
        </tr>
        <tr>
            <th>Радиаторов</th>
            <th>Секций</th>";
foreach ($Sekcion as $s) {
    if ($s != 0) {
        echo "<th>$s</th>";
    }
}
echo "</tr>
      </thead>";
echo "<tbody>";
echo "<tr>";
echo "<td rowspan='" . (count($Model[1]) + count($Model[2]) + 1) . "'>Сварка радиаторов</td>";

foreach ($Art as $j => $models) {
    foreach ($models as $artikul) {
        echo "<tr>";
        echo "<td>{$Model[$j][$artikul]}</td>";

        // Вывод общего количества радиаторов и секций
        echo "<td>" . (isset($Kol_rad[$j][$artikul]) ? $Kol_rad[$j][$artikul] : 0) . "</td>";
        echo "<td>" . (isset($Kol_sek[$j][$artikul]) ? $Kol_sek[$j][$artikul] : 0) . "</td>";
        echo "<td></td>";

        // Вывод секционности
        foreach ($Sekcion as $s) {
            if ($s != 0) {
                echo "<td>" . (isset($Kolich[$j][$artikul][$s]) ? $Kolich[$j][$artikul][$s] : 0) . "</td>";
            }
        }

        $vsego_rad += isset($Kol_rad[$j][$artikul]) ? $Kol_rad[$j][$artikul] : 0;
        $vsego_sek += isset($Kol_sek[$j][$artikul]) ? $Kol_sek[$j][$artikul] : 0;
        echo "</tr>";
    }
}

echo "<tr>
		<th rowspan='2'></th>
        <th>Брак:</th>
        <th></th>
        <th></th>
        <th rowspan='2'></th>";

        echo "<th rowspan='2' colspan='" . (count($Sekcion)) . "'></th>";
   

echo "</tr>";
		
echo "<tr>
        <th>ИТОГО:</th>
        <th>$vsego_rad</th>
        <th>$vsego_sek</th>";
		
echo "</tr>";

echo "</tbody></table>";

// Добавляем кнопку для скачивания
echo "<br><a href='?export=excel'>Скачать в Excel</a>";

// Экспорт в Excel
if (isset($_GET['export']) && $_GET['export'] == 'excel') {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Заголовки
    $sheet->setCellValue('A1', 'Наименование операции');
    $sheet->setCellValue('B1', 'Модель');
    $sheet->setCellValue('C1', 'Количество');
    $sheet->setCellValue('D1', '% брака');

    $col = 'E';
    foreach ($Sekcion as $s) {
        if ($s != 0) {
            $sheet->setCellValue($col . '1', $s);
            $col++;
        }
    }

    // Данные
    $row = 2;
    foreach ($Art as $j => $models) {
        foreach ($models as $artikul) {
            $sheet->setCellValue('A' . $row, 'Сварка радиаторов');
            $sheet->setCellValue('B' . $row, $Model[$j][$artikul]);
            $sheet->setCellValue('C' . $row, isset($Kol_rad[$j][$artikul]) ? $Kol_rad[$j][$artikul] : 0);
            $sheet->setCellValue('D' . $row, '');

            $col = 'E';
            foreach ($Sekcion as $s) {
                if ($s != 0) {
                    $sheet->setCellValue($col . $row, isset($Kolich[$j][$artikul][$s]) ? $Kolich[$j][$artikul][$s] : 0);
                    $col++;
                }
            }
            $row++;
        }
    }

    // Итого
    $sheet->setCellValue('B' . $row, 'ИТОГО:');
    $sheet->setCellValue('C' . $row, $vsego_rad);
    $sheet->setCellValue('D' . $row, $vsego_sek);

    // Сохранение файла
    $writer = new Xlsx($spreadsheet);
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="raport.xlsx"');
    header('Cache-Control: max-age=0');
    $writer->save('php://output');
    exit;
}

// Закрываем соединение
$conn->close();
?>

</body>
</html>
