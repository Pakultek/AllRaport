<!DOCTYPE html>
<html lang="ru">
<head>
    <title>Суточный рапорт</title>
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
        #back {
            display: inline-block;
            color: white;
            text-decoration: none;
            padding: .2em 2em;
            outline: none;
            border-width: 2px 0;
            border-style: solid none;
            border-color: #FDBE33 #000 #D77206;
            border-radius: 6px;
            background: linear-gradient(#F3AE0F, #E38916) #E38916;
            width: 170px;
            height: 28px;
        }
    </style>
</head>
<body>

<!-- Форма для выбора даты -->
<button id="back" style="width: 150px; height: 30px;">Обратно</button>
<p></p>
<form method="GET" action="">
    <label for="date" style="font-size:15px;">Выберите дату:</label>
    <input type="date" id="datePicker" name="date" max="<?php echo date('Y-m-d'); ?>">
    <button type="submit">Показать данные</button>
</form>

<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

// Подключение к базе данных (замените на свои данные)
$servername = "10.10.3.41";
$username = "vadim";
$password = "1q2w3e4r5t6y";
$dbname = "TUBOG";

// Создание подключения
$conn = new mysqli($servername, $username, $password, $dbname);

// Проверка подключения
if ($conn->connect_error) 
{
    die("Connection failed: " . $conn->connect_error);
}

// Запуск сессии
session_start();

// Проверка
if (isset($_GET['date'])) 
{
    $date = $_GET['date'];

    // Кнопка для скачивания документа
    echo "<br><br>";
    echo "<form action='AllRaport.php' method='post'>";
    echo "<input type='hidden' name='date' value='" . htmlspecialchars($date) . "'>";
    echo "<button type='submit' name='download_excel'>Скачать Excel-документ</button>";
    echo "</form>";
    echo "<br>";

    // Запрос для получения данных из таблиц Jurnal_planov, Plan_zadaniya, Radiatori_Zakazi, Connection
    // Сварка
	$sql = "SELECT * 
            FROM Jurnal_planov
            JOIN Plan_zadaniya ON Plan_zadaniya.id_jurnala = Jurnal_planov.N_por
            JOIN Radiatori_Zakazi ON Radiatori_Zakazi.N_pp = Plan_zadaniya.id_poz_zakaza
            JOIN Connection ON Connection.N_connect = Radiatori_Zakazi.id_connect
            WHERE DATE(Jurnal_planov.Data) = ?";
    $stmt = $conn->prepare($sql);
    $stmt->bind_param("s", $date);
    $stmt->execute();
    $result = $stmt->get_result();

	// Первичные испытания
    $sqlIsp = "SELECT * 
        FROM Pervichnoe_ispit
        JOIN Plan_zadaniya ON Plan_zadaniya.N_poz = Pervichnoe_ispit.Pozic_Zad
        JOIN Jurnal_smen ON Jurnal_smen.N_smeni = Pervichnoe_ispit.id_smeni
        JOIN Radiatori_Zakazi ON Radiatori_Zakazi.N_pp = Plan_zadaniya.id_poz_zakaza
        JOIN Connection ON Connection.N_connect = Radiatori_Zakazi.id_connect
        WHERE DATE(Jurnal_smen.Data) = ?";
    
    $stmtIsp = $conn->prepare($sqlIsp);
    $stmtIsp->bind_param("s", $date);
    $stmtIsp->execute();
    $resultIsp = $stmtIsp->get_result();

	//Окраска
    $sqlZav = "SELECT * 
        FROM Zaveska
        JOIN Jurnal_raportov ON Jurnal_raportov.N_jurnala = Zaveska.id_jurnala
        JOIN Plan_zadaniya ON Plan_zadaniya.N_poz = Zaveska.Pozic_Zad
        JOIN Radiatori_Zakazi ON Radiatori_Zakazi.N_pp = Plan_zadaniya.id_poz_zakaza
        JOIN Connection ON Connection.N_connect = Radiatori_Zakazi.id_connect
        WHERE DATE(Zaveska.Data) = ?";

    $stmtZav = $conn->prepare($sqlZav);
    $stmtZav->bind_param("s", $date);
    $stmtZav->execute();
    $resultZav = $stmtZav->get_result();
	
	// Вторичные испытания
    $sqlIsp2 = "SELECT * 
        FROM Vtorichnoe_ispit
        JOIN Plan_zadaniya ON Plan_zadaniya.N_poz = Vtorichnoe_ispit.Pozic_Zad
        JOIN Jurnal_smen ON Jurnal_smen.N_smeni = Vtorichnoe_ispit.id_smeni
        JOIN Radiatori_Zakazi ON Radiatori_Zakazi.N_pp = Plan_zadaniya.id_poz_zakaza
        JOIN Connection ON Connection.N_connect = Radiatori_Zakazi.id_connect
        WHERE DATE(Jurnal_smen.Data) = ?";
    
    $stmtIsp2 = $conn->prepare($sqlIsp2);
    $stmtIsp2->bind_param("s", $date);
    $stmtIsp2->execute();
    $resultIsp2 = $stmtIsp2->get_result();

    // Упаковка радиаторов
    $sqlUp = "SELECT * 
        FROM `Gotovaya_produc`
        JOIN `Radiatori_Zakazi` ON `Radiatori_Zakazi`.`N_pp` = `Gotovaya_produc`.`id_zakaza`
        JOIN `Connection` ON `Connection`.`N_connect` = `Radiatori_Zakazi`.`id_connect`
        JOIN `Color` ON `Color`.`N_color` = `Radiatori_Zakazi`.`id_color`
        WHERE DATE(`Gotovaya_produc`.`Data_upakovki`) = ?";
        
    $stmtUp = $conn->prepare($sqlUp);
    $stmtUp->bind_param("s", $date);
    $stmtUp->execute();
    $resultUp = $stmtUp->get_result();

    // Инициализация массивов
	//Сварка
    $Art = [];
    $Kolich = [];
    $Sekcion = [];
    $Model = [];
    $Kol_rad = [];
    $Kol_sek = [];
    $vsego_rad = 0;
    $vsego_sek = 0;
    
	//Первичные испытания
    $ArtIsp = [];
    $KolichIsp = [];
    $SekcionIsp = [];
    $ModelIsp = [];
    $Kol_radIsp = [];
    $Kol_sekIsp = [];
    $vsego_radIsp = 0;
    $vsego_sekIsp = 0;
    $vsego_rad_brak = 0;
    $vsego_rad_remont = 0;
    $vsego_sek_brak = 0;
    $vsego_sek_remont = 0;
    $Proc_brak = 0;
    $Brak = 0;
    $Kol_rad_brak = [];
    $Kol_sek_brak = [];
    $Kol_rad_remont = [];
    $Kol_sek_remont = [];
    $Kolich_brak = [];
    $Kolich_remont = [];
	
	//Вторичные испытания
	$ArtIsp2 = [];
    $KolichIsp2 = [];
    $SekcionIsp2 = [];
    $ModelIsp2 = [];
    $Kol_radIsp2 = [];
    $Kol_sekIsp2 = [];
    $vsego_radIsp2 = 0;
    $vsego_sekIsp2 = 0;
    $vsego_rad_brak2 = 0;
    $vsego_rad_remont2 = 0;
    $vsego_sek_brak2 = 0;
    $vsego_sek_remont2 = 0;
    $Proc_brak2 = 0;
    $Brak2 = 0;
    $Kol_rad_brak2 = [];
    $Kol_sek_brak2 = [];
    $Kol_rad_remont2 = [];
    $Kol_sek_remont2 = [];
    $Kolich_brak2 = [];
    $Kolich_remont2 = [];
    
	//Окраска
    $ArtZav = [];
    $KolichZav = [];
    $SekcionZav = [];
    $ModelZav = [];
    $Kol_radZav = [];
    $Kol_sekZav = [];
    $vsego_radZav = 0;
    $vsego_sekZav = 0;
    
    // Упаковка радиаторов
    $ArtUp = [];
    $KolichUp = [];
    $SekcionUp = [];
    $ModelUp = [];
    $Kol_radUp = [];
    $Kol_sekUp = [];
    $vsego_radUp = 0;
    $vsego_sekUp = 0;

    // Обработка данных
	//Сварка
    while ($row = $result->fetch_assoc()) 
	{
        if ($row["kol_fakt"] != 0) 
		{
            $artikul = mb_substr($row["Artikul"], 0, 8);
            $artikul = mb_substr($artikul, 4, 4);

            if ($row["Name_connect"] != '') 
			{
                $artikul .= '-' . $row["Name_connect"];
            }

            if (!in_array($artikul, $Art)) 
			{
                $Art[] = $artikul;
                $Model[$artikul] = $artikul;
            }

            if (!in_array($row["Sekcion"], $Sekcion)) 
			{
                $Sekcion[] = $row["Sekcion"];
            }

            if (!isset($Kolich[$artikul])) 
			{
                $Kolich[$artikul] = [];
            }
            if (isset($Kolich[$artikul][$row["Sekcion"]])) 
			{
                $Kolich[$artikul][$row["Sekcion"]] += $row["kol_fakt"];
            } 
			else 
			{
                $Kolich[$artikul][$row["Sekcion"]] = $row["kol_fakt"];
            }

            if (isset($Kol_rad[$artikul])) 
			{
                $Kol_rad[$artikul] += $row["kol_fakt"];
            } 
			else 
			{
                $Kol_rad[$artikul] = $row["kol_fakt"];
            }

            if (isset($Kol_sek[$artikul])) 
			{
                $Kol_sek[$artikul] += ($row["kol_fakt"] * $row["Sekcion"]);
            } 
			else 
			{
                $Kol_sek[$artikul] = $row["kol_fakt"] * $row["Sekcion"];
            }
        }
    }

	//Первичные испытания
    while ($rowIsp = $resultIsp->fetch_assoc()) 
	{
        $id_poz_zakaza = $rowIsp["id_poz_zakaza"];
        if ($id_poz_zakaza != 0) 
		{
            $artikulIsp = mb_substr($rowIsp["Artikul"], 0, 8);
            $artikulIsp = mb_substr($artikulIsp, 4, 4);
                    
            if ($rowIsp["Name_connect"] != '') 
			{
                $artikulIsp .= '-' . $rowIsp["Name_connect"];
            }
                        
            if (!in_array($artikulIsp, $ArtIsp)) 
			{
                $ArtIsp[] = $artikulIsp;
                $ModelIsp[$artikulIsp] = $artikulIsp;
            }
                    
            if (!in_array($rowIsp["Sekcion"], $SekcionIsp)) 
			{
                $SekcionIsp[] = $rowIsp["Sekcion"];
            }
                    
            if ($rowIsp["Status_isp"] == 1) 
			{
                if (isset($KolichIsp[$artikulIsp][$rowIsp["Sekcion"]])) 
				{
                    $KolichIsp[$artikulIsp][$rowIsp["Sekcion"]]++;
                } 
				else 
				{
                    $KolichIsp[$artikulIsp][$rowIsp["Sekcion"]] = 1;
                }
                    
                if (isset($Kol_radIsp[$artikulIsp])) 
				{
                    $Kol_radIsp[$artikulIsp]++;
                } 
				else 
				{
                    $Kol_radIsp[$artikulIsp] = 1;
                }
                        
                if (isset($Kol_sekIsp[$artikulIsp])) 
				{
                    $Kol_sekIsp[$artikulIsp] += $rowIsp["Sekcion"];
                } 
				else 
				{
                    $Kol_sekIsp[$artikulIsp] = $rowIsp["Sekcion"];
                }
            } 
			else 
			{
                if ($rowIsp["Status_isp"] == 2) 
				{
                    if (isset($Kolich_remont[$artikulIsp][$rowIsp["Sekcion"]])) 
					{
                        $Kolich_remont[$artikulIsp][$rowIsp["Sekcion"]]++;
                    } 
					else 
					{
                        $Kolich_remont[$artikulIsp][$rowIsp["Sekcion"]] = 1;
                    }
                            
                    if (isset($Kol_rad_remont[$artikulIsp])) 
					{
                        $Kol_rad_remont[$artikulIsp]++;
                    } 
					else 
					{
                        $Kol_rad_remont[$artikulIsp] = 1;
                    }
                            
                    if (isset($Kol_sek_remont[$artikulIsp])) 
					{
                        $Kol_sek_remont[$artikulIsp] += $rowIsp["Sekcion"];
                    } 
					else 
					{
                        $Kol_sek_remont[$artikulIsp] = $rowIsp["Sekcion"];
                    }
                } 
				else 
				{
                    if (isset($Kolich_brak[$artikulIsp][$rowIsp["Sekcion"]])) 
					{
                        $Kolich_brak[$artikulIsp][$rowIsp["Sekcion"]]++;
                    } 
					else 
					{
                        $Kolich_brak[$artikulIsp][$rowIsp["Sekcion"]] = 1;
                    }
                            
                    if (isset($Kol_rad_brak[$artikulIsp])) 
					{
                        $Kol_rad_brak[$artikulIsp]++;
                    } 
					else 
					{
                        $Kol_rad_brak[$artikulIsp] = 1;
                    }
                            
                    if (isset($Kol_sek_brak[$artikulIsp])) 
					{
                        $Kol_sek_brak[$artikulIsp] += $rowIsp["Sekcion"];
                    } 
					else 
					{
                        $Kol_sek_brak[$artikulIsp] = $rowIsp["Sekcion"];
                    }
                }
            }	
        }
    }
	
	//Вторичные испытания
	while ($rowIsp2 = $resultIsp2->fetch_assoc()) 
	{
        $id_poz_zakaza = $rowIsp2["id_poz_zakaza"];
        if ($id_poz_zakaza != 0) 
		{
            $artikulIsp2 = mb_substr($rowIsp2["Artikul"], 0, 8);
            $artikulIsp2 = mb_substr($artikulIsp2, 4, 4);
                    
            if ($rowIsp2["Name_connect"] != '') 
			{
                $artikulIsp2 .= '-' . $rowIsp2["Name_connect"];
            }
                        
            if (!in_array($artikulIsp2, $ArtIsp2)) 
			{
                $ArtIsp2[] = $artikulIsp2;
                $ModelIsp2[$artikulIsp2] = $artikulIsp2;
            }
                    
            if (!in_array($rowIsp2["Sekcion"], $SekcionIsp2)) 
			{
                $SekcionIsp2[] = $rowIsp2["Sekcion"];
            }
                    
            if ($rowIsp2["Status_isp"] == 1) 
			{
                if (isset($KolichIsp2[$artikulIsp2][$rowIsp2["Sekcion"]])) 
				{
                    $KolichIsp2[$artikulIsp2][$rowIsp2["Sekcion"]]++;
                } 
				else 
				{
                    $KolichIsp2[$artikulIsp2][$rowIsp2["Sekcion"]] = 1;
                }
                    
                if (isset($Kol_radIsp2[$artikulIsp2])) 
				{
                    $Kol_radIsp2[$artikulIsp2]++;
                } 
				else 
				{
                    $Kol_radIsp2[$artikulIsp2] = 1;
                }
                        
                if (isset($Kol_sekIsp2[$artikulIsp2])) 
				{
                    $Kol_sekIsp2[$artikulIsp2] += $rowIsp2["Sekcion"];
                } 
				else 
				{
                    $Kol_sekIsp2[$artikulIsp2] = $rowIsp2["Sekcion"];
                }
            } 
			else 
			{
                if ($rowIsp2["Status_isp"] == 2) 
				{
                    if (isset($Kolich_remont2[$artikulIsp2][$rowIsp2["Sekcion"]])) 
					{
                        $Kolich_remont2[$artikulIsp2][$rowIsp2["Sekcion"]]++;
                    } 
					else 
					{
                        $Kolich_remont2[$artikulIsp2][$rowIsp2["Sekcion"]] = 1;
                    }
                            
                    if (isset($Kol_rad_remont2[$artikulIsp2])) 
					{
                        $Kol_rad_remont2[$artikulIsp2]++;
                    } 
					else 
					{
                        $Kol_rad_remont2[$artikulIsp2] = 1;
                    }
                            
                    if (isset($Kol_sek_remont2[$artikulIsp2])) 
					{
                        $Kol_sek_remont2[$artikulIsp2] += $rowIsp2["Sekcion"];
                    } 
					else 
					{
                        $Kol_sek_remont2[$artikulIsp2] = $rowIsp2["Sekcion"];
                    }
                } 
				else 
				{
                    if (isset($Kolich_brak2[$artikulIsp2][$rowIsp2["Sekcion"]])) 
					{
                        $Kolich_brak2[$artikulIsp2][$rowIsp2["Sekcion"]]++;
                    } 
					else 
					{
                        $Kolich_brak2[$artikulIsp2][$rowIsp2["Sekcion"]] = 1;
                    }
                            
                    if (isset($Kol_rad_brak2[$artikulIsp2])) 
					{
                        $Kol_rad_brak2[$artikulIsp2]++;
                    } 
					else 
					{
                        $Kol_rad_brak2[$artikulIsp2] = 1;
                    }
                            
                    if (isset($Kol_sek_brak2[$artikulIsp2])) 
					{
                        $Kol_sek_brak2[$artikulIsp2] += $rowIsp2["Sekcion"];
                    } 
					else 
					{
                        $Kol_sek_brak2[$artikulIsp2] = $rowIsp2["Sekcion"];
                    }
                }
            }	
        }
    }
	
	//Окраска
    while ($rowZav = $resultZav->fetch_assoc()) 
	{
        $id_poz_zakazaZav = $rowZav["Data"];
        if ($id_poz_zakazaZav != 0) 
		{
            $key = array_search($rowZav["Sekcion"], $SekcionZav);
            if ($key === false) 
            {
                $SekcionZav[] = $rowZav["Sekcion"];
            }

            $articul_index = '';
            $temp = '';
            $temp2 = '';
            $temp = mb_substr($rowZav["Artikul"], 0, 8);
            $temp2 = mb_substr($rowZav["Artikul"], 11);
            $articul_index = $temp . '' . $temp2;

            $key = array_search($articul_index, $ArtZav);
            if ($key === false) 
            {
                $ArtZav[] = $articul_index;
            }

            // Определение раздела
            $section = '';
            if ($rowZav["id_color"] == 0) 
            {
                $section = "Белый";
            } 
            else if ($rowZav["id_color"] == 6) 
            {
                $section = "RAL";
            }
            else 
            {
                $section = "Цветные";
            }

            if (isset($KolichZav[$articul_index][$section][$rowZav["Sekcion"]])) 
            {
                $KolichZav[$articul_index][$section][$rowZav["Sekcion"]]++;
            }
            else 
            {
                $KolichZav[$articul_index][$section][$rowZav["Sekcion"]] = 1;
            }
        }
    }

    // Упаковка радиаторов
    while ($rowUp = $resultUp->fetch_assoc()) 
	{
        $artikulUp = mb_substr($rowUp["Artikul"], 0, 8);
        $artikulUp = mb_substr($artikulUp, 4, 4);

        if ($rowUp["Name_connect"] != '') 
		{
            $artikulUp .= '-' . $rowUp["Name_connect"];
        }

        if ($rowUp["id_color"] != 0) 
		{
            $artikulUp .= '-' . $rowUp["Kr_name"];
        }
		else if ($rowUp["id_color"] == 6)
		{
			$artikulUp .= '-' . $rowUp["N_RAL"];
		}

        if (!in_array($artikulUp, $ArtUp)) 
		{
            $ArtUp[] = $artikulUp;
            $ModelUp[$artikulUp] = $artikulUp;
        }

        if (!in_array($rowUp["Sekcion"], $SekcionUp)) 
		{
            $SekcionUp[] = $rowUp["Sekcion"];
        }

        if (!isset($KolichUp[$artikulUp])) 
		{
            $KolichUp[$artikulUp] = [];
        }
        if (isset($KolichUp[$artikulUp][$rowUp["Sekcion"]])) 
		{
            $KolichUp[$artikulUp][$rowUp["Sekcion"]] += 1; 
        } 
		else 
		{
            $KolichUp[$artikulUp][$rowUp["Sekcion"]] = 1; 
        }

        if (isset($Kol_radUp[$artikulUp])) 
		{
            $Kol_radUp[$artikulUp] += 1; 
        }
		else
		{
            $Kol_radUp[$artikulUp] = 1; 
        }

        if (isset($Kol_sekUp[$artikulUp])) 
		{
            $Kol_sekUp[$artikulUp] += $rowUp["Sekcion"]; 
        } 
		else 
		{
            $Kol_sekUp[$artikulUp] = $rowUp["Sekcion"]; 
        }
    }

    // Сортируем массив секционности по возрастанию
    sort($Sekcion);
    sort($SekcionIsp);
    sort($SekcionIsp2);
    sort($SekcionZav);
    sort($SekcionUp);
    

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
    foreach ($Sekcion as $s) 
	{
        if ($s != 0) 
		{
            echo "<th>$s</th>";
        }
    }
    echo "</tr>
          </thead>";
    echo "<tbody>";
    echo "<tr>";
    echo "<td rowspan='".(count($Model) + 1 )."'>Сварка радиаторов</td>";

    foreach ($Art as $artikul) 
    {
        echo "<tr>";
        echo "<td>{$Model[$artikul]}</td>"; 

        // Вывод общего количества радиаторов и секций
        echo "<td>" . (isset($Kol_rad[$artikul]) ? $Kol_rad[$artikul] : 0) . "</td>";
        echo "<td>" . (isset($Kol_sek[$artikul]) ? $Kol_sek[$artikul] : 0) . "</td>";
        echo "<td></td>";
        
        // Вывод секционности
        foreach ($Sekcion as $s) 
		{
            if ($s != 0) 
			{
                echo "<td>" . (isset($Kolich[$artikul][$s]) ? $Kolich[$artikul][$s] : 0) . "</td>";
            }
        }

        $vsego_rad += isset($Kol_rad[$artikul]) ? $Kol_rad[$artikul] : 0;
        $vsego_sek += isset($Kol_sek[$artikul]) ? $Kol_sek[$artikul] : 0;
        
        
        echo "</tr>";
    }
    
    
    
    echo "<tr>
            <th rowspan='2'></th>
            <th>Окончат. брак:</th>
            <th></th>
            <th></th>
            <th rowspan='2'></th>";

    echo "<th rowspan='2' colspan='" . (count($Sekcion)) . "'></th>";

    echo "</tr>";

    echo "<tr>
            <th>Годные:</th>
            <th>$vsego_rad</th>
            <th>$vsego_sek</th>";

    echo "</tr>";
    
    echo "<th></th>";
    
    echo "<thead>
            <tr>
                <th rowspan='2'>Наименование операции</th>
                <th rowspan='2'>Модель</th>
                <th colspan='2'>Количество</th>
                <th rowspan='2'>% брака</th>
                <th colspan='" . (count($SekcionIsp)) . "'>Секционность</th>
            </tr>
            <tr>
                <th>Радиаторов</th>
                <th>Секций</th>";
    foreach ($SekcionIsp as $s) 
	{
        if ($s != 0) 
		{
            echo "<th>$s</th>";
        }
    }
    echo "</tr>
          </thead>";
    echo "<tr>";
    
    echo "<td rowspan='".(count($ModelIsp) + 1 )."'>Первичные испытания радиаторов</td>";

    foreach ($ArtIsp as $artikulIsp) 
    {
        echo "<tr>";
        echo "<td>{$ModelIsp[$artikulIsp]}</td>"; 

        // Вывод общего количества радиаторов и секций
        echo "<td>" . (isset($Kol_radIsp[$artikulIsp]) ? $Kol_radIsp[$artikulIsp] : 0) . "</td>";
        echo "<td>" . (isset($Kol_sekIsp[$artikulIsp]) ? $Kol_sekIsp[$artikulIsp] : 0) . "</td>";
        echo "<td></td>";
        
        // Вывод секционности
        foreach ($SekcionIsp as $sIsp) 
		{
            if ($sIsp != 0) 
			{
                echo "<td>" . (isset($KolichIsp[$artikulIsp][$sIsp]) ? $KolichIsp[$artikulIsp][$sIsp] : 0) . "</td>";
            }
        }

        $vsego_radIsp += isset($Kol_radIsp[$artikulIsp]) ? $Kol_radIsp[$artikulIsp] : 0;
        $vsego_sekIsp += isset($Kol_sekIsp[$artikulIsp]) ? $Kol_sekIsp[$artikulIsp] : 0;
        
        $vsego_rad_brak +=isset($Kol_rad_brak[$artikulIsp]) ? $Kol_rad_brak[$artikulIsp] : 0;
        $vsego_sek_brak +=isset($Kol_sek_brak[$artikulIsp]) ? $Kol_sek_brak[$artikulIsp] : 0;
        
        $vsego_rad_remont += isset($Kol_rad_remont[$artikulIsp]) ? $Kol_rad_remont[$artikulIsp] : 0;
        $vsego_sek_remont += isset($Kol_sek_remont[$artikulIsp]) ? $Kol_sek_remont[$artikulIsp] : 0;
        echo "</tr>";
    }
    $Proc_brak = ($vsego_rad_brak * 100) / $vsego_radIsp;
    $Brak = round($Proc_brak, 2);
    
    echo "</tr>";
    

    echo "<tr>
            <th rowspan='3'></th>
            <th>Окончат. брак:</th>
            <th>$vsego_rad_brak</th>
            <th>$vsego_sek_brak</th>
            <th>$Brak</th>";

    echo "<th rowspan='3' colspan='" . (count($SekcionIsp)) . "'></th>";

    echo "</tr>";

    echo "<tr>
            <th>Ремонт:</th>
            <th>$vsego_rad_remont</th>
            <th>$vsego_sek_remont</th>
            <th rowspan='2'></th>
        </tr>";
    
    echo "<tr>
            
            <th>Годные:</th>
            <th>$vsego_radIsp</th>
            <th>$vsego_sekIsp</th>";
    echo "</tr>";
	
	echo "<th></th>";
    
    echo "<thead>
            <tr>
                <th rowspan='2'>Наименование операции</th>
                <th rowspan='2'>Модель</th>
                <th colspan='2'>Количество</th>
                <th rowspan='2'>% брака</th>
                <th colspan='" . (count($SekcionIsp2)) . "'>Секционность</th>
            </tr>
            <tr>
                <th>Радиаторов</th>
                <th>Секций</th>";
    foreach ($SekcionIsp2 as $s) 
	{
        if ($s != 0) 
		{
            echo "<th>$s</th>";
        }
    }
    echo "</tr>
          </thead>";
    echo "<tr>";
    
    echo "<td rowspan='".(count($ModelIsp2) + 1 )."'>Вторичные испытания радиаторов</td>";

    foreach ($ArtIsp2 as $artikulIsp2) 
    {
        echo "<tr>";
        echo "<td>{$ModelIsp2[$artikulIsp2]}</td>"; 

        // Вывод общего количества радиаторов и секций
        echo "<td>" . (isset($Kol_radIsp2[$artikulIsp2]) ? $Kol_radIsp2[$artikulIsp2] : 0) . "</td>";
        echo "<td>" . (isset($Kol_sekIsp2[$artikulIsp2]) ? $Kol_sekIsp2[$artikulIsp2] : 0) . "</td>";
        echo "<td></td>";
        
        // Вывод секционности
        foreach ($SekcionIsp2 as $sIsp2) 
		{
            if ($sIsp2 != 0) 
			{
                echo "<td>" . (isset($KolichIsp2[$artikulIsp2][$sIsp2]) ? $KolichIsp2[$artikulIsp2][$sIsp2] : 0) . "</td>";
            }
        }

        $vsego_radIsp2 += isset($Kol_radIsp2[$artikulIsp2]) ? $Kol_radIsp2[$artikulIsp2] : 0;
        $vsego_sekIsp2 += isset($Kol_sekIsp2[$artikulIsp2]) ? $Kol_sekIsp2[$artikulIsp2] : 0;
        
        $vsego_rad_brak2 += isset($Kol_rad_brak2[$artikulIsp2]) ? $Kol_rad_brak2[$artikulIsp2] : 0;
        $vsego_sek_brak2 += isset($Kol_sek_brak2[$artikulIsp2]) ? $Kol_sek_brak2[$artikulIsp2] : 0;
        
        $vsego_rad_remont2 += isset($Kol_rad_remont2[$artikulIsp2]) ? $Kol_rad_remont2[$artikulIsp2] : 0;
        $vsego_sek_remont2 += isset($Kol_sek_remont2[$artikulIsp2]) ? $Kol_sek_remont2[$artikulIsp2] : 0;
        echo "</tr>";
    }
    $Proc_brak2 = ($vsego_rad_brak2 * 100) / $vsego_radIsp2;
    $Brak2 = round($Proc_brak2, 2);
    
    echo "</tr>";
    

    echo "<tr>
            <th rowspan='3'></th>
            <th>Окончат. брак:</th>
            <th>$vsego_rad_brak2</th>
            <th>$vsego_sek_brak2</th>
            <th>$Brak2</th>";

    echo "<th rowspan='3' colspan='" . (count($SekcionIsp2)) . "'></th>";

    echo "</tr>";

    echo "<tr>
            <th>Ремонт:</th>
            <th>$vsego_rad_remont2</th>
            <th>$vsego_sek_remont2</th>
            <th rowspan='2'></th>
        </tr>";
    
    echo "<tr>
            
            <th>Годные:</th>
            <th>$vsego_radIsp2</th>
            <th>$vsego_sekIsp2</th>";
    echo "</tr>";
    
    echo "<th></th>";
    
    echo "<thead>
            <tr>
                <th rowspan='2'>Наименование операции</th>
                <th rowspan='2'>Модель</th>
                <th colspan='2'>Количество</th>
                <th rowspan='2'>% брака</th>
                <th colspan='" . (count($SekcionZav)) . "'>Секционность</th>
            </tr>
            <tr>
                <th>Радиаторов</th>
                <th>Секций</th>";
    foreach ($SekcionZav as $s) 
	{
        if ($s != 0) 
		{
            echo "<th>$s</th>";
        }
    }
    echo "</tr>
          </thead>";
    echo "<tr>";
    echo "<td rowspan='" . (count($ArtZav) + 4) . "'>Окраска радиаторов</td>";

    $sections = ["Белый", "Цветные", "RAL"]; // Разделы для окраски
    $vsego_radZav = 0;
    $vsego_sekZav = 0;

    foreach ($sections as $section) 
	{
        echo "<tr>";
        echo "<td colspan='4'></td>";
        echo "<td colspan='" . (count($SekcionZav)) . "'>$section</td>";
        echo "</tr>";

        foreach ($ArtZav as $articul_index) 
		{
            $kol_radZav = 0;
            $kol_sekZav = 0;
            $flag = false;

            foreach ($SekcionZav as $s) 
			{
                if (isset($KolichZav[$articul_index][$section][$s])) 
				{
                    $flag = true;
                    break;
                }
            }

            if ($flag) 
			{
                echo "<tr>";
                echo "<td>$articul_index</td>";
                
                // Вывод общего количества радиаторов и секций
                foreach ($SekcionZav as $s) 
				{
                    if ($s != 0) 
					{
                        $kol_radZav += isset($KolichZav[$articul_index][$section][$s]) ? $KolichZav[$articul_index][$section][$s] : 0;
                        $kol_sekZav += isset($KolichZav[$articul_index][$section][$s]) ? $KolichZav[$articul_index][$section][$s] * $s : 0;
                    }
                }

                echo "<td>$kol_radZav</td>";
                echo "<td>$kol_sekZav</td>";
                echo "<td></td>";
                
                foreach ($SekcionZav as $s) 
				{
                    if ($s != 0) 
					{
                        echo "<td>" . (isset($KolichZav[$articul_index][$section][$s]) ? $KolichZav[$articul_index][$section][$s] : 0) . "</td>";
                    }
                }

                $vsego_radZav += $kol_radZav;
                $vsego_sekZav += $kol_sekZav;

                echo "</tr>";
            }
        }
    }

    echo "<tr>
            <th rowspan='2'></th>
            <th>Снято:</th>
            <th></th>
            <th></th>";
            
    echo "<th rowspan='2' colspan='" . (count($SekcionZav) + 1) . "'></th>
          </tr>";

    echo "<tr>
            <th>Годные:</th>
            <th>$vsego_radZav</th>
            <th>$vsego_sekZav</th>";
    echo "</tr>";

    echo"<th></th>";		  
              
              
    // Вывод данных в таблицу для упаковки радиаторов
   
    echo "<thead>
        <tr>
            <th rowspan='2'>Наименование операции</th>
            <th rowspan='2'>Модель</th>
            <th colspan='2'>Количество</th>
            <th rowspan='2'>% брака</th>
            <th colspan='" . (count($SekcionUp)) . "'>Секционность</th>
        </tr>
        <tr>
            <th>Радиаторов</th>
            <th>Секций</th>";
foreach ($SekcionUp as $s) 
{
    if ($s != 0) 
	{
        echo "<th>$s</th>";
    }
}
echo "</tr>
      </thead>";
echo "<tbody>";
echo "<tr>";
echo "<td rowspan='" . (count($ModelUp) + 1) . "'>Упаковка радиаторов</td>";

foreach ($ArtUp as $articulUp) 
{
    $kol_radUp = 0;
    $kol_sekUp = 0;

    echo "<tr>";
    echo "<td>{$ModelUp[$articulUp]}</td>";

    // Вывод общего количества радиаторов и секций
    foreach ($SekcionUp as $s) 
	{
        if ($s != 0) 
		{
            $kol_radUp += isset($KolichUp[$articulUp][$s]) ? $KolichUp[$articulUp][$s] : 0;
            $kol_sekUp += isset($KolichUp[$articulUp][$s]) ? $KolichUp[$articulUp][$s] * $s : 0;
        }
    }

    echo "<td>$kol_radUp</td>";
    echo "<td>$kol_sekUp</td>";
    echo "<td></td>";

    foreach ($SekcionUp as $s) 
	{
        if ($s != 0) 
		{
            echo "<td>" . (isset($KolichUp[$articulUp][$s]) ? $KolichUp[$articulUp][$s] : 0) . "</td>";
        }
    }

    $vsego_radUp += $kol_radUp;
    $vsego_sekUp += $kol_sekUp;

    echo "</tr>";
}

echo "<tr>
        <th rowspan='2'></th>
        <th>Годные:</th>
        <th>$vsego_radUp</th>
        <th>$vsego_sekUp</th>";
echo "</tr>";

echo "</tbody></table>";
		
}


// Закрываем соединение
$conn->close();
?>

<script>
// Вызов функции после загрузки страницы
window.onload = function() 
{
    const urlParams = new URLSearchParams(window.location.search);
    const selectedDate = urlParams.get('date');

    if (selectedDate) 
	{
        document.getElementById('datePicker').value = selectedDate;
    }
	
	// Выполняем переход на страницу Remont по нажатию кнопки "Рапорт ремонта"
		document.getElementById('back').addEventListener('click', function(event) {
        event.preventDefault();
        window.location.href = 'http://10.10.3.41/rifar/TUBOG/Uchet/Raport/index.php';
    });
};
</script>

</body>
</html>