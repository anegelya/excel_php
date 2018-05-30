<?php
    require('PHPExcel.php');

    if(isset($_POST['submit'])){
        // Отримуємо кулькість файлів
        $countfiles = count($_FILES['excel-file']['name']);

        // Cтворюємо загальний асоціативний масив для квартир
        $apparts = array();
        
        // Для кожного з файлів здійснюємо парсинг
        for($l=0;$l<$countfiles;$l++){
            $target_dir = "uploads/";
            $target_file = $target_dir . basename($_FILES["excel-file"]["name"][$l]);

            if (move_uploaded_file($_FILES["excel-file"]["tmp_name"][$l], $target_file)) {
                echo "<p>Файл номер ".($l+1)." успішно завантажено.\n</p>";
                
                $Reader = PHPExcel_IOFactory::createReaderForFile($target_file);
                $Reader->setReadDataOnly(true);
                $objXLS = $Reader->load($target_file);
                $objWorksheet = $objXLS->getActiveSheet();

                $flat = null;

                // Проходимось по рядкам таблиці
                for ($i = 1; $i <= $objWorksheet->getHighestRow(); $i++) {  
                    
                    $nColumn = PHPExcel_Cell::columnIndexFromString(
                        $objWorksheet->getHighestColumn());
                    
                    // Проходимось по колонках рядка таблиці
                    for ($j = 0; $j < $nColumn; $j++) {
                        $value = $objWorksheet->getCellByColumnAndRow($j, $i)->getValue();

                        // Пропускаємо пустий рядок і рядок із заголовками таблиці
                        if($i > 2) {
                            if($j===0 && $value) {
                                /* Якщо це перша клітинка в таблиці і вона не пуста,
                                тобто відповідає базовому ID квартири,
                                то перевіряємо її на відповідність регулярним виразам,
                                зберігаємо її як поточну квартиру в локальну змінну та
                                додаємо до загального масиву квартир.
                                Також додаємо до нього масив приміщень,
                                загальну площу квартири
                                та масив поверхів, де розміщена квартира */
                                if(preg_match('/^MZK/', $value) !== 1) {
                                    $flat = $value;
                                    $apparts[$value] = array();
                                    $apparts[$flat]['rooms'] = array();
                                    $apparts[$flat]['area'] = 0;
                                    $apparts[$flat]['floor'] = array(); 
                                } else {
                                    // В іншому разі обнуляємо поточну квартиру
                                    $flat = null;
                                }
                            } else if($flat) {
                                if($j===1 && $value) {
                                    /* Якщо це третя клітинка в таблиці і вона не пуста,
                                    тобто відповідає коду приміщення,
                                    то зберігаємо її до масиву приміщень квартири,
                                    туди додеємо ї назву, площу та поверх.
                                    Площу та поверх кімнати також додаємо 
                                    до загальної площі та масиву поверхів квартири */
                                    $apparts[$flat]['rooms'][$value] = array();
                                    $name = $objWorksheet->getCellByColumnAndRow($j+1, $i)->getValue();
                                    $apparts[$flat]['rooms'][$value]['name'] = $name;
    
                                    $area = $objWorksheet->getCellByColumnAndRow($j+2, $i)->getValue();
                                    $apparts[$flat]['area'] = $apparts[$flat]['area'] + $area;
                                    $apparts[$flat]['rooms'][$value]['area'] = $area;
    
                                    $floor = $objWorksheet->getCellByColumnAndRow($j+3, $i)->getValue();
                                    $apparts[$flat]['rooms'][$value]['floor'] = $floor;
                                    if(!in_array($floor, $apparts[$flat]['floor'])) {
                                        $apparts[$flat]['floor'][] = $floor;
                                    }
                                }
                            }
                        }
                    }
                }

                echo ("<p>Файл номер ".($l+1)." додано до таблиці</p>");
            } else {
                echo "<p>Можлива атака за допомогою файлів завантаження!\n</p>";
            }
        }

        $firstTable = makeFirstTable($apparts);
        $secondTable = makeSecondTable($apparts);

        $fileName = strval($_POST['name']);

        writeCSV($firstTable, $fileName, 'first_table');
        writeCSV($secondTable, $fileName, 'second_table');
    }

    function writeCSV($table, $fileName, $tableName) {
        if($file = @fopen($fileName.'_'.$tableName.".csv", "x")) {
            foreach ($table as $row) {
                fputcsv($file, $row);
            }
            echo('<p>Файл '.$fileName.'_'.$tableName." успішно створено!</p>");
            fclose($file);
        } else {
            echo('<p>Файл із іменем '.$fileName.'_'.$tableName." вже існує. Спробуйте інше ім'я</p>");
        }
    }

    function makeFirstTable($apparts) {
        $table = array(
            array(
                'Номер квартири',
                'Тип',
                'Секція',
                'Під’їзд',
                'Поверхи',
                'Загальна площа',
                'К-ть кімнат',
                'Базовий ID'
            )
        );

        foreach ($apparts as $flatId => $flat) {

            preg_match('/^(\S+)\/(\S+)\/(\S+)\.(\S+)\_(\S+)\_(\S+)\_(\S+)/', $flatId, $idMatches);
            $row = array(
                $idMatches[2],
                $idMatches[3].'.'.$idMatches[4],
                $idMatches[5],
                $idMatches[6],
                implode(',', $flat['floor']),
                $flat['area'],
                $idMatches[3],
                $flatId
            );

            $table[] = $row;
        }

        return $table;
    }

    function makeSecondTable($apparts) {
        $table = array(
            array(
                'Базовий ID',
                'Код приміщення',
                'Тип приміщення',
                'Площа',
                'Поверх'
            )
        );

        foreach ($apparts as $flatId => $flat) {
            foreach ($flat['rooms'] as $roomId => $room) {
                $row = array(
                    $flatId,
                    $roomId,
                    $room['name'],
                    $room['area'],
                    $room['floor']
                );

                $table[] = $row;
            }
        }

        return $table;
    }
?>

<!DOCTYPE html>
<html lang="uk">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <title>Excel PHP parser</title>

    <link href="" type="text/css" rel="stylesheet">
</head>

<body>
    <form method="POST" enctype="multipart/form-data">
        <input type="file" id="excel-file" name="excel-file[]" multiple required>
        <input type="name" required name="name" placeholder="Введіть сюди назву файлу">
        <input type="submit" value="Пропарсити" name="submit">
    </form>
</body>
</html>