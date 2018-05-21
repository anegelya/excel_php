<?php

    require('PHPExcel.php');

    function ExcelPHPParse($file_name) {
        $Reader = PHPExcel_IOFactory::createReaderForFile($file_name);
        $Reader->setReadDataOnly(true);
        $objXLS = $Reader->load($file_name);
        $objWorksheet = $objXLS->getActiveSheet();

        $flats = array();
        $columns = array();

    
        for ($i = 1; $i <= $objWorksheet->getHighestRow(); $i++) {  
            
            $nColumn = PHPExcel_Cell::columnIndexFromString(
                $objWorksheet->getHighestColumn());
            
            for ($j = 0; $j < $nColumn; $j++) {
                $value = $objWorksheet->getCellByColumnAndRow($j, $i)->getValue();
                $flat;
                $explication;

                if($i !== 2) {
                    if($j===0 && $value) {
                        $flat = $value;
                        $flats[$value] = array();
                    } else if($j===2 && $value) {
                        $flats[$flat][$value] = "";
                        $explication = $value;
                        if(!in_array($value, $columns)) {
                            $columns[] = $value;
                        }
                    } else if($j===3 && $value) {
                        if(!$flats[$flat][$explication]) {
                            $flats[$flat][$explication] = $value;
                        }
                    }
                }
            }
        }

        sort($columns);

        $res = '"Квартира";';

        foreach($columns as $col) {
            $res .= '"'.$col .= '";';
        }

        $res .= "\r\n";

        foreach($flats as $key => $flat) {
            $res .= '"'.$key .= '";';
            foreach($columns as $col) {
                if($flat[$col]) {
                    $res .= '"'.$flat[$col] .= '";';
                } else {
                    $res .= ";";
                }
            }

            $res .= "\r\n";
        }

        $f = fopen("out.csv", "w"); 
        fwrite($f, $res);
        fclose($f);

        echo ("Файл створено та записано!\n");
    }

    if(isset($_POST['submit'])){

        $target_dir = "uploads/";
        $target_file = $target_dir . basename($_FILES["excel-file"]["name"]);

        if (move_uploaded_file($_FILES["excel-file"]["tmp_name"], $target_file)) {
            echo "Файл успішно завантажено.\n";
            ExcelPHPParse($target_file);
        } else {
            echo "Можлива атака за допомогою файлів завантаження!\n";
        }
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
        <input type="file" id="excel-file" name="excel-file">
        <input type="submit" value="Пропарсити" name="submit">
    </form>
</body>
</html>