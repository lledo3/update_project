<?php
    /**
     * Just in case the user running this doesn't
     * have a default timezone set...we set one.
     */
    date_default_timezone_set("America/New_York");
    /**
     * Let's not rely on relative pathing if we can
     * avoid it
     */
    include __DIR__.'/excel_reader2.php';

    $worksheets = getAllExcelFiles(__DIR__, 'xls');
    $excelDocTotals = array();
?>
<!doctype html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <title>Project 1</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- Some default styles just in case the user is not online when viewing the file -->
    <style>
        .container { margin-top: 30px; }
        summary { padding: 10px 0 2px 0; }
        table { border: 1px solid #ccc; border-collapse: collapse; margin: 5px 0 10px 0;}
        table td,
        table th { border: 1px solid #ccc; border-collapse: collapse; padding: 5px;}
        body { font-family: Arial; }
    </style>
    <link rel="stylesheet" type="text/css" href="//maxcdn.bootstrapcdn.com/bootstrap/3.2.0/css/bootstrap.min.css">
</head>
<body>
    <div class="container">
        <div class="col-md-6 col-md-offset-3">
            <?php foreach($worksheets as $worksheet):
                    $excelDocTotals = getWorksheetValues($worksheet); ?>
                    <table class="table table-bordered table-striped table-hover">
                        <summary><strong><?=$excelDocTotals['name'];?> </strong> <small>(<?=$excelDocTotals['date']; ?>)</small></summary>
                        <thead>
                            <tr>
                                <th>&nbsp;</th>
                                <th>Impressions</th>
                                <th>Earnings</th>
                                <th>eCPM</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><strong>Total</strong></td>
                                <td><?=$excelDocTotals['totalImp']; ?></td>
                                <td class="text-center">&ndash;</td>
                                <td class="text-center">&ndash;</td>
                            </tr>
                            <tr>
                                <td><strong>Filled</strong></td>
                                <td><?=$excelDocTotals['filledImp']; ?></td>
                                <td>$<?=$excelDocTotals['earnings']; ?></td>
                                <td>$<?=$excelDocTotals['totalECPM']; ?></td>
                            </tr>
                        </tbody>
                    </table>
            <?php endforeach; ?>
        </div>
    </div>
</body>
</html>

<?php
    /**
     * Grabs all the files in $directory and returns us an
     * array of only the files we want
     *
     * @return Array
     */
    function getAllExcelFiles($directory, $type)
    {
        $files = scandir($directory);
        $excelFiles = array();

        foreach($files as $sheet)
        {
            $fileInfo = pathinfo($sheet);
            if($fileInfo['extension'] === $type)
            {
                $excelFiles[] = $sheet;
            }
        }
        return $excelFiles;
    }

    /**
     * Grabs the excel file and does all the needed calculations
     * and returns an array of all those values
     *
     * @param  String $workbook
     * @return Array
     */
    function getWorksheetValues($workbook)
    {
        $values = array();
        $sheet = new Spreadsheet_Excel_Reader($workbook);

        $values['name'] = $sheet->val(4,'C');
        $values['date'] = date("m/d/y", strtotime($sheet->val(6,'C')));

        $passback = $sheet->raw(11,'F');
        $values['filledImp'] = $sheet->raw(12,'F') + $sheet->raw(13,'F') + $sheet->raw(14,'F') + $sheet->raw(15,'F') + $sheet->raw(16,'F');
        $values['totalImp'] = $passback + $values['filledImp'];

        $values['rev'] = ($sheet->raw(12,'J') + $sheet->raw(13,'J') + $sheet->raw(14,'J') + $sheet->raw(15,'J') + $sheet->raw(16,'J')) / 2;
        $values['ecpm'] = ($values['rev'] / ($values['totalImp'] / 1000));

        $values['earnings'] = round($values['rev'], 4);
        $values['totalECPM'] = round($values['ecpm'], 4);

        return $values;
    }