<!doctype html>
<html lang="en">
<head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css" integrity="sha384-MCw98/SFnGE8fJT3GXwEOngsV7Zt27NXFoaoApmYm81iuXoPkFOJwJ8ERdknLPMO" crossorigin="anonymous">
    <link rel="stylesheet" href="style.css">
    <title>Neoship</title>
</head>
<body>
<?php
    require 'vendor/autoload.php';
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
    /**  Defining type of files and names of files  **/
    $inputFileType = 'Xlsx';
    $inputFileName = 'zasielky.xlsx';
    $inputFileName1 = 'cennik.xlsx';
    /**  Create a new Reader of the type defined in $inputFileType  **/
    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($inputFileType);
    $spreadsheet = $reader->load($inputFileName);
    $spreadsheet1 = $reader->load($inputFileName1);
?>

<div class="container wrapper" style="padding-top: 50px;">
    <table class="table table-striped">
        <th>Krajina</th>
        <th>Hmotnosť</th>
        <th>Príplatkové služby</th>
        <th>Doprava bez DPH</th>
        <th>Doprava s DPH</th>
        <th>Celkovo bez DPH</th>
        <th>Celkovo s DPH</th>
        <?php
            $i=2;
            do{
                $i++;
                /**  Variables about order **/
                $country = $spreadsheet->getActiveSheet()->getCell('F'."$i")->getValue();
                /**  $weight for calculation **/
                $weight = $spreadsheet->getActiveSheet()->getCell('K'."$i")->getValue();
                /**  $weight for table **/
                $weight_default = $spreadsheet->getActiveSheet()->getCell('K'."$i")->getValue();
                $bonus = $spreadsheet->getActiveSheet()->getCell('L'."$i")->getValue();
                $value = $spreadsheet->getActiveSheet()->getCell('I'."$i")->getValue();
                $delivery = 0;

                if($country==""){break;}
                /**  Calculating total price according to price list in excel file  **/
                if($value>0) {

                    if ($country == "Slovenská republika [SK]") {
                        /**  $a- help variable for separating maximum weight value from string  **/
                        $a = 4;

                        for ($j = 4; $j <= 15; $j++) {
                            /**  $t- max weight of interval $-val is price for this weight  **/
                            $to = $spreadsheet1->getActiveSheet()->getCell('G' . $j)->getValue();
                            $val = $spreadsheet1->getActiveSheet()->getCell('J' . $j)->getValue();
                            $t = substr($to, strlen($to) - $a);
                            /**  If $j==15 there is no additive price  **/
                            if (($weight <= (double)$t) || $j == 15) {
                                /**  $max defines max price for interval  **/
                                $maximum = $spreadsheet1->getActiveSheet()->getCell('G17')->getValue();
                                $maxi = substr($maximum, strlen($maximum) - 7);
                                $max = str_replace(",", "", $maxi);
                                if ($value >= 0.01 && $value <= (double)$max) {
                                    $delivery += $spreadsheet1->getActiveSheet()->getCell('J17')->getValue();
                                }
                                $delivery += $val;
                                break;
                            }
                            /**  We have to increse $a because $weight is higher  **/
                            if ($j == 5) {
                                $a = 5;
                            }
                        }
                    }
                    if ($country == "Maďarská republika [HU]" || $country == "Česká republika [CZ]") {
                        $a = 4;
                        for ($j = 37; $j <= 45; $j++) {
                            $to = $spreadsheet1->getActiveSheet()->getCell('G' . $j)->getValue();
                            $val = $spreadsheet1->getActiveSheet()->getCell('J' . $j)->getValue();
                            $t = substr($to, strlen($to) - $a);

                            if($j==45){
                                if($weight>$t){
                                    $weight*=0.5;
                                }
                            }
                            if ($weight <= (double)$t) {
                                /**  $ci $l $f defines interval of price for calculation  **/
                                $cim = $spreadsheet1->getActiveSheet()->getCell('G47')->getValue();
                                $ci = substr($cim, strlen($cim) - 6);

                                $lim = $spreadsheet1->getActiveSheet()->getCell('G48')->getValue();
                                $li = substr($lim, strlen($lim) - 8);
                                $l = str_replace(",", "", $li);

                                $fim = $spreadsheet1->getActiveSheet()->getCell('G49')->getValue();
                                $fi = substr($fim, strlen($fim) - 8);
                                $f = str_replace(",", "", $fi);

                                if ($value >= 0.01 && $value <= (double)$ci) {
                                    $delivery += $spreadsheet1->getActiveSheet()->getCell('J47')->getValue();
                                }
                                if ($value > (double)$ci && $value <= (double)$l) {
                                    $delivery += $spreadsheet1->getActiveSheet()->getCell('J48')->getValue();
                                }
                                /**  $do defines percents from price list  **/
                                $dob = $spreadsheet1->getActiveSheet()->getCell('J49')->getValue();
                                $d = substr($dob, 0, 3);
                                $do = (double)$d;
                                if ($value > (double)$l && $value <= (double)$f) {
                                    $delivery *= (($do / 100) + 1);
                                }
                                $delivery += $val;
                                break;
                            }
                            if ($j == 37) {
                                $a = 5;
                            }
                        }
                    }
                    if ($country == "Rakúsko [AT]") {
                        $a = 4;
                        for ($j = 20; $j <= 30; $j++) {
                            $to = $spreadsheet1->getActiveSheet()->getCell('G' . $j)->getValue();
                            $val = $spreadsheet1->getActiveSheet()->getCell('J' . $j)->getValue();
                            $t = substr($to, strlen($to) - $a);
                            if($j==30){
                                if($weight>$t){
                                    $weight*=0.5;
                                }
                            }
                            if ($weight <= (double)$t) {
                                $lim = $spreadsheet1->getActiveSheet()->getCell('G32')->getValue();
                                $li = substr($lim, strlen($lim) - 8);
                                $l = str_replace(",", "", $li);

                                $fim = $spreadsheet1->getActiveSheet()->getCell('G33')->getValue();
                                $fi = substr($fim, strlen($fim) - 8);
                                $f = str_replace(",", "", $f);

                                if ($value >= 0.01 && $value <= (double)$l) {
                                    $delivery += $spreadsheet1->getActiveSheet()->getCell('J32')->getValue();
                                }
                                $dob = $spreadsheet1->getActiveSheet()->getCell('J33')->getValue();
                                $d = substr($dob, 0, 3);
                                $do = (double)$d;
                                if ($value > (double)$l && $value <= (double)$f) {
                                    $delivery *= (($do / 100) + 1);
                                }
                                $delivery += $val;
                                break;
                            }
                            if ($j == 21) {
                                $a = 5;
                            }
                        }
                    }
                    if ($bonus == "ZM") {
                        $delivery += $spreadsheet1->getActiveSheet()->getCell('J52')->getValue();
                    }
                    if ($bonus == "TsN") {
                        $delivery += $spreadsheet1->getActiveSheet()->getCell('J53')->getValue();
                    }
                }
                ?>
    <h4>
        <tr>
            <td>
                <?=$country?>
            </td>
            <td>
                <?=$weight_default?>
            </td>
            <td>
                <?=$bonus?>
            </td>
            <td>
                <?=$delivery?>
            </td>
            <td>
                <?=$delivery_dph=round($delivery*1.2, 2)?>
            </td>
            <td>
                <?=round($value*0.8+$delivery,2)?>
            </td>
            <td>
                <?=$value+$delivery_dph ?>
            </td>
        </tr>
    </h4>
<?php }while($country!="") ?>
</table>
</div>

<!-- Optional JavaScript -->
<!-- jQuery first, then Popper.js, then Bootstrap JS -->
<script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" integrity="sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo" crossorigin="anonymous"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.3/umd/popper.min.js" integrity="sha384-ZMP7rVo3mIykV+2+9J3UJ46jBk0WLaUAdn689aCwoqbBJiSnjAK/l8WvCWPIPm49" crossorigin="anonymous"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/js/bootstrap.min.js" integrity="sha384-ChfqqxuZUCnJSK3+MXmPNIyE6ZbWh2IMqE241rYiqJxyMiZ6OW/JmZQ5stwEULTy" crossorigin="anonymous"></script>
<script src="https://code.jquery.com/jquery-3.3.1.js" integrity="sha256-2Kok7MbOyxpgUVvAk/HJ2jigOSYS2auK4Pfzbm7uH60=" crossorigin="anonymous"></script>
</body>
</html>
