<?php

require 'phpoffice/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$name = $_POST['name'];
$email = $_POST['email'];
$message = $_POST['message'];

$file = "contact_data.xlsx";

/* If file exists load it */
if(file_exists($file))
{
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
    $sheet = $spreadsheet->getActiveSheet();
    $row = $sheet->getHighestRow()+1;
}
else
{
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    $sheet->setCellValue('A1','Name');
    $sheet->setCellValue('B1','Email');
    $sheet->setCellValue('C1','Message');

    $row = 2;
}

/* Insert data */
$sheet->setCellValue('A'.$row,$name);
$sheet->setCellValue('B'.$row,$email);
$sheet->setCellValue('C'.$row,$message);

/* Save Excel */
$writer = new Xlsx($spreadsheet);
$writer->save($file);

echo "<script>
alert('Message Saved Successfully');
window.location.href='index.html';
</script>";

?>