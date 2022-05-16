<?php

require 'vendor/autoload.php';
require 'conexion.php';

use PhpOffice\PhpSpreadsheet\{Spreadsheet, IOFactory};
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\FIll;


//styling arrays
//table head style
$tableHead = [
	'font'=>[
		'color'=>[
			'rgb'=>'FFFFFF'
		],
		'bold'=>true,
		'size'=>11
	],
	'fill'=>[
		'fillType' => Fill::FILL_SOLID,
		'startColor' => [
			'rgb' => '538ED5'
		]
	],
];
//even row
$evenRow = [
	'fill'=>[
		'fillType' => Fill::FILL_SOLID,
		'startColor' => [
			'rgb' => '00BDFF'
		]
	]
];
//odd row
$oddRow = [
	'fill'=>[
		'fillType' => Fill::FILL_SOLID,
		'startColor' => [
			'rgb' => '00EAFF'
		]
	]
];

$sql = "SELECT id, nombre, notaUno, notaDos, notaTres, notaCuatro, notaCinco, notaSeis, notaSiete FROM alumnos";
$resultado = $mysqli->query($sql);

$excel = new Spreadsheet();
$hojaActiva = $excel->getActiveSheet();


$hojaActiva->setCellValue('B5', 'ID');
$hojaActiva->setCellValue('C5', 'Nombre');
$hojaActiva->setCellValue('D5', 'Nota 1');
$hojaActiva->setCellValue('E5', 'Nota 2');
$hojaActiva->setCellValue('F5', 'Nota 3');
$hojaActiva->setCellValue('G5', 'Nota 3');
$hojaActiva->setCellValue('H5', 'Nota 5');
$hojaActiva->setCellValue('I5', 'Nota 6');
$hojaActiva->setCellValue('J5', 'Nota 7');
$hojaActiva->setCellValue('K5', 'Promedio');
$hojaActiva->setCellValue('L5', 'Promedio ponderado');

$fila = 6;
while($rows = $resultado->fetch_assoc()){
    $hojaActiva->setCellValue('B'.$fila, $rows['id']);
    $hojaActiva->setCellValue('C'.$fila, $rows['nombre']);
    $hojaActiva->setCellValue('D'.$fila, $rows['notaUno']);
    $hojaActiva->setCellValue('E'.$fila, $rows['notaDos']);
    $hojaActiva->setCellValue('F'.$fila, $rows['notaTres']);
    $hojaActiva->setCellValue('G'.$fila, $rows['notaCuatro']);
    $hojaActiva->setCellValue('H'.$fila, $rows['notaCinco']);
    $hojaActiva->setCellValue('I'.$fila, $rows['notaSeis']);
    $hojaActiva->setCellValue('J'.$fila, $rows['notaSiete']);
    $fila++;
}

/* Here there will be some code where you create $spreadsheet */

// redirect output to client browser
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="myfile.xlsx"');
header('Cache-Control: max-age=0');

$writer = IOFactory::createWriter($excel, 'Xlsx');
$writer->save('php://output');
exit;