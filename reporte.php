<?php
//call the autoload
require 'vendor/autoload.php';
require 'conexion.php';
//load phpspreadsheet class using namespaces
use LDAP\Result;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
//call iofactory instead of xlsx writer
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

\PhpOffice\PhpSpreadsheet\Cell\Cell::setValueBinder( new \PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder() );
\PhpOffice\PhpSpreadsheet\Calculation\MathTrig\Sum::product();

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

//conection DB
$sql = "SELECT id, nombre FROM alumnos";
$resultado = $mysqli->query($sql);

//styling arrays end
//make a new spreadsheet object
$spreadsheet = new Spreadsheet();
//get current active sheet (first sheet)
$sheet = $spreadsheet->getActiveSheet();

//set default font
$spreadsheet->getDefaultStyle()
	->getFont()
	->setName('Arial')
	->setSize(10);

$spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);

//heading
$spreadsheet->getActiveSheet()->setCellValue('D3',"Reporte académico de notas");

//merge heading
$spreadsheet->getActiveSheet()->mergeCells("D3:L3");

// set font style
$spreadsheet->getActiveSheet()->getStyle('D3')->getFont()->setSize(12);

// set cell alignment
$spreadsheet->getActiveSheet()->getStyle('D3')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

//setting column width
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(5);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(30);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('J')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('K')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('L')->setWidth(10);
$spreadsheet->getActiveSheet()->getColumnDimension('M')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('N')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('O')->setWidth(15);
$spreadsheet->getActiveSheet()->getColumnDimension('P')->setWidth(15);

//Cabecera de la tabla
$spreadsheet->getActiveSheet()
	->setCellValue('B3',"ID")->mergeCells('B3:B4')
	->setCellValue('C3',"ALUMNO")->mergeCells('C3:C4')
	->setCellValue('D4',"Nota 1")
	->setCellValue('E4',"Nota 2")
	->setCellValue('F4',"Nota 3")
	->setCellValue('G4',"Nota 4")
	->setCellValue('H4',"Nota 5")
	->setCellValue('I4',"Nota 6")
	->setCellValue('J4',"Nota 7")
	->setCellValue('K4',"Nota 8")
	->setCellValue('L4',"Nota 9")
	->setCellValue('M3',"Cognitiva DF")
	->setCellValue('N3', "Procedimental")
	->setCellValue('O3', "Social")
	->setCellValue('P3', "Definitiva")->mergeCells('P3:P4');

	
//set font style and background color
$spreadsheet->getActiveSheet()->getStyle('B3:P3')->applyFromArray($tableHead);
$spreadsheet->getActiveSheet()->getStyle('D4:L4')->applyFromArray($tableHead);

$spreadsheet->getActiveSheet()->setCellValue('M4', '%');
$spreadsheet->getActiveSheet()->setCellValue('n4', '%');
$spreadsheet->getActiveSheet()->setCellValue('O4', '%');
// $spreadsheet->getActiveSheet()->setCellValue('H4', '%');
// $spreadsheet->getActiveSheet()->setCellValue('I4', '%');
// $spreadsheet->getActiveSheet()->setCellValue('J4', '%');
// $spreadsheet->getActiveSheet()->setCellValue('K4', '%');

//the content
$fila = 5;
while($rows = $resultado->fetch_assoc()){
    $spreadsheet->getActiveSheet()
	->setCellValue('B'.$fila, $rows['id'])
    ->setCellValue('C'.$fila, $rows['nombre'])
	->setCellValue('M'.$fila, "=AVERAGE(D$fila:L$fila)")
	//->setCellValue('M'.$fila, \PhpOffice\PhpSpreadsheet\Calculation\MathTrig\Sum::product("(B$4:k$4);(C5:K5)"));
	->setCellValue('P'.$fila, "=SUMPRODUCT(M$4:O$4,M$fila:O$fila)");
    $fila++;
}
//Condición de limite minimo de aprobación
$conditional = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
$conditional->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS);
$conditional->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_LESSTHAN);
$conditional->addCondition(3,0);
$conditional->getStyle()->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
$conditional->getStyle()->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
$conditional->getStyle()->getFill()->getStartColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE);

$conditionalStyles = $spreadsheet->getActiveSheet()->getStyle("D5:P$fila")->getConditionalStyles();
$conditionalStyles[] = $conditional;

$spreadsheet->getActiveSheet()->getStyle("D5:P$fila")->setConditionalStyles($conditionalStyles);

//Condicional limite maximo de nota a obtener
$conditional = new \PhpOffice\PhpSpreadsheet\Style\Conditional();
$conditional->setConditionType(\PhpOffice\PhpSpreadsheet\Style\Conditional::CONDITION_CELLIS);
$conditional->setOperatorType(\PhpOffice\PhpSpreadsheet\Style\Conditional::OPERATOR_GREATERTHAN);
$conditional->addCondition(5,0);
$conditional->getStyle()->getFont()->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_WHITE);
$conditional->getStyle()->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
$conditional->getStyle()->getFill()->getStartColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);

$conditionalStyles = $spreadsheet->getActiveSheet()->getStyle("D5:P$fila")->getConditionalStyles();
$conditionalStyles[] = $conditional;

$spreadsheet->getActiveSheet()->getStyle("D5:P$fila")->setConditionalStyles($conditionalStyles);

//set the header first, so the result will be treated as an xlsx file.
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

//make it an attachment so we can define filename
header('Content-Disposition: attachment;filename="result.xlsx"');

//create IOFactory object
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
//save into php output
$writer->save('php://output');
