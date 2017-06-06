<?php

	include 'PHPExcel.php';
	//include 'PHPExcel/Writer/Excel2007.php';
	$objPHPExcel = new PHPExcel();

	//$template = $PHPWord->loadTemplate('template.docx');
	//$objPHPExcel = $objReader->load("template.xls" );
	//$objPHPExcel = PHPExcel_Autoloader::Load("template.xlsx");


	$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
	$objWriter->save("test.xlsx");
?>

