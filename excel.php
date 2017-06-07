<?php

	include 'PHPExcel.php';
	date_default_timezone_set('Asia/Shanghai');
	$fn='长安贷款履约保证保险投保单.xlsx';
	$fn=iconv('utf-8', 'GB2312//IGNORE', $fn);
	$objPHPexcel = PHPExcel_IOFactory::load($fn);
	$objWorksheet = $objPHPexcel->getActiveSheet();
	//$objWorksheet->getCell('C12')->setValue('John');
	//$objWorksheet->getCell('D8')->setValue('Smith');
	$objWriter = PHPExcel_IOFactory::createWriter($objPHPexcel, 'Excel5');
	$objWriter->save('write.xls');


