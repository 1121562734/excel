<?php
	include 'D:\phpStudy\WWW\test\PHPExcel\PHPExcel.php';
	date_default_timezone_set('America/Los_Angeles');

	class fill_template {
		var $startrow = 0;
		function __construct($fn) {
			$this->tpl = PHPExcel_IOFactory::load($fn);
			$this->target = clone $this->tpl;
		}
		function add_data($ar) {
			//    if(!isset($this->target)) $this->target = clone $this->tpl;
			$sheet = $this->tpl->getActiveSheet();
			//获取最大行
			$maxRows    = $sheet->getHighestRow();
			$maxColumns = $sheet->getHighestColumn();
			//获取最大列
			$maxColumns = PHPExcel_Cell::columnIndexFromString($maxColumns);
			$i = 0;
			for($iR = 2; $iR <= $maxRows; $iR++){
				//内循环获取对应的单元格信息
				//使用函数 : getCellByColumnAndRow(列索引 从0开始, 行索引 从0开始)
				for($iC = 0; $iC < $maxColumns; $iC++){
					$txt = $sheet->getCellByColumnAndRow($iC, $iR)->getValue();
					if(preg_match('/{(.+)}/e',$txt ,$match)){
						//获取ar的值
						$vo=$ar[$match[1]];
						 $this->target->getActiveSheet()->getCellByColumnAndRow($iC, $iR)->setValue($vo);
					}
				}
			}
		}
		function output($fn) {
			$t = PHPExcel_IOFactory::createWriter($this->target, 'Excel5');
			$t->save($fn);
		}
	}

	$p = new fill_template('test.xls');
	$p->add_data(array('name'=>'cpj','age'=>'qw3'));
	$p->output('xxx.xls');
