<?php
	/**
	 * 存在的单元格跟样式会被改变
	 */
	include 'PHPExcel.php';
	date_default_timezone_set('Asia/Shanghai');

	class fill_template {
		var $startrow = 0;
		function __construct($fn) {
			$fn = iconv('utf-8', 'GB2312//IGNORE', $fn);
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
			for($iR = 2; $iR <= $maxRows; $iR++){
				//内循环获取对应的单元格信息
				//使用函数 : getCellByColumnAndRow(列索引 从0开始, 行索引 从0开始)
				for($iC = 0; $iC < $maxColumns; $iC++){
					$txt = $sheet->getCellByColumnAndRow($iC, $iR)->getValue();
					if(preg_match('/\${(.+)}/e',$txt ,$match)){
						//获取ar的值
						//$vo=iconv('utf-8', 'GB2312//IGNORE', $ar[$match[1]]);
						$vo=$ar[$match[1]];
						$this->target->getActiveSheet()->getCellByColumnAndRow($iC, $iR)->setValue($vo);
					}
				}
			}
		}
		function output($fn) {
			$fn = iconv('utf-8', 'GB2312//IGNORE', $fn);
			$t = PHPExcel_IOFactory::createWriter($this->target, 'Excel5');
			$t->save($fn);
		}
	}

	$p = new fill_template('表单/长安贷款履约保证保险投保单.xls');
	$p->add_data(array('name'=>'陈鹏杰','age'=>'年龄'));
	$p->output('苏宁/长安贷款履约保证保险投保单2.xls');



