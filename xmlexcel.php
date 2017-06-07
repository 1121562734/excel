<?php
	/**
	 * 替换并转换
	 * @param $filename  文件名
	 * @param $patterns   模版中的文字
	 * @param $replacements  要替换的文字
	 */
	function out_excel($filename,$date,$outname){

		$string=input_excel($filename);
		$patterns=array();
		$replacements=array();
		foreach($date as $key =>$vo){
			$patterns[]=$key;
			$replacements[]=$vo;
		}
		$string = preg_replace($patterns,$replacements,$string);
		$outname=iconv('utf-8', 'GB2312//IGNORE', $outname);
		file_put_contents($outname, $string);
	}

	/**
	 * 获取文件
	 * @param $filename
	 * @return string
	 */
	function input_excel($filename){
		ob_start();
		$filename=iconv('utf-8', 'GB2312//IGNORE', $filename);
		include $filename;
		return ob_get_clean();
	}


	/************
	调用方法
	 **********************/
	$filename="长安贷款履约保证保险投保单.xml";
	$date=array(
		'/${name}/'=>"www",
		'/${phone}/'=>"15751215451"
	);
	$outname="长安贷款履约保证保险投保单.xls";

	out_excel($filename, $date, $outname);

