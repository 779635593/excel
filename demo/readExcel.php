<?php

require_once __DIR__ . '/vendor/autoload.php';

use zhuoxin\excel\ExcelUtil;

// 读取表格
try {
	$file_path = './test.xls';
	$res       = ExcelUtil::readExcel($file_path);
	var_dump($res);
} catch (\Exception $e) {
	var_dump($e->getMessage());
}

// 读取表格 生成器方式
try {
	$file_path = './test.xls';
	foreach (ExcelUtil::readExcelGenerator($file_path) as $row) {
		var_dump($row);
	}
} catch (\Exception $e) {
	var_dump($e->getMessage());
}
