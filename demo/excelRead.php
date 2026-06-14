<?php

require_once __DIR__ . '/vendor/autoload.php';

use zhuoxin\excel\ExcelRead;

// 读取表格
try {
	$file_path = './test.xls';
	$res       = ExcelRead::readExcel($file_path);
	var_dump($res);
} catch (\Exception $e) {
	var_dump($e->getMessage());
}

// 读取表格 生成器方式
try {
	$file_path = './test.xls';
	foreach (ExcelRead::readExcelGenerator($file_path) as $row) {
		var_dump($row);
	}
} catch (\Exception $e) {
	var_dump($e->getMessage());
}
