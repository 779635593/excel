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
