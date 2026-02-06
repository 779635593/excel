<?php

require_once __DIR__ . '/vendor/autoload.php';

use zhuoxin\excel\ExcelToCsv;

// 导出表格
try {
	$excelToCsv = new ExcelToCsv('order');
	// 设置表头
	$excelToCsv->setHeader(['id', '姓名', '性别']);
	$datas = [
		[
			'id'   => 1,
			'name' => 'xiaoming',
			'sex'  => '男',
		],
		[
			'id'   => 2,
			'name' => 'xiaoli',
			'sex'  => '女',
		],
	];
	foreach ($datas as $data) {
		// 追加数据
		$excelToCsv->addData($data);
	}
	// 关闭
	$excelToCsv->close();
} catch (\Exception $e) {
	var_dump($e->getMessage());
}
