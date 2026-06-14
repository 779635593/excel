<?php

require_once __DIR__ . '/vendor/autoload.php';

use zhuoxin\excel\CsvExport;

// 导出 Csv
try {
	$excelExport = new CsvExport('order', './excelout');
	// 设置表头
	$excelExport->setHeader(['id', '姓名', '性别']);
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
		$excelExport->addData($data);
	}
	// 关闭
	$excelExport->close();
} catch (\Exception $e) {
	var_dump($e->getMessage());
}
