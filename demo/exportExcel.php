<?php

require_once __DIR__ . '/vendor/autoload.php';

use zhuoxin\excel\ExcelExport;

// 导出表格
try {
	$excelExport = new ExcelExport('order');
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
