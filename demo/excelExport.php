<?php

require_once __DIR__ . '/vendor/autoload.php';

use zhuoxin\excel\ExcelExport;

// 导出 Excel
try {
	$file        = 'order';
	$excelExport = new ExcelExport($file, './excelout');
	// 设置图片属性
	$excelExport->setImageAttr(['img1', 'img2']);
	// 设置表头
	$excelExport->setHeader(['id', '姓名', '图片2', '性别', '图片1']);

	$datas = [
		[
			'id'   => 1,
			'name' => 'xiaoming',
			'img2' => './img/1.jpg',
			'sex'  => '男',
			'img1' => './img/1.jpg',
		],
		[
			'id'   => 2,
			'name' => 'xiaoli',
			'img2' => './img/2.jpg',
			'sex'  => '女',
			'img1' => './img/4.jpg',
		],
		[
			'id'   => 3,
			'name' => 'xiaoli3',
			'img2' => './img/3.jpg',
			'sex'  => '女3',
			'img1' => './img/3.jpg',
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
