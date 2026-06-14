<?php

namespace zhuoxin\excel;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;

// 读取 Excel 类
class ExcelRead
{

	/**
	 * 读取文件
	 *
	 * @param        $filePath  // 文件路径
	 * @param  bool  $keepHead  // 是否保留头部，是从第1行读取，否则从第2行读取
	 *
	 * @return array
	 * @throws \Exception
	 */
	public static function readExcel($filePath, bool $keepHead = false): array
	{
		$excelData = [];
		foreach (self::readExcelGenerator($filePath, $keepHead) as $rowData) {
			$excelData[] = $rowData;
		}

		return $excelData;
	}

	/**
	 * 读取文件 生成器方式
	 *
	 * @param        $filePath  // 文件路径
	 * @param  bool  $keepHead  // 是否保留头部，是从第1行读取，否则从第2行读取
	 *
	 * @return \Generator
	 * @throws \Exception
	 */
	public static function readExcelGenerator($filePath, bool $keepHead = false): \Generator
	{
		try {
			if ( ! file_exists($filePath)) {
				throw new \Exception('文件不存在');
			}
			// 1. 加载Excel文件
			$spreadsheet = IOFactory::load($filePath);
			// 指定读取的 sheet
			$spreadsheet->setActiveSheetIndex(0);
			// 获取 sheet 对象
			$sheet = $spreadsheet->getActiveSheet();
			// 2. 获取Excel的最大行数和列数
			// 总行数
			$rowTotalNum = $sheet->getHighestRow();
			// 总列数字母（格式如：C）
			$highestColumn = $sheet->getHighestColumn();
			// 总列数字母转数字（例：C -> 3）
			$columnTotalNum = Coordinate::columnIndexFromString($highestColumn);
			// 3. 循环读取每行数据
			// 是否保留头部，从第1行读取，否则从第2行读取
			$startRow = $keepHead ? 1 : 2;
			for ($row = $startRow; $row <= $rowTotalNum; $row++) {
				// 行数据
				$rowData = [];
				// 读取行的每列数据
				for ($column = 1; $column <= $columnTotalNum; $column++) {
					$rowData[] = $sheet->getCellByColumnAndRow($column, $row)->getValue();
				}
				// 使用 yield 逐行返回数据
				yield $rowData;
			}
		} catch (\Throwable $e) {
			throw new \Exception('读取文件错误：' . $e->getMessage());
		}
	}

}