<?php

namespace zhuoxin\excel;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;

// Excel 工具类，仅用于读取文件，导出使用 ExcelToCsv ，不满足再使用 PhpOffice 自行编写
class ExcelUtil
{

	/**
	 * 读取文件
	 * 是否保留头部，从第1行读取，否则从第2行读取
	 *
	 * @param        $file_path  // 文件路径
	 * @param  bool  $keepHead   // 是否保留头部，是从第1行读取，否则从第2行读取
	 *
	 * @return array
	 * @throws \Exception
	 */
	public static function readExcel($file_path, bool $keepHead = false): array
	{
		try {
			if ( ! file_exists($file_path)) {
				throw new \Exception('文件不存在');
			}
			// 1. 加载Excel文件
			$spreadsheet = IOFactory::load($file_path);
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
			// 表格数据
			$excelData = [];
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
				// 追加到总数据中
				$excelData[] = $rowData;
			}

			// 打印读取到的数据，后续可循环插入数据库
			return $excelData;
		} catch (\Throwable $e) {
			throw new \Exception('读取文件错误：' . $e->getMessage());
		}
	}

}