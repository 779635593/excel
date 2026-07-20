<?php

namespace zhuoxin\excel;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Excel 导出类
// 导出图片或少量数据使用，数据量大使用CsvExport类
class ExcelExport
{

	// 文件名
	private string $fileName;

	// 保存目录，路径为空则是浏览器下载
	private string $saveDir = '';

	// 表格对象
	private Spreadsheet $spreadsheet;

	// 表格sheet对象
	private Worksheet $sheet;

	// 当前行数
	private int $currentRow = 1;

	// 图片列名 集合
	private array $imgColumns = [];

	// 图片路径为空时的提示信息
	private string $imgEmptyMsg = '';

	// 图片 宽
	private int $imgWidth = 60;

	// 图片 高
	private int $imgHeight = 60;

	// 行高
	private int $rowHeight = 70;

	public function __construct(string $fileName, string $saveDir = '', bool $dateSuffix = true)
	{
		// 清除文件名后缀，避免重复.xlsx
		$baseName = rtrim($fileName, '.xlsx');
		// 时间后缀
		$suffix = $dateSuffix ? '_' . date('YmdHis') : '';
		// 生成文件名+年月日时分秒.xlsx
		$this->fileName = $baseName . $suffix . '.xlsx';
		// 保存路径
		if ( ! empty($saveDir)) {
			$saveDir = rtrim($saveDir, '/');
			if ( ! is_dir($saveDir)) {
				mkdir($saveDir, 0755, true);
			}
			$this->saveDir = $saveDir;
		}

		set_time_limit(0);
		// 内存最大1G
		ini_set('memory_limit', '1G');
		// 取消执行时间限制
		ini_set('max_execution_time', 0);
		// 防止用户关闭浏览器中断导出
		ignore_user_abort(true);

		// 实例化表格对象
		$this->spreadsheet = new Spreadsheet();
		// sheet对象
		$this->sheet = $this->spreadsheet->getActiveSheet();
		$this->sheet->setTitle($fileName);
	}

	/**
	 * 设置表头
	 *
	 * @param  array  $headerArr
	 *
	 * @return void
	 */
	public function setHeader(array $headerArr)
	{
		// 列索引
		$colIdx = 0;
		foreach ($headerArr as $title) {
			$letter = $this->numToLetter($colIdx);
			$this->sheet->setCellValue("{$letter}{$this->currentRow}", $title);
			$this->sheet->getColumnDimension($letter)->setWidth(16);
			$colIdx++;
		}
		// 行号自增
		$this->currentRow++;
	}

	/**
	 * 单行写入
	 *
	 * @param  array  $rowData  // 行数据
	 *
	 * @return void
	 */
	public function addData(array $rowData)
	{
		// 是否有图片
		$haveImg = false;
		// 列索引
		$colIdx = 0;
		foreach ($rowData as $key => $cellVal) {
			// 获取列字母
			$letter = $this->numToLetter($colIdx);
			// 在图片列集合中，插入图片
			if (in_array($key, $this->imgColumns)) {
				// 图片路径为空
				if (empty($cellVal)) {
					$this->sheet->setCellValue("{$letter}{$this->currentRow}", $this->imgEmptyMsg);
				} else {
					try {
						// 插入图片到单元格
						$this->insertImage($letter, $cellVal);
						// 标识有图片
						$haveImg = true;
					} catch (\Exception $exception) {
						// 将异常信息写入该位置
						$errMsg = sprintf("图片异常，图片路径：%s,错误信息：%s ", $cellVal, $exception->getMessage());
						$this->sheet->setCellValue("{$letter}{$this->currentRow}", $errMsg);
					}
				}
			} else {
				// 非图片列，插入文字。如果长数字显示字符串，在传入数字时需加引号
				$this->sheet->setCellValue("{$letter}{$this->currentRow}", $cellVal);
			}
			$colIdx++;
		}
		// 如有图片写入 则设置当前行高
		if ($haveImg) {
			$this->sheet->getRowDimension($this->currentRow)->setRowHeight($this->rowHeight);
		}
		// 行号自增
		$this->currentRow++;
	}

	/**
	 * 关闭输出
	 *
	 * @return void
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function close()
	{
		$writer = new Xlsx($this->spreadsheet);
		// 浏览器下载模式
		if (empty($this->saveDir)) {
			header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
			header("Content-Disposition: attachment;filename=\"{$this->fileName}\"");
			header('Cache-Control: max-age=0');
			$writer->save('php://output');
			exit;
		} else {
			// 文件保存模式
			$writer->save($this->getFilePath());
		}
	}

	// 获取文件路径
	public function getFilePath(): string
	{
		return $this->saveDir . DIRECTORY_SEPARATOR . $this->fileName;
	}

	/**
	 * 设置图片属性
	 *
	 * @param  array   $imgColumns   // 图片列名 集合：数据中的关联索引/没索引则为下标
	 * @param  string  $imgEmptyMsg  // 图片路径为空时的提示信息
	 * @param  int     $w            // 图片 宽
	 * @param  int     $h            // 图片 高
	 * @param  int     $rowH         // 行高
	 *
	 * @return void
	 */
	public function setImageAttr(array $imgColumns, string $imgEmptyMsg = '', int $w = 60, int $h = 60, int $rowH = 70)
	{
		$this->imgColumns  = $imgColumns;
		$this->imgEmptyMsg = $imgEmptyMsg;
		$this->imgWidth    = $w;
		$this->imgHeight   = $h;
		$this->rowHeight   = $rowH;
	}

	/**
	 * 插入图片到单元格
	 *
	 * @param $columnLetter  // 列字母
	 * @param $imgPath       // 图片路径
	 *
	 * @return void
	 * @throws \Exception
	 */
	private function insertImage($columnLetter, $imgPath)
	{
		if (empty($imgPath) || ! file_exists($imgPath)) {
			throw new \Exception('文件不存在');
		}
		$draw = new Drawing();
		// 设置单元格位置
		$draw->setCoordinates("{$columnLetter}{$this->currentRow}");
		// 图片路径
		$draw->setPath($imgPath);
		$draw->setWidth($this->imgWidth);
		$draw->setHeight($this->imgHeight);
		// 单元格内部横向偏移
		$draw->setOffsetX(5);
		// 单元格内部纵向偏移
		$draw->setOffsetY(5);
		// 绑定图片对象到工作表，挂载生效（不写这张图不会显示）
		$draw->setWorksheet($this->sheet);
		unset($draw);
	}

	/**
	 * 列下标转Excel列字母 0=A,1=B...
	 *
	 * @param  int  $num  // 列下标
	 *
	 * @return string
	 */
	private function numToLetter(int $num): string
	{
		$letter = '';
		while (true) {
			$mod    = $num % 26;
			$letter = chr(ord('A') + $mod) . $letter;
			$num    = intval($num / 26) - 1;
			if ($num < 0) {
				break;
			}
		}

		return $letter;
	}

	// 关闭句柄
	public function __destruct()
	{
		$this->close();
	}

}