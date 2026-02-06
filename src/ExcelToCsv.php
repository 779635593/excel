<?php

namespace zhuoxin\excel;

// Excel 导出类，超大数据导出使用
// 下载文件后缀是 .xlsx，双击直接用 Excel 打开，无任何提示；
// 本质是 CSV 文件，但 Excel 完美兼容，用户完全无感知，这是生产环境的「障眼法」
class ExcelToCsv
{

	// 输出文件句柄
	private $file_handler = null;

	/**
	 * 导出Csv类
	 *
	 * @param  string  $filename  // 文件名
	 */
	public function __construct(string $filename = "")
	{
		$filename = $filename . date('YmdHis');
		set_time_limit(0);
		// 64M足够，百万行也够用
		ini_set('memory_limit', '64M');
		// 取消执行时间限制
		ini_set('max_execution_time', 0);
		// 防止用户关闭浏览器中断导出
		ignore_user_abort(true);

		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header("Content-Disposition: attachment; filename={$filename}.xlsx");
		header('Cache-Control: no-cache');
		header('Pragma: no-cache');
		header('Expires: 0');

		// 打开文件句柄
		$this->file_handler = fopen('php://output', 'w');
		// UTF8 BOM头 解决Excel中文乱码，两种写法都正确，任选其一即可
		fwrite($this->file_handler, chr(0xEF) . chr(0xBB) . chr(0xBF));
		// fwrite($this->file_handler, chr(239).chr(187).chr(191));
	}

	/**
	 * 设置表格头
	 *
	 * @param  array  $header  ['爱好', '姓名']
	 */
	public function setHeader(array $header = [])
	{
		// 写入文件头
		if ($header) {
			$this->addData($header);
		}
	}

	/**
	 * 写入行数据
	 *
	 * @param $data  // 写入行数据 ['音乐','小明']/ ['like'=>'音乐','name'=>'小明']
	 *
	 * @return false|void
	 */
	public function addData($data)
	{
		if ( ! is_array($data) || empty($data)) {
			return false;
		}
		// 强制转为【纯值的索引数组】，兼容关联数组/索引数组
		$data = array_values($data);
		// 循环处理每个单元格数据
		foreach ($data as &$value) {
			// 空值处理：null/false 转为空字符串
			if ($value === null || $value === false) {
				$value = '';
			}
			// 长数字处理：手机号/身份证号(≥5位纯数字) 加制表符，Excel强制识别为文本，防科学计数法
			if (is_numeric($value) && strlen((string) $value) >= 5) {
				$value = "\t" . $value;
			}
			// 去除首尾空格，避免脏数据
			$value = trim((string) $value);
		}
		fputcsv($this->file_handler, $data);
	}

	// 主动关闭句柄
	public function close()
	{
		if ($this->file_handler) {
			// 强制刷新缓冲区，确保所有数据都输出完毕
			ob_flush();
			flush();
			fclose($this->file_handler);
			$this->file_handler = null;
			exit;
		}
	}

	// 关闭句柄
	public function __destruct()
	{
		$this->close();
	}

}