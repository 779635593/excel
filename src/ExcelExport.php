<?php

namespace zhuoxin\excel;

// Excel 导出类，大数据导出使用
// 下载文件后缀是 .xlsx，双击直接用 Excel 打开，无任何提示；
// 本质是 CSV 文件，但 Excel 完美兼容，用户完全无感知，这是生产环境的「障眼法」
class ExcelExport
{

	// 输出文件句柄
	private $file_handler = null;

	// 保存文件路径
	private $file_path = null;

	// 是否浏览器模式,默认false
	private $is_browser_mode = false;

	/**
	 * 导出Csv类
	 *
	 * @param  string  $filename   // 文件名
	 * @param  string  $save_path  // 保存路径，为空则浏览器下载
	 */
	public function __construct(string $filename = "", string $save_path = "")
	{
		// 生成文件名+年月日时分秒.xlsx
		$filename = $filename . date('YmdHis') . '.xlsx';;
		set_time_limit(0);
		// 取消执行时间限制
		ini_set('max_execution_time', 0);
		// 防止用户关闭浏览器中断导出
		ignore_user_abort(true);

		// 浏览器下载模式
		if (empty($save_path)) {
			header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
			header("Content-Disposition: attachment; filename={$filename}");
			header('Cache-Control: no-cache');
			header('Pragma: no-cache');
			header('Expires: 0');
			// 打开文件句柄 输出流
			$this->file_handler = fopen('php://output', 'w');
		} else {
			// 文件保存模式
			$this->is_browser_mode = false;
			$save_path             = rtrim($save_path, '/');
			if ( ! is_dir($save_path)) {
				mkdir($save_path, 0755, true);
			}
			$this->file_path = $save_path . '/' . $filename;
			// 打开文件句柄 文件流
			$this->file_handler = fopen($this->file_path, 'w');
		}
		// UTF8 BOM头 解决Excel中文乱码
		fwrite($this->file_handler, chr(0xEF) . chr(0xBB) . chr(0xBF));
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
			// 长数字转为字符串
			if (is_numeric($value) && strlen((string) $value) > 8) {
				// 用="包裹数字，强制Excel以文本显示
				$value = '="' . $value . '"';
			}
		}
		fputcsv($this->file_handler, $data);
	}

	/**
	 * 获取文件路径（仅文件保存模式有效）
	 */
	public function getFilePath(): string
	{
		return $this->file_path;
	}

	// 主动关闭句柄
	public function close()
	{
		if ($this->file_handler) {
			// 浏览器模式：需要刷新缓冲区和exit
			if ($this->is_browser_mode) {
				// 强制刷新缓冲区，确保所有数据都输出完毕
				ob_flush();
				flush();
			}
			fclose($this->file_handler);
			$this->file_handler = null;
			if ($this->is_browser_mode) {
				exit;
			}
		}
	}

	// 关闭句柄
	public function __destruct()
	{
		$this->close();
	}

}