<?php

namespace zhuoxin\excel;

// CSV 导出类
// 大数据专用，若导出图片使用ExcelExport类
// 下载文件后缀是 .xlsx，双击直接用 Excel 打开，无任何提示；
// 本质是 CSV 文件，但 Excel 完美兼容，用户完全无感知，这是生产环境的「障眼法」
class CsvExport
{

	// 输出文件句柄
	private $fileHandler;

	// 保存文件路径
	private string $filePath;

	// 是否浏览器模式,默认false
	private bool $isBrowserMode = false;

	/**
	 * 导出Csv类
	 *
	 * @param  string  $fileName    // 导出文件名
	 * @param  string  $saveDir     // 保存目录，为空则浏览器下载
	 * @param  bool    $dateSuffix  // 生成时间后缀，默认true
	 */
	public function __construct(string $fileName, string $saveDir = '', bool $dateSuffix = true)
	{
		// 清除文件名后缀，避免重复.xlsx
		$baseName = rtrim($fileName, '.xlsx');
		// 时间后缀
		$suffix = $dateSuffix ? '_' . date('YmdHis') : '';
		// 生成文件名+年月日时分秒.xlsx
		$fileName = $baseName . $suffix . '.xlsx';

		set_time_limit(0);
		// 取消执行时间限制
		ini_set('max_execution_time', 0);
		// 防止用户关闭浏览器中断导出
		ignore_user_abort(true);

		// 浏览器下载模式
		if (empty($saveDir)) {
			header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
			header("Content-Disposition: attachment; filename={$fileName}");
			header('Cache-Control: no-cache');
			header('Pragma: no-cache');
			header('Expires: 0');
			// 打开文件句柄 输出流
			$this->fileHandler = fopen('php://output', 'w');
		} else {
			// 文件保存模式
			$this->isBrowserMode = false;
			$saveDir             = rtrim($saveDir, '/');
			if ( ! is_dir($saveDir)) {
				mkdir($saveDir, 0755, true);
			}
			$this->filePath = $saveDir . DIRECTORY_SEPARATOR . $fileName;
			// 打开文件句柄 文件流
			$this->fileHandler = fopen($this->filePath, 'w');
		}
		// UTF8 BOM头 解决Excel中文乱码
		fwrite($this->fileHandler, chr(0xEF) . chr(0xBB) . chr(0xBF));
	}

	/**
	 * 设置表格头
	 *
	 * @param  array  $header  ['爱好', '姓名']
	 *
	 * @return void
	 */
	public function setHeader(array $header = [])
	{
		// 写入文件头
		$this->addData($header);
	}

	/**
	 * 写入行数据
	 *
	 * @param $rowData  // 写入行数据 ['音乐','小明']/ ['like'=>'音乐','name'=>'小明']
	 *
	 * @return void
	 */
	public function addData(array $rowData)
	{
		// 如果长数字显示字符串，在传入数字时需加引号
		fputcsv($this->fileHandler, $rowData);
	}

	// 获取文件路径（仅文件保存模式有效）
	public function getFilePath(): string
	{
		return $this->filePath;
	}

	// 主动关闭句柄
	public function close()
	{
		if ($this->fileHandler) {
			// 浏览器模式：需要刷新缓冲区和exit
			if ($this->isBrowserMode) {
				// 强制刷新缓冲区，确保所有数据都输出完毕
				ob_flush();
				flush();
			}
			fclose($this->fileHandler);
			$this->fileHandler = null;
			if ($this->isBrowserMode) {
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