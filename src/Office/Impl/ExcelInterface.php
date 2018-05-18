<?php
namespace App\Office\Impl;


interface ExcelInterface
{
    // 导入
    public function import($originFile);

    /**
     * 导出excel
     *
     * @param array   $exportOptions 导出数据设置
     *
     * @param string  $filename      导出文件名
     * @param string  $targetDir     导出目录(默认直接下载)
     * @param boolean $debug         是否调试(直接显示在浏览器上)
     */
    public function export(array $exportOptions, $filename = '', $targetDir = '', $debug = false);
}