<?php

namespace App\Office;

use App\Office\Impl\ExcelInterface;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Writer\Html;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf;

class Excel implements ExcelInterface
{
    /**
     * @var array 传入数据规范
     */
    private $exportOptions = [
        // 导出字段
        'fields' => [],
        // 导出数据
        'data'   => [],
        //
        'style'  => [
        ],
        'title'  => '', // excel 标题
    ];

    private $defaultOptions = [
        'startCell'  => 'A1',
        'title'      => '', // excel 标题
        'subTitle'   => '',
        'fontFamily' => '楷体',
        'height'     => 30,
        'width'      => 30,     // 单元格宽度
        'template'   => '',
        'merges'     => [
            'cell'     => [],
            'cellText' => [],
        ],
        'column'     => [
            'A' => '00000000', 'B' => 'FF000000', 'D' => 'FF0FF000'
        ],
        'style'      => [
            'title'   => [
                // excel 要求格式
                'font'      => ['bold'  => true,
                                'size'  => 16,
                                'color' => [
                                    'argb' => 'FF25281B'
                                ]
                ],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                    'vertical'   => Alignment::VERTICAL_CENTER
                ],
                'rowHeight' => '30'
            ],
            'fields'  => [
                // excel 要求格式
                'font'      => ['bold'  => true,
                                'size'  => 14,
                                'color' => [
                                    'argb' => 'FF25281B'
                                ]
                ],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                    'vertical'   => Alignment::VERTICAL_CENTER
                ],
                'rowHeight' => '35'
            ],
            'content' => [
                // excel 要求格式
                'font'      => [
                    'bold'  => false,
                    'size'  => 12,
                    'color' => [
                        'argb' => '00000000',
                    ]
                ],
                'alignment' => [
                    'horizontal' => Alignment::HORIZONTAL_CENTER,
                ],
                'rowHeight' => '30',   // 行高
            ],
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,  //设置border样式
                    'color'       => ['argb' => 'C3C3C3C3'],     //设置border颜色
                ],
            ],
            'link'    => [      // 包含链接单元格的宽度
                'width' => 50
            ],
            'image'   => [  // 单元格图片设置
                'height' => 100,
                'width'  => 50
            ]
        ],

        'creator'     => '', // 创建人
        'lastModify'  => '', // 最后修改人.
        'subject'     => '', // 题目
        'description' => '', // 描述
        'keywords'    => '', //关键词
        'category'    => '', // 种类
    ];

    public function __construct($style = [])
    {
        array_walk($this->defaultOptions, function () use ($style) {
            foreach ($style as $key => $roomStyle) {
                if (is_array($roomStyle)) {
                    foreach ($roomStyle as $skey => $subStyle) {
                        if (is_array($subStyle)) {
                            foreach ($subStyle as $k => $v) {
                                $this->defaultOptions[$key][$skey][$k] = $v;
                            }
                        } else {
                            $this->defaultOptions[$key][$skey] = $subStyle;
                        }
                    }
                } else {
                    $this->defaultOptions[$key] = $roomStyle;
                }
            }
        });
    }

    /**
     * 导入excel表格
     *
     * @param $originFile
     *
     * @return array|null
     */
    public function import($originFile)
    {
        if (!file_exists($originFile))
            $result = null;

        try {
            $readerHandel = IOFactory::createReaderForFile($originFile);
            $spreadsheet  = $readerHandel->load($originFile);
            $sheet        = $spreadsheet->getActiveSheet();

            $result = $sheet->toArray();

        } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
            $result = null;
        } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
            $result = null;
        }

        return $result;
    }

    /**
     * 导出excel
     *
     * @param array   $exportOptions 导出数据设置
     *
     * @param string  $filename      导出文件名
     * @param string  $targetDir     导出目录(默认直接下载)
     * @param boolean $debug         是否调试
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function export(array $exportOptions, $filename = '', $targetDir = '', $debug = false)
    {
        $exportOptions = array_merge($this->exportOptions, $exportOptions);

        if ($this->defaultOptions['template']) {
            $spreadsheet = IOFactory::load($this->defaultOptions['template']);
        } else {
            $spreadsheet = new Spreadsheet();
        }

        $fields   = $exportOptions['fields'];
        $rowDatas = $exportOptions['data'];
        list($startColumn, $startRow) = Coordinate::coordinateFromString($this->defaultOptions['startCell']);
        $rowIndex = $startRow;
        // 设置表头标题
        if ($title = $exportOptions['title']) {
            $titleColumn = $startColumn;
            $maxColumn   = count($fields) - 1;
            for ($i = 0; $i < $maxColumn; $i++) {
                ++$titleColumn;
            }

            $spreadsheet
                ->getActiveSheet()
                ->getRowDimension($startRow)
                ->setRowHeight($this->defaultOptions['style']['title']['rowHeight']);

            $spreadsheet
                ->getActiveSheet()
                ->mergeCells($startColumn . $startRow . ':' . $titleColumn . $startRow)// 合并单元格
                ->setCellValue($startColumn . $startRow, $title);
            if (!$this->defaultOptions['template'])
                $spreadsheet->getActiveSheet()
                    ->getStyle($startColumn . $startRow)
                    ->applyFromArray([
                        'font'      => $this->defaultOptions['style']['title']['font'],
                        'alignment' => $this->defaultOptions['style']['title']['alignment'],
                        'borders'   => $this->defaultOptions['style']['borders']
                    ]);

            $rowIndex++;
        }

        // 有合并行
        if ($this->defaultOptions['merges']['cellText']) {
            foreach ($this->defaultOptions['merges']['cell'] as $merge) {
                $spreadsheet->getActiveSheet()->mergeCells($merge);
            }
            $spreadsheet
                ->getActiveSheet()
                ->getRowDimension($rowIndex)
                ->setRowHeight($this->defaultOptions['style']['fields']['rowHeight']);
            foreach ($this->defaultOptions['merges']['cellText'] as $cell => $val) {
                $spreadsheet->getActiveSheet()->setCellValue($cell, $val);
                if (!$this->defaultOptions['template'])
                    $spreadsheet->getActiveSheet()
                        ->getStyle($cell)
                        ->applyFromArray([
                            'font'      => $this->defaultOptions['style']['fields']['font'],
                            'alignment' => $this->defaultOptions['style']['fields']['alignment'],
                            'borders'   => $this->defaultOptions['style']['borders'],
                        ]);
            }
            $rowIndex++;
        }

        // 设置字段高度
        $spreadsheet
            ->getActiveSheet()
            ->getRowDimension($rowIndex)
            ->setRowHeight($this->defaultOptions['style']['fields']['rowHeight']);

        $currentColumn = $startColumn;
        // 设置字段
        foreach ($fields as $k => $v) {
            $cell = $currentColumn . $rowIndex;

            $spreadsheet->getActiveSheet()
                ->setCellValue($cell, $v);
            if (!$this->defaultOptions['template'])
                $spreadsheet->getActiveSheet()
                    ->getStyle($cell)
                    ->applyFromArray([
                        'font'      => $this->defaultOptions['style']['fields']['font'],
                        'alignment' => $this->defaultOptions['style']['fields']['alignment'],
                        'borders'   => $this->defaultOptions['style']['borders']
                    ]);
            // 设置每列的宽度
            $spreadsheet->getActiveSheet()
                ->getColumnDimension($currentColumn)
                ->setWidth($this->defaultOptions['width']);

            ++$currentColumn;
        }

        $fields && $rowIndex++;

        // 设置内容
        foreach ($rowDatas as $rowData) {
            $currentColumn = $startColumn;
            foreach ($rowData as $cellValue) {
                $cell = $currentColumn . $rowIndex;
                $spreadsheet
                    ->getActiveSheet()
                    ->getRowDimension($rowIndex)
                    ->setRowHeight($this->defaultOptions['style']['content']['rowHeight']);

                if (!$this->defaultOptions['template'])
                    $spreadsheet->getActiveSheet()
                        ->getStyle($cell)
                        ->applyFromArray([
                            'font'      => $this->defaultOptions['style']['content']['font'],
                            'alignment' => $this->defaultOptions['style']['content']['alignment'],
                            'borders'   => $this->defaultOptions['style']['borders']
                        ]);

                $this->setCellValues($spreadsheet, $cell, $cellValue ?: '');

                if (array_key_exists($currentColumn, $this->defaultOptions['column']))
                    $spreadsheet->getActiveSheet()
                        ->getStyle($cell)
                        ->getFont()
                        ->getColor()
                        ->setARGB($this->defaultOptions['column'][$currentColumn]);

                $currentColumn++;
            }
            $rowIndex++;
        }

        // 文档属性
        $spreadsheet
            ->getProperties()
            ->setCreator($this->defaultOptions['creator'])
            ->setLastModifiedBy($this->defaultOptions['lastModify'])
            ->setTitle($this->defaultOptions['title'])
            ->setSubject($this->defaultOptions['subject'])
            ->setDescription($this->defaultOptions['description'])
            ->setKeywords($this->defaultOptions['keywords'])
            ->setCategory($this->defaultOptions['category']);
        $this->defaultOptions['subTitle'] && $spreadsheet->getActiveSheet()->setTitle($this->defaultOptions['subTitle']);

        // 默认属性
        $spreadsheet
            ->getDefaultStyle()
            ->getFont()
            ->setName(iconv('gbk', 'utf-8', $this->defaultOptions['fontFamily']))
            ->setSize(0);

        $fileExt = ucfirst(pathinfo($filename, PATHINFO_EXTENSION));

        $this->outPut($spreadsheet, $filename, $targetDir, $fileExt, $debug);
    }

    /**
     * 导出表格
     *
     * @param Spreadsheet $spreadsheet
     * @param string      $filename  导出文件名
     * @param string      $targetDir 导出文件目录
     * @param string      $type      导出类型 excel|pdf
     * @param bool        $debug     是否直接显示在浏览器上
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function outPut(Spreadsheet $spreadsheet, $filename, $targetDir = '', $type = 'excel', $debug = false)
    {
        if ($debug) {
            $objHtmlWriter = new Html($spreadsheet);
            $objHtmlWriter->save("php://output");

            return;
        }

        switch ($type) {
            case $type == 'Xlsx':
                $objWriter = IOFactory::createWriter($spreadsheet, "Xlsx");
                if (!$targetDir) {
                    header("Content-Type: application/vnd.ms-excel;");
                    $filename = $filename ?: time() . '.xlsx';
                    header("Content-Disposition:attachment;filename=" . $filename);
                }
                break;
            case $type == 'Xls':
                $objWriter = IOFactory::createWriter($spreadsheet, "Xls");
                if (!$targetDir) {
                    header("Content-Type: application/vnd.ms-excel;");
                    $filename = $filename ?: time() . '.xls';
                    header("Content-Disposition:attachment;filename=" . $filename);
                }
                break;
            case $type == 'Pdf':
                IOFactory::registerWriter("Pdf", Mpdf::class);
                $objWriter = IOFactory::createWriter($spreadsheet, 'Pdf');
                $objWriter->setSheetIndex(0);
                if (!$targetDir) {
                    $filename = $filename ?: time() . '.pdf';
                    header("Content-Disposition:attachment;filename=" . $filename);
                    header("Content-Type: application/pdf");
                }
                break;
            default:
                throw new \PhpOffice\PhpSpreadsheet\Writer\Exception("Do not support file type");
        }

        if ($targetDir) {
            $objWriter->save($targetDir . '/' . $filename);
        } else {
            header("Pragma: public");
            header("Expires: 0");
            header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
            header("Content-Type:application/force-download");
            header("Content-Type:application/octet-stream");
            header("Content-Type:application/download");
            header("Content-Transfer-Encoding:binary");

            $objWriter->save("php://output");
        }

        $spreadsheet->disconnectWorksheets();
        unset($spreadsheet);
    }

    /**
     * 设置单元格数据
     *
     * @param Spreadsheet $spreadsheet
     * @param string      $cell 单元格
     * @param string      $val  单元格值
     *
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function setCellValues(Spreadsheet $spreadsheet, $cell, $val)
    {
        if ($val) {
            switch (true) {
                // 布尔类型
                case is_bool($val):
                    $this->writeBoolValue($spreadsheet, $cell, $val);
                    break;

                // 对象格式
                case is_object($val):
                    $this->writeObjectValue($spreadsheet, $cell, $val);
                    break;

                // 连接格式
                case is_string($val) && preg_match('/^(http|https)(.*?)\.com|cn|cc$/i', $val):
                    self::writeLinkValue($spreadsheet, $cell, $val);
                    break;

                // 图片格式
                case is_string($val) && preg_match('/^.*?\.jpeg|jpg|png|gif$/i', $val) && file_exists($val):
                    $this->writeImageValue($spreadsheet, $cell, $val);
                    break;

                // 纯文本格式
                default:
                    $this->writeTextValue($spreadsheet, $cell, $val);
            }
        }

        // 内容垂直居中
        $spreadsheet
            ->getActiveSheet()
            ->getStyle($cell)
            ->getAlignment()
            ->setVertical(Alignment::VERTICAL_CENTER);
    }

    /**
     * 设置图片格式数据
     *
     * @param Spreadsheet $spreadsheet
     * @param string      $cell      单元格
     * @param string      $imagePath 图片地址
     * @param string      $imageName 图片名称
     * @param string      $imageDesc 图片描述
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function writeImageValue(Spreadsheet $spreadsheet, $cell, $imagePath, $imageName = '', $imageDesc = '')
    {
        $objDrawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
        $objDrawing
            ->setWorksheet($spreadsheet->getActiveSheet())
            ->setName($imageName)
            ->setDescription($imageDesc)
            ->setCoordinates($cell);

        $objDrawing
            ->setHeight($this->defaultOptions['style']['image']['height'] / 100)// 设置为下面的0.01倍 不知道原因
            ->setWidth($this->defaultOptions['style']['image']['width']);

        $objDrawing->setPath(substr($imagePath, 0, 1) == '/' ? substr($imagePath, 1) : $imagePath, true);
    }

    /**
     * 设置文本格式数据
     *
     * @param Spreadsheet $spreadsheet
     * @param string      $cell 单元格
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function writeTextValue(Spreadsheet $spreadsheet, $cell, $val)
    {
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue($cell, $val);
        if (!$this->defaultOptions['template'])
            $spreadsheet->getActiveSheet()
                ->getStyle($cell)
                ->applyFromArray([
                    'font'      => $this->defaultOptions['style']['content']['font'],
                    'alignment' => $this->defaultOptions['style']['content']['alignment'],
                ]);
    }

    /**
     * 设置对象格式数据
     *
     * @param Spreadsheet $spreadsheet
     * @param string      $cell 单元格
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function writeObjectValue(Spreadsheet $spreadsheet, $cell, $val)
    {
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue($cell, $val);
        if (!$this->defaultOptions['template'])
            $spreadsheet->getActiveSheet()
                ->getStyle($cell)
                ->applyFromArray([
                    'font'      => $this->defaultOptions['style']['content']['font'],
                    'alignment' => $this->defaultOptions['style']['content']['alignment'],
                ]);
    }

    /**
     * 设置布尔格式数据
     *
     * @param Spreadsheet $spreadsheet
     * @param string      $cell 单元格
     * @param string      $val
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function writeBoolValue(Spreadsheet $spreadsheet, $cell, $val)
    {
        $spreadsheet
            ->getActiveSheet()
            ->setCellValue($cell, $val ? '是' : '否');
        if (!$this->defaultOptions['template'])
            $spreadsheet->getActiveSheet()
                ->getStyle($cell)
                ->applyFromArray([
                    'font'      => $this->defaultOptions['style']['content']['font'],
                    'alignment' => $this->defaultOptions['style']['content']['alignment'],
                ]);
    }

    /**
     * 设置连接格式数据
     *
     * @param Spreadsheet $spreadsheet
     * @param string      $cell 单元格
     * @param string      $val
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function writeLinkValue(Spreadsheet $spreadsheet, $cell, $val)
    {
        // 设置该列的宽度
        $spreadsheet->getActiveSheet()
            ->getColumnDimension(substr($cell, 0, 1))
            ->setWidth($this->defaultOptions['style']['link']['width']);

        $spreadsheet->getActiveSheet()
            ->setCellValue($cell, iconv('gbk', 'utf-8', $val))
            ->getCell($cell)
            ->getHyperlink()
            ->setUrl($val);

        // 设置单元格内容格式
        if (!$this->defaultOptions['template'])
            $spreadsheet
                ->getActiveSheet()
                ->getStyle($cell)
                ->applyFromArray([
                    'font'      => $this->defaultOptions['style']['content']['font'],
                    'alignment' => $this->defaultOptions['style']['content']['alignment'],
                ]);
    }

}
