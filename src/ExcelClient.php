<?php
/**
 * User: ogenes<ogenes.yi@gmail.com>
 * Date: 2022/6/17
 */

namespace Ogenes\Exceler;

use Exception;
use InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Chart\Properties;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

class ExcelClient extends Base
{
    public function read(string $filePath, array $config = null, $ext = null): array
    {
        if (empty($config)) {
            throw new InvalidArgumentException('错误的excel配置信息');
        }
        $ext || $ext = pathinfo($filePath, PATHINFO_EXTENSION);
        
        $sheets = [];
        try {
            $ext === 'xlsx' ?
                $reader = IOFactory::createReader('Xlsx') :
                $reader = IOFactory::createReader('Xls');
            $reader->setReadDataOnly(true);  // 设置为只读
            $excel = $reader->load($filePath);
            foreach ($config as $name => $cells) {
                $sheets[$name] = is_int($name) ? $excel->getSheet($name) : $excel->getSheetByName($name);
                if (is_null($sheets[$name])) {
                    throw new InvalidArgumentException('该excel文件中找不到[' . $name . ']表');
                }
            }
        } catch (Exception $e) {
            throw new InvalidArgumentException('读取文件内容失败。详细错误消息:' . $e->getMessage());
        }
        
        $rowSet = [];
        foreach ($sheets as $name => $sheet) {
            $cells = array_flip($config[$name]);
            $rowSet[$name] = $fields = [];
            $i = 0;
            foreach ($sheet->getRowIterator() as $rowIterator) {
                $cellIterator = $rowIterator->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(false);
                if ($i === 0) {
                    $j = 0;
                    foreach ($cellIterator as $cell) {
                        $v = trim($cell->getValue());
                        if (!empty($v)) {
                            $fields[$j] = $v;
                        }
                        $j++;
                    }
                    $diff = array_diff(array_keys($cells), $fields);
                    if (!empty($diff)) {
                        foreach ($diff as $k => $v) {
                            $column = $k + 1;
                            $diff[$k] = "第" . $column . "列 " . $v;
                        }
                        throw new InvalidArgumentException("[{$name}]表中不存在列:" . implode(',', $diff));
                    }
                    if (count($fields) !== count(array_unique($fields))) {
                        throw new InvalidArgumentException("[{$name}]表中存在重复的列名");
                    }
                } else {
                    $row = [];
                    $empty = true;
                    $j = 0;
                    foreach ($cellIterator as $cell) {
                        $v = trim($cell->getValue());
                        if (isset($fields[$j], $cells[$fields[$j]])) {
                            $row[$cells[$fields[$j]]] = $v;
                            if (!empty($v)) {
                                $empty = false;
                            }
                        }
                        $j++;
                    }
                    if (!$empty) {
                        $rowSet[$name][] = $row;
                    }
                }
                $i++;
            }
        }
        return $rowSet;
    }
    
    public function readRemote(string $uri, array $config = null, $ext = null): array
    {
        empty($ext) && $ext = 'xlsx';
        $tmpPath = "/tmp/" . md5($uri . uniqid('read_remote_', true)) . '.' . 'xlsx';
        $filepath = DownloadClient::getInstance()->downloadFile($uri, $tmpPath);
        $ret = $this->read($filepath, $config, $ext);
        @unlink($filepath);
        return $ret;
    }
    
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function export(string $filename, array $config, array $data, string $filepath = './'): string
    {
        if (empty($config)) {
            throw new InvalidArgumentException("错误的excel配置信息'");
        }
        $excel = new Spreadsheet();
        $sheetIndexMap = [];
        $sheetIndex = 0;
        foreach ($config as $sheetName => $sheetConfig) {
            $excel->setActiveSheetIndex($sheetIndex);
            $sheet = $excel->getActiveSheet();
            $sheet->setTitle($sheetName);
            $sheetIndexMap[$sheetIndex] = $sheetName;
            
            $columnIndex = 'A';
            foreach ($sheetConfig as $columnItem) {
                $sheet->setCellValue($columnIndex . '1', $columnItem['columnName']);
                $styleArray = [
                    'borders' => [
                        'outline' => [
                            'borderStyle' => Border::BORDER_THIN,
                            'color' => ['argb' => '00595959'],
                        ],
                    ],
                    'font' => [
                        'name' => '微软雅黑',
                        'color' => ['argb' => '5E5E5E'],
                        'bold' => true,
                        'size' => 8
                    ],
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_CENTER,
                        'vertical' => Alignment::VERTICAL_CENTER,
                        'wrapText' => false,
                    ]
                ];
                $sheet->getStyle($columnIndex . '1')->applyFromArray($styleArray);
                
                //设置单元格宽
                $cellLength = $columnItem['width'] ?? 30;
                $sheet->getColumnDimension($columnIndex)->setWidth($cellLength);
                
                $columnIndex++;
            }
            $excel->createSheet();
            $sheetIndex++;
        }
        
        $sheetIndex = 0;
        foreach ($data as $sheetData) {
            $excel->setActiveSheetIndex($sheetIndex);
            $sheet = $excel->getActiveSheet();
            if (isset($sheetIndexMap[$sheetIndex])) {
                $configKey = $sheetIndexMap[$sheetIndex];
                $sheetConfig = $config[$configKey];
                $rowIndex = 2;
                $maxColumn = 'A';
                $maxRow = 2;
                foreach ($sheetData as $row) {
                    $columnIndex = 'A';
                    foreach ($sheetConfig as $item) {
                        $text = array_key_exists($item['bindKey'], $row) ? $row[$item['bindKey']] : '未找到的bindKey';
                        $sheet->setCellValue($columnIndex . $rowIndex, $text);
                        $setWrapText = $item['setWrapText'] ?? false;
                        $styleArray = [
                            'borders' => [
                                'outline' => [
                                    'borderStyle' => Border::BORDER_THIN,
                                    'color' => ['argb' => '00595959'],
                                ],
                            ],
                            'font' => [
                                'name' => '微软雅黑',
                                'color' => ['argb' => '00595959'],
                                'size' => 9
                            ],
                            'numberFormat' => [
                                'formatCode' => $item['format'] ?? Properties::FORMAT_CODE_GENERAL,
                            ],
                            'alignment' => [
                                'horizontal' => $item['align'] ?? Alignment::HORIZONTAL_LEFT,
                                'vertical' => Alignment::VERTICAL_CENTER,
                            ]
                        ];
                        $setWrapText && $styleArray['alignment']['wrapText'] = true;
                        $sheet->getStyle($columnIndex . $rowIndex)->applyFromArray($styleArray);
                        isset($item['height']) && $sheet->getRowDimension($rowIndex)->setRowHeight($item['height']);
                        $maxColumn = $columnIndex;
                        $columnIndex++;
                    }
                    $maxRow = $rowIndex;
                    $rowIndex++;
                }
                $sheet->setAutoFilter("A1:{$maxColumn}{$maxRow}");
                $sheetIndex++;
            }
        }
        if (ob_get_length() > 0) {
            ob_end_clean();
        }
        ob_start();
        $dir = $filepath . date('/Y/m/d/');
        if (!is_dir($dir) && !mkdir($dir, 0700, true) && !is_dir($dir)) {
            throw new InvalidArgumentException(sprintf('Directory "%s" was not created', $dir));
        }
        $filePath = $dir . $filename . '.xlsx';
        $writer = IOFactory::createWriter($excel, 'Xlsx');
        $writer->save($filePath);
        return $filePath;
    }
    
}
