<?php
/**
 * User: ogenes<ogenes.yi@gmail.com>
 * Date: 2022/6/18
 */

namespace Ogenes\Exceler;

use InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class ExportClient extends ExportService
{
    protected $config = [];
    protected $data = [];
    
    /**
     * @desc 请指定导出配置
     *
     * [
     * 'bindKey' => 'price', // 字段名
     * 'columnName' => '售价', // 表头名
     * 'horizontal' => Alignment::HORIZONTAL_RIGHT, //水平方向，参考 PhpOffice\PhpSpreadsheet\Style\Alignment
     * 'format' => Properties::FORMAT_CODE_ACCOUNTING, // 格式，参考 PhpOffice\PhpSpreadsheet\Chart\Properties
     * 'width' => 20 // 宽度
     * ]
     * @param array $config
     * @return $this
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setConfig(array $config): self
    {
        $this->config = $config;
        return $this;
    }
    
    /**
     * @return array
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function getConfig(): array
    {
        return $this->config;
    }
    
    /**
     * @desc 请指定导出的数据内容
     *
     * @param array $data
     * @return $this
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function setData(array $data): self
    {
        $this->data = $data;
        return $this;
    }
    
    /**
     * @return array
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function getData(): array
    {
        return $this->data;
    }
    
    /**
     * @desc 导出excel表格到文件
     * @return string
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function export(): string
    {
        $config = $this->getConfig();
        if (empty($config)) {
            throw new InvalidArgumentException("错误的excel配置信息'");
        }
        $excel = new Spreadsheet();
        $sheetIndexMap = $this->formatHeader($excel);
        $this->formatContent($excel, $sheetIndexMap);
        if (ob_get_length() > 0) {
            ob_end_clean();
        }
        ob_start();
        $dir = $this->getFilepath();
        if (!is_dir($dir) && !mkdir($dir, 0700, true) && !is_dir($dir)) {
            throw new InvalidArgumentException(sprintf('Directory "%s" was not created', $dir));
        }
        $filePath = $dir . '/' . $this->getFilename() . '.xlsx';
        $writer = IOFactory::createWriter($excel, 'Xlsx');
        $writer->save($filePath);
        return $filePath;
    }
    
    /**
     * @desc 导出excel表格到浏览器
     *
     * @return void
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    public function output()
    {
        $config = $this->getConfig();
        if (empty($config)) {
            throw new InvalidArgumentException("错误的excel配置信息'");
        }
        $excel = new Spreadsheet();
        $sheetIndexMap = $this->formatHeader($excel);
        $this->formatContent($excel, $sheetIndexMap);
        if (ob_get_length() > 0) {
            ob_end_clean();
        }
        ob_start();
        
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header(sprintf('Content-Disposition: attachment;filename="%s.xlsx"', $this->getFilename()));
        header('Cache-Control: max-age=0');
        
        $writer = IOFactory::createWriter($excel, 'Xlsx');
        $writer->save('php://output');
    }
    
    /**
     * @desc 生成表头
     *
     * @param Spreadsheet $excel
     * @return array
     * @throws Exception
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    protected function formatHeader(Spreadsheet $excel): array
    {
        $config = $this->getConfig();
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
                $styleArray = [];
                $this->headerFont && $styleArray['font'] = $this->headerFont;
                $this->headerBorders && $styleArray['borders'] = $this->headerBorders;
                $this->headerAlignment && $styleArray['alignment'] = $this->headerAlignment;
                $this->headerFill && $styleArray['fill'] = $this->headerFill;
                $styleArray && $sheet->getStyle($columnIndex . '1')->applyFromArray($styleArray);
                $cellLength = $columnItem['width'] ?? $this->width;
                $sheet->getColumnDimension($columnIndex)->setWidth($cellLength, $this->unit);
                
                $columnIndex++;
            }
            $excel->createSheet();
            $sheetIndex++;
        }
        return $sheetIndexMap;
    }
    
    /**
     * @desc 生成内容
     * @param Spreadsheet $excel
     * @param array $sheetIndexMap
     * @return void
     * @throws Exception
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    protected function formatContent(Spreadsheet $excel, array $sheetIndexMap)
    {
        $config = $this->getConfig();
        $data = $this->getData();
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
                $columnMap = [];
                foreach ($sheetData as $row) {
                    $columnIndex = 'A';
                    foreach ($sheetConfig as $item) {
                        $hyperlink = '';
                        if (array_key_exists($item['bindKey'], $row)) {
                            $text = $row[$item['bindKey']];
                            if (is_array($text)) {
                                $hyperlink = $text['hyperlink'] ?? '';
                                $text = $text['text'] ?? '';
                            }
                        } elseif (strpos($item['bindKey'], '=') !== false) {
                            $text = str_replace(
                                array_map(static function ($key) {
                                    return "{{$key}}";
                                }, array_keys($columnMap)),
                                array_map(static function ($val) use ($rowIndex) {
                                    return "{$val}{$rowIndex}";
                                }, $columnMap),
                                $item['bindKey']
                            );
                        } else {
                            $text = '未找到的bindKey';
                        }
                        $sheet->setCellValue($columnIndex . $rowIndex, $text);
                        $styleArray = $this->getContentStyle($item, $row['cellStyle'] ?? []);
                        $styleArray && $sheet->getStyle($columnIndex . $rowIndex)->applyFromArray($styleArray);
                        $hyperlink && $sheet->getCell($columnIndex . $rowIndex)->getHyperlink()->setUrl($hyperlink);
                        $maxColumn = $columnIndex;
                        $columnMap[$item['bindKey']] = $columnIndex;
                        $columnIndex++;
                    }
                    $maxRow = $rowIndex;
                    $rowIndex++;
                }
//                $sheet->setAutoFilter("A1:{$maxColumn}{$maxRow}");
                $sheetIndex++;
            }
        }
    }
    
    /**
     * @desc 生成内容格式
     *
     * @param array $item
     * @return array
     *
     * @author: ogenes<ogenes.yi@gmail.com>
     * @date: 2022/6/18
     */
    protected function getContentStyle(array $item, array $columnStyle): array
    {
        $styleArray = [
            'font' => $this->font,
            'borders' => $this->borders,
            'alignment' => $this->alignment,
            'fill' => $this->fill,
            'numberFormat' => ['formatCode' => $this->formatCode],
        ];
        
        $existStyle = $item['style'] ?? [];
        $fields = [
            'font',
            'borders',
            'alignment',
            'fill',
        ];
        foreach ($fields as $field) {
            !empty($existStyle[$field]) && $styleArray[$field] = array_merge($styleArray[$field], $existStyle[$field]);
            $cellStyle = $columnStyle[$item['bindKey']] ?? ($columnStyle['*'] ?? []);
            !empty($cellStyle[$field]) && $styleArray[$field] = array_merge($styleArray[$field], $cellStyle[$field]);
        }
        if (!empty($item['format'])) {
            $styleArray['numberFormat']['formatCode'] = $item['format'];
        }
        return $styleArray;
    }
    
}