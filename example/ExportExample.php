<?php

use Ogenes\Exceler\ExportClient;
use PhpOffice\PhpSpreadsheet\Chart\Properties;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

require "vendor/autoload.php";

/**
 * User: ogenes<ogenes.yi@gmail.com>
 * Date: 2022/6/17
 */
class ExportExample
{
    
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function run(): string
    {
        $data['sheet1'] = [
            ['goodsName' => '半裙', 'price' => 1490, 'actualStock' => 2,],
            ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1,]
        ];
        
        $config['sheet1'] = [
            ['bindKey' => 'goodsName', 'columnName' => '商品名称'],
            ['bindKey' => 'price', 'columnName' => '售价', 'horizontal' => Alignment::HORIZONTAL_RIGHT, 'format' => Properties::FORMAT_CODE_ACCOUNTING],
            ['bindKey' => 'actualStock', 'columnName' => '实际库存', 'horizontal' => Alignment::HORIZONTAL_RIGHT, 'width' => 15],
        ];
        return ExportClient::getInstance()
            ->setFilepath(__DIR__ . '/file/' . date('Y/m/d/'))
            ->setFilename('file' . date('His'))
            ->setData($data)
            ->setConfig($config)
            ->export();
    }
}

try {
    $filepath = (new ExportExample())->run();
    print_r($filepath);
} catch (\Exception $e) {
    print_r($e->getMessage());
}
echo PHP_EOL;