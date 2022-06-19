<?php

use Ogenes\Exceler\ExportClient;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

require "vendor/autoload.php";

/**
 * User: ogenes <ogenes.yi@gmail.com>
 * Date: 2022/6/19
 */
class CellExample
{
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function run(): string
    {
        $data['sheet1'] = [
            ['goodsName' => '半裙', 'price' => 1490, 'actualStock' => 2, 'img' => __DIR__ . '/file/example.jpeg'],
            ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1, 'img' => __DIR__ . '/file/example.jpeg']
        ];
        
        $config['sheet1'] = [
            ['bindKey' => 'goodsName', 'columnName' => '商品名称' ],
            ['bindKey' => 'img', 'columnName' => '图片', 'drawing' => ['w' => 80, 'h' => '80'] ],
            ['bindKey' => 'price', 'columnName' => '售价', 'format' => NumberFormat::FORMAT_DATE_YYYYMMDDSLASH],
            ['bindKey' => 'actualStock', 'columnName' => '实际库存'],
            ['bindKey' => '={price}*{actualStock}', 'columnName' => '库存额'], //会自动转化为公式
        ];
        $client = ExportClient::getInstance();
        return $client->setFilepath(__DIR__ . '/file/' . date('Y/m/d/'))
            ->setFilename('cellDemo' . date('His'))
            ->setData($data)
            ->setConfig($config)
            ->export();
    }
}


try {
    $filepath = (new CellExample())->run();
    print_r($filepath);
} catch (\Exception $e) {
    print_r($e->getMessage());
}
echo PHP_EOL;