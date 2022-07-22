<?php

use Ogenes\Exceler\ExportClient;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;

require "vendor/autoload.php";

/**
 * User: ogenes <ogenes.yi@gmail.com>
 * Date: 2022/7/23
 */
class DrawExample
{
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function run(): string
    {
        ini_set('memory_limit', '2KB');
        $data['sheet1'] = [
            ['goodsName' => ['text' => '半裙', 'comment' => 'this is comment'], 'price' => 1490, 'actualStock' => 2, 'img' => ''],
            ['goodsName' => '半裙1', 'price' => 1590, 'actualStock' => 1, 'img'=> 'https://static-legacy.dingtalk.com/media/lADPD3zUN8-C7xXNAzzNAzw_828_828.jpg'],
            ['goodsName' => '半裙', 'price' => 899, 'actualStock' => 2, 'img'=> 'https://static-legacy.dingtalk.com/media/lADPBbCc1reZWcPNAgPNAes_491_515.jpg'],
        ];
    
        $config['sheet1'] = [
            ['bindKey' => 'goodsName', 'columnName' => '商品名称', 'width' => 40 ],
            ['bindKey' => 'img', 'columnName' => '图片','width' => 20, 'drawing' => ['w' => 80, 'h' => 80, 'remote' => true]],
            ['bindKey' => 'price', 'columnName' => '售价', 'format' => NumberFormat::FORMAT_DATE_YYYYMMDDSLASH],
            ['bindKey' => 'actualStock', 'columnName' => '实际库存'],
        ];
        $client = ExportClient::getInstance();
        return $client->setFilepath(__DIR__ . '/file/' . date('Y/m/d/'))
            ->setFilename('ImgDemo' . date('His'))
            ->setData($data)
            ->setConfig($config)
            ->export();
    }
}

try {
    $filepath = (new DrawExample())->run();
    print_r($filepath);
} catch (\Exception $e) {
    print_r($e->getMessage());
}
echo PHP_EOL;