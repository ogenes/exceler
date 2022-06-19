<?php

use Ogenes\Exceler\ExportClient;
use PhpOffice\PhpSpreadsheet\Style\Border;

require "vendor/autoload.php";

/**
 * User: ogenes <ogenes.yi@gmail.com>
 * Date: 2022/6/19
 */
class BorderExample
{
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function run(): string
    {
        $borders = [];
        $borders['bottom']['borderStyle'] = Border::BORDER_THIN;
        $borders['bottom']['color'] = ['argb' => 'F56C6C'];
//        仅对指定行的售价生效
//        $style['price']['borders'] = $borders;
        //对指定行所有字段生效
        $style['*']['borders'] = $borders;
        $data['sheet1'] = [
            ['goodsName' => '半裙', 'price' => 1490, 'actualStock' => 2,],
            ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1 ]
        ];
        
        $config['sheet1'] = [
            ['bindKey' => 'goodsName', 'columnName' => '商品名称'],
            ['bindKey' => 'price', 'columnName' => '售价'],
            ['bindKey' => 'actualStock', 'columnName' => '实际库存'],
        ];
        $client = ExportClient::getInstance();
        $client->setStyleHeaderBorders($borders);
        $borders['bottom']['borderStyle'] = Border::BORDER_DASHDOT;
        $client->setStyleBorders($borders);
        
        return $client->setFilepath(__DIR__ . '/file/' . date('Y/m/d/'))
            ->setFilename('borderDemo' . date('His'))
            ->setData($data)
            ->setConfig($config)
            ->export();
    }
}


try {
    $filepath = (new BorderExample())->run();
    print_r($filepath);
} catch (\Exception $e) {
    print_r($e->getMessage());
}
echo PHP_EOL;