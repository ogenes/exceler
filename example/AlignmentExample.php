<?php

use Ogenes\Exceler\ExportClient;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

require "vendor/autoload.php";

/**
 * User: ogenes <ogenes.yi@gmail.com>
 * Date: 2022/6/19
 */
class AlignmentExample
{
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function run(): string
    {
        $alignment = [];
        $alignment['shrinkToFit'] = true;

        $data['sheet1'] = [
            ['goodsName' => '半裙', 'price' => 1490, 'actualStock' => 2, 'remark' => "像裙子\r\n像连衣裙\r\n像连衣裙\r\n像连衣裙\r\n像连衣裙"],
            ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1, 'remark' => "像裙子\r\n像连衣裙"]
        ];
        $config['sheet1'] = [
            ['bindKey' => 'goodsName', 'columnName' => '商品名称'],
            ['bindKey' => 'price', 'columnName' => '售价'],
            ['bindKey' => 'actualStock', 'columnName' => '实际库存'],
            ['bindKey' => 'remark', 'columnName' => '备注', 'style' => ['alignment' => $alignment] ],
        ];
        $client = ExportClient::getInstance();
        
        return $client->setFilepath(__DIR__ . '/file/' . date('Y/m/d/'))
            ->setFilename('alignmentDemo' . date('His'))
            ->setData($data)
            ->setConfig($config)
            ->export();
    }
}


try {
    $filepath = (new AlignmentExample())->run();
    print_r($filepath);
} catch (\Exception $e) {
    print_r($e->getMessage());
}
echo PHP_EOL;