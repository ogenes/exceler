<?php

use Ogenes\Exceler\ExportClient;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

require "vendor/autoload.php";

/**
 * User: ogenes <ogenes.yi@gmail.com>
 * Date: 2022/6/19
 */
class FillExample
{
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function run(): string
    {
        $fill = [
            'fillType' => Fill::FILL_GRADIENT_LINEAR,
            'rotation' => 90,
            'startColor' => [
                'argb' => 'FFA0A0A0',
            ],
            'endColor' => [
                'argb' => 'FFFFFFFF',
            ]
        ];

        $data['sheet1'] = [
            ['goodsName' => '半裙', 'price' => 1490, 'actualStock' => 2],
            ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1]
        ];
        $config['sheet1'] = [
            ['bindKey' => 'goodsName', 'columnName' => '商品名称'],
            ['bindKey' => 'price', 'columnName' => '售价', 'style' => ['fill' => $fill]],
            ['bindKey' => 'actualStock', 'columnName' => '实际库存'],
        ];
        $client = ExportClient::getInstance();
        
        return $client->setFilepath(__DIR__ . '/file/' . date('Y/m/d/'))
            ->setFilename('fillDemo' . date('His'))
            ->setData($data)
            ->setConfig($config)
            ->export();
    }
}


try {
    $filepath = (new FillExample())->run();
    print_r($filepath);
} catch (\Exception $e) {
    print_r($e->getMessage());
}
echo PHP_EOL;