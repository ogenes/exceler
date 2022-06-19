<?php

use Ogenes\Exceler\ExportClient;

require "vendor/autoload.php";

/**
 * User: ogenes <ogenes.yi@gmail.com>
 * Date: 2022/6/19
 */
class FontExample
{
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function run(): string
    {
        $font = [];
        $font['name'] = '宋体';
        $font['size'] = 14;
//        仅对指定行的实际库存生效
        $font['superscript'] = true;
        $style['actualStock']['font'] = $font;
        //对指定行所有字段生效
//        $style['*']['font'] = $font;
        $data['sheet1'] = [
            ['goodsName' => '半裙', 'price' => 1490, 'actualStock' => 2,],
            ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1, 'cellStyle' => $style ]
        ];
        
        $config['sheet1'] = [
            ['bindKey' => 'goodsName', 'columnName' => '商品名称'],
            ['bindKey' => 'price', 'columnName' => '售价'],
            ['bindKey' => 'actualStock', 'columnName' => '实际库存'],
        ];
        $client = ExportClient::getInstance();
//        $client->setStyleHeaderFont($font);
//        $client->setStyleFont($font);
        
        return $client->setFilepath(__DIR__ . '/file/' . date('Y/m/d/'))
            ->setFilename('fontDemo' . date('His'))
            ->setData($data)
            ->setConfig($config)
            ->export();
    }
}


try {
    $filepath = (new FontExample())->run();
    print_r($filepath);
} catch (\Exception $e) {
    print_r($e->getMessage());
}
echo PHP_EOL;