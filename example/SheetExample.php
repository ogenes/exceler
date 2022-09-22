<?php

use Ogenes\Exceler\ExportClient;

require "vendor/autoload.php";

/**
 * User: ogenes <ogenes.yi@gmail.com>
 * Date: 2022/6/30
 */
class SheetExample
{
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function run(): string
    {
        ini_set('memory_limit', '2KB');
        $data['sheet1'] = [
            ['goodsName' => ['text' => '半裙', 'comment' => 'this is comment'], 'price' => 1490, 'actualStock' => 2],
            ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1, 'hyperlink' => 'https://www.baidu.com/s?wd=半裙2'],
            ['goodsName' => ['text' => '半裙', 'comment' => 'this is comment'], 'price' => 1490, 'actualStock' => 2],
            ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1, 'hyperlink' => 'https://www.baidu.com/s?wd=半裙2'],
            ['goodsName' => ['text' => '半裙', 'comment' => 'this is comment'], 'price' => 1490, 'actualStock' => 2],
            ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1, 'hyperlink' => 'https://www.baidu.com/s?wd=半裙2']
        ];
        for ($i = 0; $i <= 20000; $i++) {
            echo $i . PHP_EOL;
            $data['sheet1'][] = ['goodsName' => ['text' => '半裙', 'comment' => 'this is comment'], 'price' => 1490, 'actualStock' => 2];
        }
        
        $config['sheet1'] = [
            ['bindKey' => 'goodsNme', 'align' => 'left', 'columnName' => '商品名称商品名称商品名称商品名称商品名称商品名称商品名称', 'width' => 40],
            ['bindKey' => 'price', 'columnName' => '售价', 'align' => 'right',],
            ['bindKey' => 'actualStock', 'columnName' => '实际库存实际库存实际库存实际库存实际库存实际库存实际库存实际库存实际库存实际库存', 'align' => 'right',],
        ];
        

        
        $client = ExportClient::getInstance();
        
//        $redis = new Redis();
//        $redis->connect('redis');
//        $client->setRedis($redis);
        
        $client->setProtection(true);
        return $client->setFilepath(__DIR__ . '/file/' . date('Y/m/d/'))
            ->setFilename('sheetDemo' . date('His'))
            ->setData($data)
            ->setConfig($config)
            ->export();
    }
}

try {
    $filepath = (new SheetExample())->run();
    print_r($filepath);
} catch (\Exception $e) {
    print_r($e->getMessage());
}
echo PHP_EOL;