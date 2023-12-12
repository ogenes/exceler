<?php

use Ogenes\Exceler\ExportClient;

require "vendor/autoload.php";

/**
 * Created by exceler.
 * User: Ogenes
 * Date: 2023/12/12
 */
class FormulaExample
{
    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function run(): string
    {
        $data['sheet1'] = [
          ['goodsName' => '半裙', 'price' => 1490, 'actualStock' => 2],
          ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1]
        ];
        
        $config['sheet1'] = [
          ['bindKey' => 'goodsName', 'columnName' => '商品名称'],
          ['bindKey' => '={price}*{actualStock}', 'columnName' => '库存额'], //会自动转化为公式，所在行的 price * actualStock
          ['bindKey' => 'price', 'columnName' => '售价'],
          ['bindKey' => 'actualStock', 'columnName' => '实际库存'],
        ];
        $client = ExportClient::getInstance();
        $client->setStyleWidth(50);
        $client->setStyleUnit('');
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
