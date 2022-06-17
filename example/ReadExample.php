<?php

use Ogenes\Exceler\ExcelClient;

require "vendor/autoload.php";


/**
 * User: ogenes<ogenes.yi@gmail.com>
 * Date: 2022/6/17
 */
class ReadExample
{
    public function run(): array
    {
        $config['sheet1'] = [
            'goodsName' => '商品名称',
            'color' => '颜色',
            'price' => '售价',
            'actualStock' => '实际库存',
        ];
        $filepath = __DIR__ . '/file/example.xlsx';
        return ExcelClient::getInstance()->read($filepath, $config);
    }
}

$data = (new ReadExample())->run();
print_r($data);