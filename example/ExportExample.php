<?php

use Ogenes\Exceler\ExcelClient;
use PhpOffice\PhpSpreadsheet\Chart\Properties;

require "vendor/autoload.php";

/**
 * User: ogenes<ogenes.yi@gmail.com>
 * Date: 2022/6/17
 */
class ExportExample
{
    public function run(): string
    {
        $data['sheet1'] = [
            [
                'goodsName' => '半裙',
                'price' => 1490,
                'actualStock' => 2,
            ],
            [
                'goodsName' => '半裙',
                'price' => 1590,
                'actualStock' => 1,
            ]
        ];
        
        $config['sheet'] = [
            ['bindKey' => 'goodsName', 'columnName' => '商品名称', 'width' => 30],
            ['bindKey' => 'price', 'columnName' => '售价', 'align' => 'right', 'format' => Properties::FORMAT_CODE_ACCOUNTING],
            ['bindKey' => 'actualStock', 'columnName' => '实际库存', 'align' => 'right'],
        ];
        return ExcelClient::getInstance()->export('newfile', $config, $data, __DIR__ . '/file/');
    }
}

$filepath = (new ExportExample())->run();
print_r($filepath);