<?php
/**
 * User: john <john.yi@55haitao.com>
 * Date: 2022/6/17
 */

use Ogenes\Exceler\ExcelClient;

require "vendor/autoload.php";

$client = new ExcelClient();
$config['sheet1'] = [
    'goodsName' => '商品名称',
    'spu' => 'SPU',
    'sku' => 'SKU',
    'color' => '颜色',
    'size' => '尺码',
    'actualStock' => '实际库存',
    'onTheWay' => '在途库存',
];
$filepath = './example.xlsx';
$data = $client->read($filepath, $config);

$exportConfig['sheet'] = [
    ['bindKey' => 'goodsName', 'columnName' => '商品名称', 'align' => 'center', 'width' => 20, 'format' => 'General'],
    ['bindKey' => 'spu', 'columnName' => 'SPU', 'align' => 'center', 'width' => 20, 'format' => 'General'],
    ['bindKey' => 'sku', 'columnName' => 'SKU', 'align' => 'center', 'width' => 20, 'format' => 'General'],
    ['bindKey' => 'color', 'columnName' => '颜色', 'align' => 'center', 'width' => 20, 'format' => 'General'],
    ['bindKey' => 'size', 'columnName' => '尺码', 'align' => 'center', 'width' => 20, 'format' => 'General'],
    ['bindKey' => 'actualStock', 'columnName' => '实际库存', 'align' => 'center', 'width' => 20, 'format' => 'General'],
    ['bindKey' => 'onTheWay', 'columnName' => '在途库存', 'align' => 'center', 'width' => 20, 'format' => 'General'],
];
$newfile = $client->export('newfile', $exportConfig, $data, __DIR__);
print_r($newfile);

