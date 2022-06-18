# exceler
封装 phpspreadsheet 工具类

## 安装
```shell
composer require ogenes/exceler
```

## 简单读取
```php
function run(): array
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
```

## 简单导出
```php

public function run(): string
{
    $data['sheet1'] = [
        ['goodsName' => '半裙', 'price' => 1490, 'actualStock' => 2,],
        ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1,]
    ];
    
    $config['sheet1'] = [
        ['bindKey' => 'goodsName', 'columnName' => '商品名称'],
        ['bindKey' => 'price', 'columnName' => '售价'],
        ['bindKey' => 'actualStock', 'columnName' => '实际库存'],
    ];
    return ExportClient::getInstance()
        ->setFilepath(__DIR__ . '/file/' . date('Y/m/d/'))
        ->setFilename('file' . date('His'))
        ->setData($data)
        ->setConfig($config)
        ->export();
}

```

## 输出到浏览器
```php
    ExportClient::getInstance()
        ->setFilename('file' . date('His'))
        ->setData($data)
        ->setConfig($config)
        ->output();
```