# exceler
封装 phpspreadsheet 工具类

## 安装
```shell
composer require ogenes/exceler
```

## 读取excel文件
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

## 导出excel文件
```php

function run(): string
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

```