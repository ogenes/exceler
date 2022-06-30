# exceler
封装 phpspreadsheet 工具类

## 安装
```shell
composer require ogenes/exceler
```

## 文档

[wiki](https://github.com/ogenes/exceler/wiki)

## 简单读取
```php
    $config['sheet1'] = [
        'goodsName' => '商品名称',
        'color' => '颜色',
        'price' => '售价',
        'actualStock' => '实际库存',
    ];
    $filepath = __DIR__ . '/file/example.xlsx';
    return ExcelClient::getInstance()->read($filepath, $config);
```

## 简单导出
```php
        $data['sheet1'] = [
            ['goodsName' => '半裙', 'price' => 1490, 'actualStock' => 2,],
            ['goodsName' => '半裙', 'price' => 1590, 'actualStock' => 1,]
        ];
        
        $config['sheet1'] = [
            ['bindKey' => 'goodsName', 'columnName' => '商品名称'],
            ['bindKey' => 'price', 'columnName' => '售价'],
            ['bindKey' => 'actualStock', 'columnName' => '实际库存'],
        ];
        $client = ExportClient::getInstance();
        return $client->setFilepath(__DIR__ . '/file/' . date('Y/m/d/'))
            ->setFilename('file' . date('His'))
            ->setData($data)
            ->setConfig($config)
            ->export();
```

## 输出到浏览器
```php
    ExportClient::getInstance()
        ->setFilename('file' . date('His'))
        ->setData($data)
        ->setConfig($config)
        ->output();
```



## DEMO

以导出收支明细为例

```php
## 可以再次封装定义企业内的Excel固定模板，
class ExcelHelper
{
    public static function export(array $data, array $config, string $filename): string
    {
        $client = ExportClient::getInstance();
        $fill = [
            'fillType' => Fill::FILL_GRADIENT_LINEAR,
            'startColor' => [
                'argb' => 'FFFE00',
            ],
            'endColor' => [
                'argb' => 'FFFE00',
            ]
        ];
        $client->setStyleHeaderFont([
                'name' => '宋体',
                'size' => 11,
                'bold' => true,
                'color' => ['argb' => '000000'],
            ])
            ->setStyleFont([
                'name' => '宋体',
                'size' => 10,
                'color' => ['argb' => '000000'],
            ])
            ->setStyleHeaderFill($fill);
        
        $client->setFreezeHeader(true);
       	return $client->setFilepath(storage_path('excel') . date('/Y/m/d/'))
                ->setFilename($filename)
                ->setData($data)
                ->setConfig($config)
                ->export();
    }
    
}
```



```php
## 导出时定义好config即可。

$data['sheet1'] = $list;

$config['sheet1'] = [
  ['bindKey' => 'orderId', 'columnName' => '订单ID', 'width' => 20, 'align' => Alignment::HORIZONTAL_RIGHT],
  ['bindKey' => 'withdrawDate', 'columnName' => '交易日期', 'width' => 15, 'align' => Alignment::HORIZONTAL_LEFT],
  ['bindKey' => 'statusCn', 'columnName' => '订单状态', 'width' => 10, 'align' => Alignment::HORIZONTAL_LEFT],
  ['bindKey' => 'amount', 'columnName' => '订单金额(USD)', 'width' => 18, 'align' => Alignment::HORIZONTAL_RIGHT, 'format' => NumberFormat::FORMAT_NUMBER_00],
  ['bindKey' => 'cost', 'columnName' => '成本(USD)', 'width' => 18, 'align' => Alignment::HORIZONTAL_RIGHT, 'format' => NumberFormat::FORMAT_NUMBER_00],
  ['bindKey' => '={amount}-{cost}', 'columnName' => '收入(USD)', 'width' => 18, 'align' => Alignment::HORIZONTAL_RIGHT, 'format' => NumberFormat::FORMAT_NUMBER_00],
  ['bindKey' => '=({amount}-{cost})/{amount}', 'columnName' => '毛利率', 'width' => 18, 'align' => Alignment::HORIZONTAL_RIGHT, 'format' => NumberFormat::FORMAT_PERCENTAGE_00],
  ['bindKey' => 'uid', 'columnName' => '用户ID', 'width' => 30, 'align' => Alignment::HORIZONTAL_RIGHT],
  ['bindKey' => 'username', 'columnName' => '用户名', 'width' => 30, 'align' => Alignment::HORIZONTAL_LEFT],
  ['bindKey' => 'note', 'columnName' => '备注', 'width' => 30, 'align' => Alignment::HORIZONTAL_LEFT],

];
$filename = "收支明细" . date('YmdHis') . '.xlsx';
$fileFullpath = ExcelHelper::export($data, $config, $filename);

```

![image-20220630132744221](https://ogenes.oss-cn-beijing.aliyuncs.com/img/2022/202206301327293.png)

## data 明细

data 中有一个预留字段 cellStyle ，用来定义单元格样式， 注意避开。

如果是超链接， 可在data 中指定value为数组， 数组包括 text 和 hyperlink 两个属性。

如果该单元格有备注，也可在data 中指定value为数组， 数组包括 text 和 commnet 两个属性，备注赋给 comment。

```php
$data['sheet1'] = [
  $row1,
  $row2,
  ...
];

$row1 = [
  'field1' => 'value1',
  'field2' => 'value2',
  ...
  'urlField1' => [
    'text' => 'urlValue1',
    'hyperlink' => 'url'
  ],
  'cellStyle' => $style,
];

```



## config 明细

bindKey 和 columnName 为必填字段， 表示data中的键与表头字段的映射。

width 表示宽度。

style 字段可以设置列的样式。

format 指定单元格格式，默认为 General 。

drawing 表示该列为图片。 包括 name、x、y、w、h五个字段。 分别表示图片名、横向偏移、纵向偏移、宽度、高度。

```php
$config['sheet1'] = [
  $conf1,
  $conf2,
  ...
];

$conf1 = [
  'bindKey' => 'a',
  'columnName' => 'A',
  'width' => 30,
  'style' => [
    'font' => [],
    'borders' => [],
    'alignment' => [],
    'fill' => [],
  ],
  'format' => 'General',
  'drawing' => [
    'name' => 'logo',
    'x' => 10,
    'y' => 10,
    'w' => 80,
    'h' => 80,
  ]
]
```




