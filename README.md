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

## DEMO

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




