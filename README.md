# BetterExcel

Export data to Excel quickly, simply, naturally!

- [BetterExcel](#betterexcel)
  - [使用方法](#使用方法)
    - [一个完整例子](#一个完整例子)
    - [表头](#表头)
    - [样式](#样式)
    - [高级表格单元](#高级表格单元)
      - [图片](#图片)
      - [日期](#日期)
- [已知局限性](#已知局限性)
- [测试](#测试)

## 使用方法

### 一个完整例子

```php
<?php

use Modules\BetterExcel\BetterExcel;
use Modules\BetterExcel\Style;
use Modules\BetterExcel\Cells\Date;
use Carbon\Carbon;

// 1. 定义表格头
$columns = [
    'id' => [
        'label' => 'ID',
        // 以 Fluent Style API设置表头样式。
        'style' =>  (new Style())->bold()->underline()->font('red', 12)->align('center'),
    ],
    'first_name' => [
        'label' => 'Name',
        //  以数组的形式设置表头样式。
        'style' => [
            'font' => [
                'styles' => ['italic', 'bold'],
                'size' => 12,
            ],
            'underline' => 'single',
            'alignment' => [
                'horizontal' => 'center',
                'vertical' => 'center',
            ],
        ]
    ],
    'last_name' => [
        'label' => 'Last Name',
        'style' =>  (new Style())->italic()->underline()->font('cyan', 12)->align('center'),
    ],
    'born_at' => [
        'label' => 'Born At',
        'style' =>  (new Style())->italic()
            ->font('purple', 12)->align('center', 'center')
            ->width(20),
    ],
];

# 2. 假设，用 Generator加载大量数据
$list = function() {
    // The born_at column shows you how to export datetime.
    yield [ 'id' => 1, 'first_name' => 'Jane', 'last_name' => 'Doe', 'born_at' => Date::fromTimeStamp(time())];
    yield [ 'id' => 2, 'first_name' => 'John', 'last_name' => 'Doe', 'born_at' => Date::fromCarbon(Carbon::now())];
};

$betterExcel = new BetterExcel($list(), [
    'path' => __DIR__
]);
$betterExcel->setHeader($columns);

# 3. 导出为 excel
echo $betterExcel->export('test.xls');
```

It outputs something like below

![alt text](./docs/image.png)

### 表头

结构如下:

```php
$columns = [
    'id' => [
        'label' => 'ID', # excel 的表头显示名。

        // 以 Fluent Style API设置表头样式。
        'style' =>  (new Style())->bold()->underline()->font('red', 12)->align('center'),
    ],
    'name' => [
        'label' => 'Name',
        //  以数组的形式设置表头样式。
        'style' => [
            'font' => [
                'styles' => ['italic', 'bold'],
                'size' => 12,
            ],
            'underline' => 'single',
            'alignment' => [
                'horizontal' => 'center',
                'vertical' => 'center',
            ],
        ]
    ],
    'born_at' => [
        'label' => 'Born At',
        'style' =>  (new Style())->italic()
            ->font('purple', 12)->align('center', 'center')
            ->width(20),
    ],
];
```

上述表头的列`id`, 表示获取某一个行数据，键为`id` 的值, 如。
```
[ 'id' => 1, 'name' => 'Jane', 'born_at' => Date::fromTimeStamp(time())]
```
如果列与行数据的键不一致，也可以在表头定义中，设置`path`字段， 如
```php
$columns = [
    'id' => [
        'label' => 'ID',
        'path' => 'user_id',
        // 以 Fluent Style API设置表头样式。
        'style' =>  (new Style())->bold()->underline()->font('red', 12)->align('center'),
    ],
]
```

`path` 也可以为 closure，它会在该列的行单元格被渲染的时候执行，如：
```
$columns = [
    'id' => [
        'label' => 'ID',
        'path' => function($row, $rowIndex, $columnIndex, $column) {
            return $row['user_id'];
        },
    ],
]
```

返回值会被写入Excel中， 但如果返回的是`PhpOption\None`, 意味着，你将自己负责渲染这个cell， closure的参数已经告诉你单元格具体的行，列号，以及分配的列字母($column->getLetter()), 你完全可以自己渲染。

### 样式

* 简单数组

```
Style::fromArray(
    [
        'font' => [
            'styles' => ['italic', 'bold'],
            'size' => 12,
        ],
        'underline' => 'single',
        'alignment' => [
            'horizontal' => 'center',
            'vertical' => 'center',
        ],
    ]
);
```

* Fluent Style API
  
```
(new Style())->bold()->underline()->font('red', 12)->align('center'),
```

所有的可用方法，请参考`Modules\BetterExcel\Style`, 这个类只是[PHP-Xlswriter
样式列表](https://xlswriter-docs.viest.me/zh-cn/yang-shi-lie-biao)的简单封装。

另外，关于`颜色`，`下划线`、`对齐`，你可以在`Modules/BetterExcel/XlsWriterFormatConstantsTrait.php`快速找到所有预定义的选项。

### 高级表格单元

除了前面的`path`字段处理单元格的方式，你还可以在数据中直接返回一个带`render`方法对象, 这个方法与 `path`的 closure用法一模一样，唯一的区别是， 第一个参数是 XlsWriter对象，而不是整行数据，你的直觉会告诉你为什么是XlsWriter对象，而不是整行数据。

这里的高级单元格，都是通过这中方法实现的。

#### 图片

在数据中，返回`Modules\BetterExcel\Cells\EmbedImage` 对象即可。

```php

[ 'id' => 1, 'name' => 'Jane', 'avatar' => EmbedImage::fromUrl('https://gravatar.com/avatar/fc8c88023f9efa517dbd7cece0c54167?s=400&d=robohash&r=x')];

```

#### 日期

在数据中，返回`Modules\BetterExcel\Cells\Date` 对象即可。

```php
[ 'id' => 1, 'name' => 'Jane', 'born_at' => Date::fromTimeStamp(time())];
[ 'id' => 2, 'name' => 'John', 'born_at' => Date::fromCarbon(Carbon::now())];
```

# 已知局限性

-   未支持多个 sheet。

# 测试

```
vendor/bin/phpspec run Modules/Export/spec/BetterExcelSpec.php
```
