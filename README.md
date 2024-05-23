# BetterExcel

Export to Excel quickly, simply, naturally!

## 使用方法

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

# 已知局限性

-   未支持多个 sheet。

# 测试

```
vendor/bin/phpspec run Modules/Export/spec/BetterExcelSpec.php
```
