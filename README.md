BetterExcel
==========

Export to Excel quickly, easily, naturally！

## 使用方法

```php

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
        'style' =>  (new Style())->italic()->underline('double')->font('cyan', 12)->align('center')->border(null),
    ],
];

# 2. 假设，用 Generator加载大量数据
$list = function() {
    yield [ 'id' => 1, 'first_name' => 'Jane', 'last_name' => 'Doe' ];
    yield [ 'id' => 2, 'first_name' => 'John', 'last_name' => 'Doe'];
};

$betterExcel = new BetterExcel($list());
$betterExcel->setHeader($columns);

# 3. 导出为 excel
echo $betterExcel->export('test.xls');
```

# 已知局限性

* 未支持多个sheet。


# 测试

```
vendor/bin/phpspec run Modules/Export/spec/BetterExcelSpec.php
```
