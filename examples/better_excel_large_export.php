<?php

// A demo to show memory usage for BetterExcel

// 导出180字段，与raw_large_export文件基本等同。

require_once __DIR__.'/../../../vendor/autoload.php';

use Modules\BetterExcel\BetterExcel;
use Modules\BetterExcel\Style;
use Modules\BetterExcel\Cells\Date;

$list = function() {
    foreach(range(1, 1000000) as $i) {
        yield [ 'id' => $i, 'first_name' => 'Jane', 'last_name' => 'Doe', 'born_at' => Date::fromTimeStamp(time())]
            + array_reduce(range(5, 180), function($carry, $key) {
                $carry['born_at'.$key-4] = Date::fromTimeStamp(time());

                return $carry;
            })
        ;
    }
};

$betterExcel = new BetterExcel($list(), ['path' => __DIR__, 'const_memory' => true]);

$headers = [
    'id' => [
        'label' => 'ID',
        'style' =>  (new Style())->italic()->underline()->font('red', 12)->align('center'),
    ],
    'first_name' => [
        'label' => 'Name',
        'style' => [
            'font' => [
               'style' => ['italic', 'bold'],
               'size' => 12,
               'name' => 'Arial',
            ],
            'underline' => 'single',
            'alignment' => [
                // 'horizontal' => 'center',
                'vertical' => 'center',
            ],
        ]
    ],
    'last_name' => [
        'label' => 'Last Name',
        'style' =>  (new Style())->italic()->underline('double')->font('cyan', 12)->align('center')->border(null),
    ],
    'born_at' => [
        'label' => 'Born At',
        'style' =>  (new Style())->italic()->underline('double')
            ->font('purple', 12)->align('center', 'center')->border(null)
            ->width(20),
    ],
];

$bornAt = $headers['born_at'];

$headers = $headers
    + array_reduce(range(5, 180), function($carry, $index) use($bornAt) {
        $carry ['born_at'.$index - 4 ] = $bornAt;
        return $carry;
    },)
;

// dump($headers);exit;
$betterExcel->setHeader($headers);

echo "my-pid-".getmypid().PHP_EOL;
// sleep(4);

$betterExcel->export('test.xls');

// sleep(15);
