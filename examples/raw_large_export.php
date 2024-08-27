<?php

# 在const meomory 下, insertDate, insertText 会有内存泄露的问题。
# 执行下面的程序，你可以用top命令，观察到内存持续增长。

$config = ['path' => __DIR__];

$excel = new \Vtiful\Kernel\Excel($config);

$filename = 'million_rows.xlsx';
$excel->constMemory($filename);
// $excel->fileName($filename);

$excel->header(['id', 'first_name', 'last_name', 'born_at']);

echo "my pid-".getmypid();

sleep(3);

$list = function() {
    for ($i = 1; $i <= 1000000; $i++) {
        yield [$i,  'Jane', 'Doe', time()];
    }
};

foreach($list() as $row) {
    $currentRow = $excel->getCurrentLine();

    $excel->insertDate(
        $currentRow,
        3,
        $row[3]
    );
    $row[3] = null;

    $excel->data([$row]);
}

$excel->output();

echo "Excel file with 1 million rows created successfully!\n";
