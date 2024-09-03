<?php

// A demo to show memory usage for using raw excel API.

use Vtiful\Kernel\Format;

$config = ['path' => __DIR__];

$excel = new \Vtiful\Kernel\Excel($config);

$filename = 'million_rows.xlsx';
$excel->constMemory($filename);
// $excel->fileName($filename);

$excel->header(['id', 'first_name', 'last_name', 'born_at']);

echo "my pid-". getmypid().PHP_EOL;

// echo "I will sleep for 5 seconds...";

// sleep(10);

$list = function() {
    for ($i = 1; $i <= 1000000; $i++) {
        yield array_merge([$i,  'Jane', 'Doe', time()], array_reduce(range(5, 180), function($carry, $key) {
            $carry[] = time();

            return $carry;
        }));
    }
};

$handle = $excel->getHandle();

$format = new Format($handle);
$format->italic()
    ->underline(Format::UNDERLINE_SINGLE)
    ->align(Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER);

echo "Current memory usage: " . memory_get_usage(true)/1000/1000 . "MB\n";

foreach($list() as $row) {
    $currentRow = $excel->getCurrentLine();

    $dates = array_splice($row, 4, 180, array_fill(0, 176, null) );

    foreach($dates as $date) {
        $excel->insertDate(
            $currentRow,
            3,
            $date,
            null,
            $format->toResource(),
        );
    }

    $excel->data([$row]);
}

echo "Current memory usage: " . memory_get_usage(true)/1000/1000 . "MB\n";

$excel->output();

echo "I will sleep for 10 seconds...";
sleep(15);

echo "Excel file with 1 million rows created successfully!\n";
