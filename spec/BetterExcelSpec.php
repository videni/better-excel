<?php

namespace spec\Modules\BetterExcel;

use Modules\BetterExcel\BetterExcel;
use Modules\BetterExcel\Style;
use PhpSpec\ObjectBehavior;

class BetterExcelSpec extends ObjectBehavior
{
    function let()
    {
        $list = function() {
            yield [ 'id' => 1, 'first_name' => 'Jane', 'last_name' => 'Doe' ];
            yield [ 'id' => 2, 'first_name' => 'John', 'last_name' => 'Doe'];
        };

        $this->beConstructedWith($list());

        $this->setOptions(['path' => __DIR__ ]);
    }

    function it_is_initializable()
    {
        $this->shouldHaveType(BetterExcel::class);
    }

    function it_exports_to_excel()
    {
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
                'style' =>  (new Style())->italic()->underline('double')->font('cyan', 12)->align('center'),
            ],
        ];

        $this->setHeader($headers);

        $this->export('test.xls')->shouldBeEqualTo(__DIR__ . '/test.xls');
    }
}
