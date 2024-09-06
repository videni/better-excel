<?php

namespace Modules\BetterExcel\Cells;

use Carbon\Carbon;
use Carbon\CarbonTimeZone;
use Modules\BetterExcel\CellInfo;
use Modules\BetterExcel\Style;
use Webmozart\Assert\Assert;
use Modules\BetterExcel\XlsWriter;

class Date extends BaseCell
{
    protected $value;

    /**
     * @var string
     */
    protected $format;


    public function __construct($value, $format = null, Style $style = null)
    {
        $this->value = $value;
        $this->format = $format;
        $this->style = $style;
    }

    public function render(XlsWriter $writer, CellInfo $info): void
    {
        Assert::isInstanceOf($writer, XlsWriter::class);

        $writer->insertDate(
            $info->rowIndex,
            $info->columnIndex,
            $this->value,
            $this->format,
            $this->getFormattedStyle($writer, $info)
        );
    }

    public static function fromTimeStamp(int $timestamp, $format = null, Style $style = null)
    {
        return new static($timestamp, $format, $style);
    }

    public static function fromCarbon(Carbon $carbon, $timezone = 'UTC', $format = null, Style $style = null)
    {
        $tz = new CarbonTimeZone($timezone);

        $localTime = $carbon->copy()->timezone($tz);

        $offset = $localTime->offset;

        $localTimestamp = $carbon->getTimestamp() + $offset;

        return new static($localTimestamp, $format, $style);
    }
}
