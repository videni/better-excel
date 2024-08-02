<?php

namespace Modules\BetterExcel\Cells;

use Carbon\Carbon;
use Carbon\CarbonTimeZone;
use Modules\BetterExcel\CellInfo;
use Modules\BetterExcel\Style;
use Webmozart\Assert\Assert;
use Modules\BetterExcel\XlsWriter;

class Date
{
    protected $value;

    /**
     * @var string
     */
    protected $format;

    /**
     * @var Style
     */
    protected $style;

    public function __construct($value, $format = null, Style $style = null)
    {
        $this->value = $value;
        $this->format = $format ?? 'yyyy-m-d hh:mm:ss';
        $this->style = $style ?? (new Style())->align();
    }

    public function render($writer, CellInfo $info): void
    {
        Assert::isInstanceOf($writer, XlsWriter::class);

        $style = $this->style ?? $info->column->getStyle();

        $writer->insertDate(
            $info->rowIndex,
            $info->columnIndex,
            $this->value,
            $this->format,
            $style ? $writer->formatStyle($style): null
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
