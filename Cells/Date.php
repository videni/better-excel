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
        $this->format = $format;
        $this->style = $style;
    }

    public function render($writer, CellInfo $info): void
    {
        Assert::isInstanceOf($writer, XlsWriter::class);

        $column = $info->column;

        // 如果当前 cell 有设置 style ，则优先使用它。
        $formattedStyle = $this->style?->getFormattedStyle();
        if ($this->style && !$formattedStyle) {
            $formattedStyle = $writer->formatStyle($this->style);
        }
        // 否则， 使用 header 设置的样式。
        else {
            $formattedStyle = $column->getStyle()?->getFormattedStyle();
        }

        $writer->insertDate(
            $info->rowIndex,
            $info->columnIndex,
            $this->value,
            $this->format,
            $formattedStyle
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
