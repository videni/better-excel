<?php

namespace Modules\BetterExcel\Cell;

use Carbon\Carbon;
use Modules\BetterExcel\Style;
use Modules\BetterExcel\Column;
use Webmozart\Assert\Assert;
use Modules\BetterExcel\XlsWriter;
use PhpOption\None;
use PhpOption\Option;

class Date
{
    private $value;

    /**
     * @var string
     */
    private $format;

    /**
     * @var Style
     */
    private $style;

    public function __construct($value, $format = null, Style $style = null)
    {
        $this->value = $value;
        $this->format = $format ?? 'yyyy/m/d hh:mm:ss';
        $this->style = $style ?? (new Style())->align('center');
    }

    public function render($writer, $rowIndex, $columnIndex, Column $_column): Option
    {
        Assert::isInstanceOf($writer, XlsWriter::class);

        $writer->insertDate(
            $rowIndex,
            $columnIndex,
            $this->value,
            $this->format,
            $this->style ? $writer->formatStyle($this->style): null
        );

        return None::create();
    }

    public static function fromTimeStamp(int $timestamp, $format = null, Style $style = null)
    {
        return new static($timestamp, $format, $style);
    }

    public static function fromCarbon(Carbon $carbon, $format = null, Style $style = null)
    {
        return new static($carbon->getTimestamp(), $format, $style);
    }
}
