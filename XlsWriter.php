<?php

namespace Modules\BetterExcel;

use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

class XlsWriter
{
    protected static $colors = [
        'black' => Format::COLOR_BLACK,
        'blue' => Format::COLOR_BLUE,
        'brown' => Format::COLOR_BROWN,
        'cyan' => Format::COLOR_CYAN,
        'gray' => Format::COLOR_GRAY,
        'green' => Format::COLOR_GREEN,
        'lime' => Format::COLOR_LIME,
        'magenta' => Format::COLOR_MAGENTA,
        'navy' => Format::COLOR_NAVY,
        'orange' => Format::COLOR_ORANGE,
        'pink' => Format::COLOR_PINK,
        'purple' => Format::COLOR_PURPLE,
        'red' => Format::COLOR_RED,
        'silver' => Format::COLOR_SILVER,
        'white' => Format::COLOR_WHITE,
        'yellow' => Format::COLOR_YELLOW,
    ];

    private $excel;

    public function __construct(string $filename, $options = [])
    {
        $this->excel = new Excel($options);

        $this->excel->fileName($filename);
    }

    public function writeHeader(array $header)
    {
        $predefinedColumns = $this->getPredefinedColumns();
        $maxColumns = count($predefinedColumns);
        $columnsCount = count($header);
        if ($columnsCount > $maxColumns) {
            throw new \Exception(sprintf('The number of columns exceeds the maximum number(%d) of columns', $maxColumns));
        };

        $mappings = array_combine(array_slice($predefinedColumns, 0, $columnsCount), $header);

        foreach ($mappings as $key => $column) {
            $column->setColumnIndex($key);
            $range = sprintf('%1$s1:%1$s1', $key);

            $this->excel->setColumn(
                $range,
                $column->getStyle()->getWidth(),
                $this->format($column->getStyle())
            );
        }

        $labels = array_reduce($header, function($carry, $column){
            $carry[] = $column->getLabel();

            return $carry;
         }, []);
        $this->excel->header($labels);
    }

    public function writeOneRow(array $row)
    {
        $this->excel->data([$row]);
    }

    public function saveToFile($filename = null)
    {
        return $this->excel->output();
    }

    private function getPredefinedColumns()
    {
        static $newSortKey = null ;
        if ($newSortKey) {
            return $newSortKey;
        }
        $sortKey = [
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
        ];
        $newSortKey = $sortKey;
        foreach ($sortKey as $item) {
            foreach ($sortKey as $itemOne) {
                array_push($newSortKey, $item.$itemOne);
            }
        }

        return $newSortKey;
    }

    public function __call($name, $arguments)
    {
        return $this->excel->{$name}(...$arguments);
    }

    public function format(Style $style)
    {
        $handle = $this->excel->getHandle();

        $format = new Format($handle);

        $formatters = [
            Style::PATH_FONT_COLOR => function($value) use($format) {
                if (is_string($value)) {
                    // default to back if color not supported
                    $format->fontColor($this->toColorConstant($value));
                }
            },
            Style::PATH_FONT_SIZE => fn($value) => $format->fontSize($value),
            Style::PATH_FONT_NAME => fn($value) => $format->font($value),
            Style::PATH_FONT_STYLES => function($styles) use($format){
                foreach ($styles as $style) {
                    $format->{strtolower($style)}();
                }
            },
            Style::PATH_ALIGN => function($value) use($format){
                [$horizontal, $vertical] = $value;

                $format->align(
                    $this->toHorizontalAlignmentConstant($horizontal),
                    $this->toVerticalAlignmentConstant($vertical)
                );
            },
            Style::PATH_UNDERLINE => fn($value) => $format->underline($this->toUnderlineConstant($value)),
            Style::PATH_WRAP => fn($value) => $value && $format->wrap(),
            Style::PATH_BORDER => fn($value) => $format->border($this->toBorderConstant($value)) ,
            Style::PATH_NUMBER => function($value) use($format){
                $format->number($value);
            },
            Style::PATH_BACKGROUND => function($value) use($format){
                $format->background($this->toColorConstant($value));
            },
        ];

        return $style->apply(function($formats) use($format, $formatters){
            foreach ($formats as $key => $value) {
                $formatters[$key]($value);
            }

            return $format->toResource();
        });
    }

    private function toColorConstant(string $name)
    {
        return static::$colors[$name]??Format::COLOR_BLACK;
    }

    private function toHorizontalAlignmentConstant($horizontal)
    {
        $alignments = [
            'left' => Format::FORMAT_ALIGN_LEFT,
            'center' => Format::FORMAT_ALIGN_CENTER,
            'right' => Format::FORMAT_ALIGN_RIGHT,
            'fill' => Format::FORMAT_ALIGN_FILL,
            'justify' => Format::FORMAT_ALIGN_JUSTIFY,
            'center_across' => Format::FORMAT_ALIGN_CENTER_ACROSS,
            'distributed' => Format::FORMAT_ALIGN_DISTRIBUTED,
        ];

        return $alignments[$horizontal]??Format::FORMAT_ALIGN_CENTER;
    }

    private function toVerticalAlignmentConstant($vertical)
    {
        $alignments =  [
            'top' => Format::FORMAT_ALIGN_VERTICAL_TOP,
            'bottom' => Format::FORMAT_ALIGN_VERTICAL_BOTTOM,
            'center' => Format::FORMAT_ALIGN_VERTICAL_CENTER,
            'justify' => Format::FORMAT_ALIGN_VERTICAL_JUSTIFY,
            'distributed' => Format::FORMAT_ALIGN_VERTICAL_DISTRIBUTED,
        ];

        return $alignments[$vertical]??Format::FORMAT_ALIGN_VERTICAL_CENTER;
    }

    private function toUnderlineConstant($underline)
    {
        $underlineStyles = [
            'single' => Format::UNDERLINE_SINGLE,
            // 'double' => Format::UNDERLINE_DOUBLE,
            'single_accounting' => Format::UNDERLINE_SINGLE_ACCOUNTING,
            'double_accounting' => Format::UNDERLINE_DOUBLE_ACCOUNTING,
        ];

        return $underlineStyles[$underline]??null;
    }

    private function toBorderConstant($border)
    {
        $borderStyles = [
            'thin' => Format::BORDER_THIN,
            'medium' => Format::BORDER_MEDIUM,
            'dashed' => Format::BORDER_DASHED,
            'dotted' => Format::BORDER_DOTTED,
            'thick' => Format::BORDER_THICK,
            'double' => Format::BORDER_DOUBLE,
            'hair' => Format::BORDER_HAIR,
            'medium_dashed' => Format::BORDER_MEDIUM_DASHED,
            'dash_dot' => Format::BORDER_DASH_DOT,
            'medium_dash_dot' => Format::BORDER_MEDIUM_DASH_DOT,
            'dash_dot_dot' => Format::BORDER_DASH_DOT_DOT,
            'medium_dash_dot_dot' => Format::BORDER_MEDIUM_DASH_DOT_DOT,
            'slant_dash_dot' => Format::BORDER_SLANT_DASH_DOT,
        ];

        return $borderStyles[$border]??null;
    }
}
