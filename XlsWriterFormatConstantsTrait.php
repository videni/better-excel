<?php

namespace Videni\BetterExcel;

use Vtiful\Kernel\Format;

/**
 * Map human readable string to Vtiful\Kernel\Format constants
 */
trait XlsWriterFormatConstantsTrait
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
            // Actually , this const does not exist
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

    private function createFormatter()
    {
        $formats = [
            Style::PATH_FONT_COLOR => function($value, Format $format) {
                // default to back if color not supported
                $format->fontColor(is_string($value)? $this->toColorConstant($value): $value);
            },
            Style::PATH_FONT_SIZE => fn($value, Format $format) => $format->fontSize($value),
            Style::PATH_FONT_NAME => fn($value, Format $format) => $format->font($value),
            Style::PATH_FONT_STYLES => function($styles, $format){
                foreach ($styles as $style) {
                    $format->{strtolower($style)}();
                }
            },
            Style::PATH_ALIGN => function($value, Format $format){
                [$horizontal, $vertical] = $value;

                $format->align(
                    $this->toHorizontalAlignmentConstant($horizontal),
                    $this->toVerticalAlignmentConstant($vertical)
                );
            },
            Style::PATH_UNDERLINE => fn($value, Format $format) => $format->underline($this->toUnderlineConstant($value)),
            Style::PATH_WRAP => fn($value, Format $format) => $value && $format->wrap(),
            Style::PATH_BORDER => fn($value, Format $format) => $format->border($this->toBorderConstant($value)) ,
            Style::PATH_NUMBER => function($value, Format $format){
                $format->number($value);
            },
            Style::PATH_BACKGROUND => function($value, Format $format){
                $format->background($this->toColorConstant($value));
            },
        ];

        return function($path) use($formats) {
            return $formats[$path];
        };
    }
}

