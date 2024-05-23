<?php

namespace Modules\BetterExcel;

use PhpOption\None;
use PhpOption\Option;
use PhpOption\Some;

/**
 * This class allows you to set styles for Excel cells quickly and easily.
 *
 * It is a simple wrapper around the \Vtiful\Kernel\Format class
 */
class Style
{
    public const FONT_STYLES = ['italic', 'bold'];

    /**
     * https://xlswriter-docs.viest.me/zh-cn/yang-shi-lie-biao/xia-hua-xian
     */
    public const UNDERLINE_STYLES = ['single', 'double', 'single_accounting', 'double_accounting'];

    /**
     * https://xlswriter-docs.viest.me/zh-cn/yang-shi-lie-biao/bian-kuang-yang-shi-chang-liang
     */
    public const BORDER_STYLES = [
        'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair',
        'medium_dashed', 'dash_dot', 'medium_dash_dot', 'dash_dot_dot',
        'medium_dash_dot_dot', 'slant_dash_dot'
    ];

    /**
     * https://xlswriter-docs.viest.me/zh-cn/yang-shi-lie-biao/yan-se-chang-liang
     */
    public const COLOR_STYLE = [
        'black', 'blue', 'brown', 'cyan', 'gray', 'green', 'lime', 'magenta', 'navy', 'orange', 'pink', 'purple', 'red', 'silver', 'white', 'yellow'
    ];

    public const PATH_FONT_STYLES = 'font.styles';
    public const PATH_FONT_SIZE = 'font.size';
    public const PATH_FONT_NAME = 'font.name';
    public const PATH_FONT_COLOR = 'font.color';
    public const PATH_FONT_UNDERLINE = 'font.underline';
    public const PATH_STRIKEOUT = 'font.strikeout';
    public const PATH_WRAP = 'wrap';
    public const PATH_ALIGN = 'align';
    public const PATH_BORDER = 'border';
    public const PATH_BACKGROUND = 'background';
    public const PATH_NUMBER = 'number';
    public const PATH_WIDTH = 'width';
    public const PATH_HEIGHT = 'height';
    public const PATH_UNDERLINE = 'underline';

    private $formats = [];

    public function __construct()
    {
        $this->formats = [];
    }

    public function italic()
    {
        $this->setFormat(self::PATH_FONT_STYLES, 'italic', true);

        return $this;
    }

    public function font($color, $size, $name = null)
    {
        $this->fontColor($color);
        $this->fontSize($size);
        $this->fontName($name);

        return $this;
    }

    public function fontStyles($styles)
    {
        $styles = is_array($styles) ? $styles : [$styles];
        foreach ($styles as $style) {
            $this->throwIfInvalidOption($style, self::FONT_STYLES, 'font style');

            $this->{$style}();
        }

        return $this;
    }
    public function fontColor($color)
    {
        if (is_string($color) ) {
            $color = strtolower($color);
            if (!in_array($color, self::COLOR_STYLE)) {
                throw new \InvalidArgumentException(
                    sprintf('Invalid color: %s, predefined colors are: %s', $color, join(', ',  self::COLOR_STYLE))
                );
            }
        }
        $this->setFormat(self::PATH_FONT_COLOR, $color);

        return $this;
    }

    public function fontSize(int $size)
    {
        $this->setFormat(self::PATH_FONT_SIZE, $size);

        return $this;
    }

    public function fontName(string $name = null)
    {
        $this->setFormat(self::PATH_FONT_NAME, $name);

        return $this;
    }

    public function bold()
    {
        $this->setFormat(self::PATH_FONT_STYLES, 'bold', true);

        return $this;
    }

    public function align($horizontal, $vertical = null)
    {
        $horizontalAlignments = [
            'left', 'center', 'right', 'fill', 'justify', 'center_across', 'distributed'
        ];
        $this->assetOption($horizontal, $horizontalAlignments, 'horizontal alignment');

        $verticalAlignments = [
            'top', 'bottom', 'center', 'justify', 'distributed'
        ];
        $this->assetOption($vertical, $verticalAlignments, 'vertical alignment');

        $this->setFormat(self::PATH_ALIGN, [$horizontal, $vertical]);

        return $this;
    }

    public function strikeout()
    {
        $this->setFormat(self::PATH_STRIKEOUT, 'strikeout', true);

        return $this;
    }

    /**
     * @param string $underline default "single"
     * @return $this
     */
    public function underline($underline = 'single')
    {
        $this->assetOption($underline, self::UNDERLINE_STYLES, 'underline style');

        $this->setFormat('underline', $underline);

        return $this;
    }

    public function wrap()
    {
        $this->setFormat('wrap', true);

        return $this;
    }

    public function border($border)
    {
        $this->assetOption($border, self::BORDER_STYLES, 'border style');
        $this->setFormat(self::PATH_BORDER, $border);

        return $this;
    }

    public function background($background)
    {
        $this->assetOption($background, self::COLOR_STYLE, 'background color');
        $this->setFormat(self::PATH_BACKGROUND, $background);

        return $this;
    }

    /**
     *
     * https://xlswriter-docs.viest.me/zh-cn/yang-shi-lie-biao/number
     *
     * example:
     * "0.000"
     * "#,##0"
     * "#,##0.00"
     * "0.00"
     *
     * @param string $format
     * @return $this
     */
    public function number(string $format)
    {
        $this->setFormat(self::PATH_NUMBER, $format);

        return $this;
    }

    private function assetOption($option, $choices, $hint)
    {
        //nullable option is allowed
        if ($option === null) {
            return;
        }
        $this->throwIfInvalidOption($option, $choices, $hint);
    }

    private function throwIfInvalidOption($option, array $choices, string $hint)
    {
        $option = strtolower($option);
        if (!in_array($option, $choices)) {
            throw new \InvalidArgumentException(
                sprintf('Invalid %s: %s, valid choices are: %s', $hint, $option, join(', ',  $choices))
            );
        }
    }

    public function apply(\Closure $callback)
    {
        return $callback($this->formats);
    }

    /**
     * Set column width, it is only needed for column,
     * usually, for column width, it means how many characters can be displayed
     *
     * @param int $width
     * @return self
     */
    public function width(int $width)
    {
        $this->setFormat(self::PATH_WIDTH, $width);

        return $this;
    }

    /**
     * Set row height, it is only needed for row
     *
     * @param int $height
     * @return self
     */
    public function height(int $height)
    {
        $this->setFormat(self::PATH_HEIGHT, $height);

        return $this;
    }

    public function getWidth()
    {
        return $this->formats[self::PATH_WIDTH]??null;
    }

    private function setFormat(string $field, $value, $multiple = false)
    {
        if ($multiple && !in_array($value, $this->formats[$field]??[])) {
            $this->formats[$field][] = $value;
        } else {
            $this->formats[$field] = $value;
        }
    }

    public static function fromArray(array $formats)
    {
        $self = new self();

        // null is valid value, I use NONE to decide if it is set manually
        $self->getFormatFromArray($formats, self::PATH_FONT_STYLES, None::create())
            ->forAll(function($value) use ($self){
                $self->fontStyles($value);
            });
        $self->getFormatFromArray($formats, self::PATH_FONT_SIZE, None::create())
            ->forAll(function($value) use ($self){
                $self->fontSize($value);
            });
        $self->getFormatFromArray($formats, self::PATH_FONT_NAME, None::create())
            ->forAll(function($value) use ($self){
                $self->fontName($value);
            });
        $self->getFormatFromArray($formats, self::PATH_FONT_COLOR, None::create())
            ->forAll(function($value) use ($self){
                $self->fontColor($value);
            });
        $self->getFormatFromArray($formats, self::PATH_FONT_UNDERLINE, None::create())
            ->forAll(function($value) use ($self){
                $self->underline($value);
            });
        $self->getFormatFromArray($formats, self::PATH_STRIKEOUT, None::create())
            ->forAll(function($value) use ($self){
                $value && $self->strikeout();
            });
        $self->getFormatFromArray($formats, self::PATH_WRAP, None::create())
            ->forAll(function($value) use ($self){
                $value && $self->wrap();
            });
        $self->getFormatFromArray($formats, self::PATH_ALIGN, None::create())
            ->forAll(function($value) use ($self){
                $ $self->align($value['horizontal']??null, $value['vertical']??null);
            });
        $self->getFormatFromArray($formats, self::PATH_BORDER, None::create())
            ->forAll(function($value) use ($self){
                $self->border($value);
            });
        $self->getFormatFromArray($formats, self::PATH_NUMBER, None::create())
            ->forAll(function($value) use ($self){
                $self->number($value);
            });
        $self->getFormatFromArray($formats, self::PATH_WIDTH, None::create())
            ->forAll(function($value) use ($self){
                $self->width($value);
            });
        $self->getFormatFromArray($formats, self::PATH_HEIGHT, None::create())
            ->forAll(function($value) use ($self){
                $self->height($value);
            });

        return $self;
    }

    private function getFormatFromArray(array $formats, string $path, $default = null): Option
    {
       $data = data_get($formats, $path, $default);

       return $data instanceof None ? $data : Some::create($data);
    }
}
