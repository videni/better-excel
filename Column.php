<?php

namespace Videni\BetterExcel;

class Column
{
    private $code;
    private $label;
    private $style;
    private $path;
    private $letter;
    private $resolver;

    public function __construct(string $code, string $label, string $path = null, callable $resolver = null, Style $style = null)
    {
        $this->code = $code;
        $this->label = $label;
        $this->style = $style;
        $this->path = $path;
        $this->resolver = $resolver;
    }

    public function getLabel(): string
    {
        return $this->label;
    }

    public function getStyle(): ?Style
    {
        return $this->style;
    }

    public function getPath(): string
    {
        return $this->path ?? $this->code;
    }

    public function setPath(string $path = null): self
    {
        $this->path = $path;

        return $this;
    }

    public function getCode()
    {
        return $this->code;
    }

    public function setLetter($letter)
    {
        $this->letter = $letter;

        return $this;
    }

    /**
     *  Get the column index, such as A, B, C, etc.
     *
     *  For some methods(setRow) in \Vtiful\Kernel\Excel, the column letter is required instead
     *  of integer based column index. this method here is to make your life easier.
     *
     * @return mixed
     */
    public function getLetter()
    {
        return $this->letter;
    }

    public function setResolver($resolver): self
    {
        $this->resolver = $resolver;

        return $this;
    }

    public function hasResolver()
    {
        return $this->resolver !== null;
    }

    public function getResolver(): ?callable
    {
        return $this->resolver;
    }

    public static function fromArray(string $code, array $column)
    {
        $style = $column['style']??[];
        $style = $style instanceof Style ? $style: Style::fromArray((array)$style);

        return new static($code,
            $column['label']?? $code,
            $column['path']?? $code,
            $column['resolver']?? null,
            $style
        );
    }

    public function getUnresolvedValue($row, $default = null)
    {
        return data_get($row, $this->getPath(), $default);
    }
}
