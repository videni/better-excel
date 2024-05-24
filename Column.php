<?php

namespace Modules\BetterExcel;

class Column
{
    private $code;
    private $label;
    private $style;
    private $path;
    private $letter;

    public function __construct(string $code, string $label, string|\Closure $path = null, Style $style = null)
    {
        $this->code = $code;
        $this->label = $label;
        $this->style = $style;
        $this->path = $path;
    }

    public function getLabel(): string
    {
        return $this->label;
    }

    public function getStyle(): ?Style
    {
        return $this->style;
    }

    public function getPath()
    {
        return $this->path;
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
}
