<?php

namespace Modules\BetterExcel;

class Column
{
    private $code;
    private $label;
    private $style;
    private $path;
    private $columnIndex;

    public function __construct(string $code, string $label, $path = null, Style $style = null)
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

    public function setColumnIndex($index)
    {
        $this->columnIndex = $index;

        return $this;
    }

    public function getColumnIndex()
    {
        return $this->columnIndex;
    }
}
