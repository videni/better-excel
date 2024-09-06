<?php

namespace Modules\BetterExcel;

class CellInfo
{
    public $resolvedValue;
    public $columnIndex;
    public $rowIndex;
    /**
     *
     * @var Style|null
     */
    public $style;

    /**
     * @var string|null
     */
    public $columnLetter;

    public static function create($resolvedValue, $columnIndex, $rowIndex, Style $style = null, $columnLetter = null)
    {
        $self = new CellInfo();

        $self->resolvedValue = $resolvedValue;
        $self->columnIndex = $columnIndex;
        $self->rowIndex = $rowIndex;
        $self->style = $style;
        $self->columnLetter = $columnLetter;

        return $self;
    }

    public function isRenderingObject(): bool
    {
        return  is_object($this->resolvedValue) && method_exists($this->resolvedValue, 'render');
    }

    public function render($writer)
    {
        if ($this->isRenderingObject()) {
            $this->resolvedValue->render($writer, $this);

            return null;
        }

        return $this->resolvedValue;
    }
}
