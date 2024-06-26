<?php

namespace Modules\BetterExcel;

class CellInfo
{
    public $resolvedValue;
    public $columnIndex;
    public $rowIndex;
    /**
     *
     * @var Column
     */
    public $column;

    public static function create($resolvedValue, $columnIndex, $rowIndex,  Column $column)
    {
        $self = new CellInfo();

        $self->resolvedValue = $resolvedValue;
        $self->columnIndex = $columnIndex;
        $self->rowIndex = $rowIndex;
        $self->column = $column;

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
