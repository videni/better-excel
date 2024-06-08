<?php

namespace Modules\BetterExcel;

use Generator;
use Illuminate\Support\Enumerable;
use Illuminate\Support\Str;
use PhpOption\Option;

class BetterExcel
{
    /**
     * @var Enumerable|Generator|array
     */
    protected $data;

    private $withHeader = true;

    private $header = [];

    private $options = [];

    private $rowIndex = 0;

    /**
     * @param array|Generator|Enumerable|null $data
     * @param array $options
     */
    public function __construct(array|Generator|Enumerable $data = null, array $options = [])
    {
        $this->data = $data;
        $this->options = $options;
    }

    public function export(string $filename, callable $callback = null)
    {
        $writer = $this->createWriter($filename, $this->options);
        if ($this->withHeader) {
            $this->writeHeader($writer, $this->getHeader());
        }

        $this->writeRows($writer, $this->data, $callback);

        return $writer->saveToFile($filename);
    }

    protected function createWriter($filename, $options = [])
    {
        if (Str::endsWith($filename, 'xls')) {
            return new XlsWriter($filename, $options);
        }

        throw new \Exception('Currently only support Excel export');
    }

    public function setHeader(array $columns)
    {
        $this->header =$this->createHeader($columns);
    }

    protected function createHeader($columns = [])
    {
        $results = [];
        foreach($columns as $key => $column) {
            $results[] = $column instanceof Column ? $column: Column::fromArray($key, $column);
        }

        return $results;
    }

    public function withoutHeader()
    {
        $this->withHeader = false;

        return $this;
    }

    public function getHeader(): array
    {
        return $this->header;
    }

    public function setOptions($options = [])
    {
        $this->options = $options;

        return $this;
    }

    /**
     * @param  $writer
     * @param array|Generator|Enumerable $data
     * @param callable|null $callback
     * @return void
     */
    protected function writeRows($writer, array|Generator|Enumerable $data, ?callable $callback = null)
    {
        $delta = count($this->header) > 0 ? 1: 0;

        $rowIndex = $this->rowIndex;

        foreach($data as $row) {
            $rowIndex = $this->withHeader ? $rowIndex+ $delta: $rowIndex;
            if ($callback) {
                $row = $callback($writer, $row, $rowIndex, $this->header);
            }
            [$simple, $complex] = $this->transformRow($row);

            $writer->writeOneRow(array_values($simple));

            foreach($complex as $cell) {
                [$value, $columnIndex, $column] = $cell;
                $value->render($writer, $rowIndex, $columnIndex, $column);
            }

            $this->rowIndex++;
        }
    }

    protected function writeHeader($writer, $headers = [])
    {
        $writer->writeHeader($headers);
    }

    protected function transformRow($row)
    {
        $columns = $this->getHeader();

        $simple = [];
        $complex = [];

        foreach($columns as $columnIndex => $column) {
            $value = $column->getUnresolvedValue($row);
            if ($resolver = $column->getResolver())  {
                $value = $resolver($value, $row);
            }
            if (is_object($value) && method_exists($value, 'render')) {
                $complex[] = [$value, $columnIndex, $column];
               //the cell will render itself, so  we set it to null to keep the data position
               $simple[$column->getCode()] = null;
               continue;
            }
            $simple[$column->getCode()] = $value;
        }

        return [$simple, $complex];
    }
}
