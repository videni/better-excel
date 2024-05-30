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

    private function createHeader($columns = [])
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
    private function writeRows($writer, array|Generator|Enumerable $data, ?callable $callback = null)
    {
        $delta = count($this->header) > 0 ? 1: 0;

        foreach($data as $index => $row) {
            $rowIndex = $this->withHeader? $index + $delta: $index;

            $row = $this->transformRow($writer, $row, $rowIndex, $callback);

            $writer->writeOneRow($row);
        }
    }

    private function writeHeader($writer, $headers = [])
    {
        $writer->writeHeader($headers);
    }

    private function transformRow($writer, $row, $rowIndex, $callback = null)
    {
        $columns = $this->getHeader();
        if ($callback) {
            $row = $callback($row);
        }

        $newRow = [];
        foreach($columns as $columnIndex => $column) {
            $path = $column->getPath();
            $value = data_get($row, $path);
            if ($column->hasResolver())  {
                $resolver = $column->getResolver();
                $value = $resolver($value, $row);
            }

            if (is_object($value) && method_exists($value, 'render')) {
                $value = $value->render($writer, $rowIndex, $columnIndex, $column);
                // If its "None" option, it means you are responsible to render this cell yourself
                if ($value instanceof Option) {
                   $value->forAll(function($v) use (&$newRow, $column){
                       $newRow[$column->getCode()] = $v;
                   });

                   continue;
                }
            }

            $newRow[$column->getCode()] = $value;
        }

        return $newRow;
    }
}
