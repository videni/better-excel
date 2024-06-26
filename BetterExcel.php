<?php

namespace Modules\BetterExcel;

use Generator;
use Illuminate\Support\Enumerable;
use Illuminate\Support\Str;

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
        $rowIndex = $this->rowIndex;
        if ($this->withHeader && count($this->header) > 0) {
            $rowIndex =  $rowIndex + 1;
        }

        foreach($data as $row) {
            if ($callback) {
                $row = $callback($writer, $row, $rowIndex, $this->header);
            }
            // Allow to render the row by itself
            if (is_object($row) && method_exists($row, 'render')) {
                $row = $row->render($writer, $rowIndex, $this->header);
            }

            $data = $this->transformRow($row, $rowIndex);

            $writer->writeOneRow($data);

            $rowIndex++;
        }

        $this->rowIndex = $rowIndex;
    }

    protected function writeHeader($writer, $headers = [])
    {
        $writer->writeHeader($headers);
    }

    protected function transformRow($row, $rowIndex)
    {
        $columns = $this->getHeader();

        $data = [];

        foreach($columns as $columnIndex => $column) {
            $value = $column->getUnresolvedValue($row);
            if ($resolver = $column->getResolver())  {
                $value = $resolver($value, $row);
            }

            $data[] = CellInfo::create($value, $columnIndex, $rowIndex, $column);
        }

        return $data;
    }
}
