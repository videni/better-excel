<?php

namespace Modules\BetterExcel;

use Generator;
use Illuminate\Support\Collection;
use Illuminate\Support\Str;

class BetterExcel
{
    /**
     * @var Collection|Generator|array
     */
    protected $data;

    private $withHeader = true;

    private $header = [];

    private $options = [];

    /**
     * @param array|Generator|Collection|null $data
     * @param array $options
     */
    public function __construct(array|Generator|Collection $data = null, array $options = [])
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

    public function setHeader($columns = [])
    {
        foreach($columns as $key => $column) {
            $style = $column['style']??[];
            $style = $style instanceof Style ? $style: Style::fromArray((array)$style);
            $this->header[] = new Column($key, $column['label']??$key, $column['path']??$key, $style);
        }
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
     * @param array|Generator|Collection $data
     * @param callable|null $callback
     * @return void
     */
    private function writeRows($writer, array|Generator|Collection $data, ?callable $callback = null)
    {
        foreach($data as $index => $row) {
            $row = $this->transformRow($writer, $row, $index, $callback);

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
        foreach($columns as $column) {
            $value = null;
            if (\is_callable($column->getPath())) {
                $value = $column->getPath()($row, $column, $rowIndex);
            } else {
                $value = data_get($row, $column->getPath());
            }

            if (is_object($value) && method_exists($value, 'render')) {
                $value = $value->render($writer, $column, $rowIndex);
            }

            $newRow[$column->getCode()] = $value;
        }

        return $newRow;
    }
}
