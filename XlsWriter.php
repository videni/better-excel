<?php

namespace Modules\BetterExcel;

use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

class XlsWriter
{
    use XlsWriterFormatConstantsTrait;

    private $excel;

    public function __construct(string $filename, $options = [])
    {
        $this->excel = new Excel($options);

        $this->excel->fileName($filename);
    }

    public function writeHeader(array $header)
    {
        $columnsLetters = $this->getPredefinedColumnLetters();

        $maxColumns = count($columnsLetters);
        $columnsCount = count($header);
        if ($columnsCount > $maxColumns) {
            throw new \Exception(sprintf('The number of columns exceeds the maximum number(%d) of columns', $maxColumns));
        };

        $mappings = array_combine(array_slice($columnsLetters, 0, $columnsCount), $header);

        foreach ($mappings as $letter => $column) {
            $column->setLetter($letter);
            $range = sprintf('%1$s1:%1$s1', $letter);

            $this->excel->setColumn(
                $range,
                $column->getStyle()->getWidth(),
                $this->formatStyle($column->getStyle())
            );
        }

        $labels = array_reduce($header, function($carry, $column){
            $carry[] = $column->getLabel();

            return $carry;
         }, []);
        $this->excel->header($labels);
    }

    public function writeOneRow(array $row)
    {
        $this->excel->data([$row]);
    }

    public function saveToFile($filename = null)
    {
        return $this->excel->output();
    }

    private function getPredefinedColumnLetters()
    {
        static $newSortKey = null ;
        if ($newSortKey) {
            return $newSortKey;
        }
        $sortKey = [
            'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
        ];
        $newSortKey = $sortKey;
        foreach ($sortKey as $item) {
            foreach ($sortKey as $itemOne) {
                array_push($newSortKey, $item.$itemOne);
            }
        }

        return $newSortKey;
    }

    public function __call($name, $arguments)
    {
        return $this->excel->{$name}(...$arguments);
    }

    public function formatStyle(Style $style)
    {
        $handle = $this->excel->getHandle();

        $format = new Format($handle);

        return $style->apply(function($formats) use($format){
            $formatter = $this->createFormatter();

            foreach ($formats as $letter => $value) {
                $formatter($letter)($value, $format);
            }

            return $format->toResource();
        });
    }
}
