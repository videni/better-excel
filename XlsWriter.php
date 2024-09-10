<?php

namespace Videni\BetterExcel;

use Vtiful\Kernel\Excel;
use Vtiful\Kernel\Format;

class XlsWriter
{
    use XlsWriterFormatConstantsTrait;

    private $excel;

    public function __construct(string $filename, $options = [])
    {
        $constMemory = $options['const_memory']?? false;
        $defaultSheetName = $options['sheet_name']?? 'Sheet1';
        $enableZip64 = $options['enable_zip64']?? true;

        $this->excel = new Excel([
            // path is required, but here just simply let the Excel class complains
            'path' => $options['path']?? null,
        ]);

        if ($constMemory) {
            $this->excel->constMemory($filename, $defaultSheetName, (bool)$enableZip64);
        } else {
            $this->excel->fileName($filename, $defaultSheetName);
        }
    }

    public function writeHeader(array $columns)
    {
        $columnsLetters = $this->getPredefinedColumnLetters();

        $maxColumns = count($columnsLetters);
        $columnsCount = count($columns);
        if ($columnsCount > $maxColumns) {
            throw new \Exception(sprintf('The number of columns exceeds the maximum number(%d) of columns', $maxColumns));
        };

        $mappings = array_combine(array_slice($columnsLetters, 0, $columnsCount), $columns);

        foreach ($mappings as $letter => $column) {
            /**
             * @var \Videni\BetterExcel\Column $column
             */
            $column->setLetter($letter);
            // Set style for the whole column, for example, "A:A" is for the whole column A,
            $range = sprintf('%1$s:%1$s', $letter);

            $this->excel->setColumn(
                $range,
                $column->getStyle()->getWidth(),
                $this->formatStyle($column->getStyle())
            );
        }

        $labels = array_reduce($columns, function($carry, $column){
            $carry[] = $column->getLabel();

            return $carry;
         }, []);
        $this->excel->header($labels);
    }

    public function writeOneRow(array $data)
    {
        $newData = [];
        foreach($data as $cell) {
            $newData[] = method_exists($cell, 'render') ? $cell->render($this): $cell;
        }

        $this->excel->data([$newData]);
    }

    public function saveToFile($filename = null)
    {
        return $this->excel->output();
    }

    private function getPredefinedColumnLetters()
    {
        static $columnLetters = null ;
        if ($columnLetters) {
            return $columnLetters;
        }

        $letters = range('A', 'Z');

        $columnLetters = [];
        array_push($columnLetters, ...$letters);

        foreach ($letters as $letter1) {
            foreach ($letters as $letter2) {
                array_push($columnLetters, $letter1.$letter2);
            }
        }

        return $columnLetters;
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

            foreach ($formats as $path => $value) {
                $formatter($path)($value, $format);
            }

            return $format->toResource();
        });
    }
}
