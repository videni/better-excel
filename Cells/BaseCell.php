<?php
namespace Modules\BetterExcel\Cells;

use Modules\BetterExcel\CellInfo;
use Modules\BetterExcel\Style;
use Modules\BetterExcel\XlsWriter;

abstract class BaseCell
{
    /**
     * Cell 的样式
     *
     * @var Style
     */
    protected $style = null;

    public abstract function render(XlsWriter $writer, CellInfo $info): void;

    /**
     * 如果当前 cell 有 cell 级别的style则优先使用它，否则使用 header中设置的style。
     * 注意： Cell级别的 style，你需要确保同一个实例被所有相关的 Cell共享，否则，
     * 导出大量数据时，会消耗大量的内存。
     *
     * @param XlsWriter $writer
     * @param CellInfo $info
     * @return mixed
     */
    public function getFormattedStyle(XlsWriter $writer, CellInfo $info)
    {
        // 如果当前 cell 有设置 style ，则优先使用它。
        $style = $this->style;
        if ($style) {
            $formattedStyle = $style->getFormattedStyle();

            return $formattedStyle ?: $writer->formatStyle($style);
        }

        // 否则，使用 header 设置的样式。
        return $info->style?->getFormattedStyle();
    }

    public function getStyle(): ?Style
    {
        return $this->style;
    }
}
