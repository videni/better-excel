<?php

namespace Modules\BetterExcel\Cells;

use Modules\BetterExcel\CellInfo;
use Modules\BetterExcel\XlsWriter;
use Modules\BetterExcel\Style;
use Webmozart\Assert\Assert;

class EmbedImage
{
    public const DEFAULT_IMAGE_SIZE = 100;

    protected $path;
    protected $style;
    protected $imageSize;
    protected $excelImageProcessor;

    public function __construct(string|\Closure $path, Style $style =null, $imageSize = null)
    {
        $this->excelImageProcessor = new ExcelImageProcessor();

        if (is_callable($path)) {
            $path->bindTo($this);
        }

        $this->path = $path;
        $this->style = $style ?? (new Style())->align('left', 'bottom');
        $this->imageSize = $imageSize ?? self::DEFAULT_IMAGE_SIZE;
    }

    /**
     *  Download the image from the url and insert it to the excel.
     *
     * @param string $imgUrl
     * @param int|null $imageSize
     * @param Style|null $style
     * @param string|Null $title
     * @return self
     */
    public static function fromUrl(string $imgUrl, $imageSize = null, Style $style =null, $title = null)
    {
        $callback = function(EmbedImage $self, $writer, $rowIndex, $columnIndex) use($imgUrl, $title) {
            $writer->insertUrl(
                $rowIndex,
                $columnIndex,
                $imgUrl,
                $title ?? "View original image",
                null,
                $writer->formatStyle((new Style())->align('center', 'bottom'))
            );

            try {
                $localPath = $self->excelImageProcessor->downloadImageToTmpDir($imgUrl);
                $localPath = $self->excelImageProcessor->convertToPNGFormatLocally($localPath);
            } catch(\Exception $e) {
                // Don't do anything, at least we have a link in the cell.
                return null;
            }

            return $localPath;
        };

        return new static($callback, $style, $imageSize);
    }

    /**
     * Insert a image to excel from the local path.
     *
     * @param  $path
     * @param int $imageSize
     * @param Style|null $style
     * @return self
     */
    public function fromPath($path, $imageSize = null, Style $style =null)
    {
        return new static($path, $style, $imageSize);
    }

    public function render(XlsWriter $writer, CellInfo $info): void
    {
        Assert::isInstanceOf($writer, XlsWriter::class);

        $localPath = $this->path;
        if (\is_callable($localPath)) {
            $localPath =  $localPath($this, $writer, $info->rowIndex, $info->columnIndex);
        }
        if ($localPath === null || file_exists($localPath) === false) {
            return;
        }

        $style = $this->style ?? $info->style;
        $height = self::DEFAULT_IMAGE_SIZE;
        if ($style && null !== $style->getHeight()) {
            $height = $style->getHeight();
        }

        [$scaleWidth, $scaleHeight] = $this->excelImageProcessor->calculateImageScaleFactor($localPath, $this->imageSize);

        $writer
            ->insertImage($info->rowIndex,  $info->columnIndex, $localPath, $scaleWidth, $scaleHeight)
            // Set the image cell height, the width is set by the header, that is why
            // I don't set the width here.
            ->setRow(
                //必须要加 1， 为什么设置行号要加1？ask the author of XlsWriter library.
                sprintf('%s', $info->columnLetter. $info->rowIndex + 1),
                $height,
                // $style ? $writer->formatStyle($style): null
            );
    }
}
