<?php

namespace Modules\BetterExcel\Cells;

use GuzzleHttp\Client;
use Modules\BetterExcel\XlsWriter;
use Modules\BetterExcel\Column;
use Modules\BetterExcel\Style;
use Webmozart\Assert\Assert;

class EmbedImage
{
    public const DEFAULT_IMAGE_SIZE = 100;

    protected $path;
    protected $style;
    protected $imageSize;

    public function __construct(string|\Closure $path, Style $style =null, $imageSize = null)
    {
        if (is_callable($path)) {
            $path->bindTo($this);
        }

        $this->path = $path;
        $this->style = $style ?? (new Style())->align('center', 'center');
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
        $imageProcessor = function($self, $writer, $rowIndex, $columnIndex) use($imgUrl, $title) {
            $writer->insertUrl(
                $rowIndex,
                $columnIndex,
                $imgUrl,
                $title ?? "View original image",
                null,
                $writer->formatStyle((new Style())->align('center', 'bottom'))
            );

            try {
                $localPath = $self->downloadImageToTmpDir($imgUrl);
                $localPath = $self->convertToThumbnailImageLocally($localPath, $self->imageSize);
            } catch(\Exception $e) {
                // Don't do anything, at least we have a link in the cell.
                return null;
            }

            return $localPath;
        };

        return new static($imageProcessor, $style, $imageSize);
    }

    /**
     * Insert a image to excel from the local path.
     *
     * @param string $path
     * @param int $imageSize
     * @param Style|null $style
     * @return self
     */
    public function fromPath(string $path, $imageSize = null, Style $style =null)
    {
        return new static($path, $style, $imageSize);
    }

    public function render(XlsWriter $writer, $rowIndex, $columnIndex, Column $column): void
    {
        Assert::isInstanceOf($writer, XlsWriter::class);

        $localPath = $this->path;
        if (\is_callable($localPath)) {
            $localPath =  $localPath($this, $writer, $rowIndex, $columnIndex);
        }

        if ($localPath === null || file_exists($localPath) === false) {
            return;
        }

        $height = self::DEFAULT_IMAGE_SIZE;
        if ($this->style && null !== $this->style->getHeight()) {
            $height = $this->style->getHeight();
        }

        $writer
            ->insertImage($rowIndex,  $columnIndex, $localPath)
            // Set the image cell height, the width is set by the header, that is why
            // I don't set the width here.
            ->setRow(
                //必须要加 1， 为什么设置行的高度要加 1？ask the author of XlsWriter library.
                sprintf($column->getLetter(). $rowIndex + 1),
                $height,
                $this->style ? $writer->formatStyle($this->style): null
            );
    }

    /**
     * @return string | \Exception
     */
    public function downloadImageToTmpDir($imgUrl)
    {
        $tempPath = sys_get_temp_dir() . DIRECTORY_SEPARATOR . basename($imgUrl);
        if (file_exists($tempPath)) {
            return $tempPath;
        }

        // TODO: Refactor this object according to flyweight pattern
        $client = new Client();
        $client->get($imgUrl, ["sink" => $tempPath]);

        return $tempPath;
    }

    public function convertToThumbnailImageLocally($imagePath, $imageSize)
    {
        $outputFile = dirname($imagePath).DIRECTORY_SEPARATOR. md5($imagePath) . '.png';
        if (file_exists($outputFile)) {
            return $outputFile;
        }

        // TODO: Refactor this object according to flyweight pattern,
        // so this object can be reused  to convert other images.
        $imagick = new \Imagick();

        $imagick->readImage($imagePath);

        $imagick->resizeImage($imageSize, $imageSize, \Imagick::FILTER_LANCZOS, 1);
        $imagick->setImageFormat('png');

        $imagick->writeImage($outputFile);

        // Clear the resource taken by the image itself,
        // the $imagick object can be used to process other images.
        $imagick->clear();

        $imagick->destroy();

        return $outputFile;
    }
}
