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

    private $path;
    private $style;
    private $imageSize;

    public function __construct(string|\Closure $path, Style $style =null, $imageSize = null)
    {
        $this->path = $path;
        $this->style = $style ?? (new Style())->align('center', 'center');
        $this->imageSize = $imageSize ?? self::DEFAULT_IMAGE_SIZE;
    }

    public static function fromUrl(string $imgUrl, $imageSize = null, Style $style =null, $title = null)
    {
        return new static(function($writer, $rowIndex, $columnIndex) use($imgUrl, $title) {
            $writer->insertUrl(
                $rowIndex,
                $columnIndex,
                $imgUrl,
                $title ?? "View original image",
                null,
                $writer->formatStyle((new Style())->align('center', 'bottom'))
            );

            try {
                $localPath = $this->downloadImageToTmpDir($imgUrl);
                $localPath = $this->convertToThumbnailImageLocally($localPath, $this->imageSize);
            } catch(\Exception $e) {
                // Don't do anything, at least we have a link in the cell.
                return null;
            }

            return $localPath;
        }
        , $style, $title, $imageSize);
    }

    public function fromPath(string $path, $imageSize = null, Style $style =null)
    {
        return new static($path, $style, $imageSize);
    }

    public function render($writer, $rowIndex, $columnIndex, Column $column)
    {
        Assert::isInstanceOf($writer, XlsWriter::class);

        $localPath = \is_callable($this->path) ? call_user_func($this->path, [$writer, $rowIndex, $columnIndex]): $this->path;

        if ($localPath === null || file_exists($localPath) === false) {
            return;
        }

        $height = self::DEFAULT_IMAGE_SIZE;
        if ($this->style && $height = $this->style->getHeight()) {
            $height = $height;
        }

        $writer
            ->insertImage($rowIndex,  $columnIndex, $localPath)
            // Set the image cell height, the width is set by the header, that is why
            // I don't set the width here.
            ->setRow(
                sprintf($column->getLetter(). $rowIndex),
                $height,
                $this->style ? $writer->formatStyle($this->style): null
            );
    }

    /**
     * @return string | \Exception
     */
    private function downloadImageToTmpDir($imgUrl)
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

    private function convertToThumbnailImageLocally($imagePath, $imageSize)
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
