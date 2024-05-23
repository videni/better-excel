<?php

namespace Modules\BetterExcel\Cells;

use GuzzleHttp\Client;
use Modules\BetterExcel\XlsWriter;
use Modules\BetterExcel\Column;
use Modules\BetterExcel\Style;
use Webmozart\Assert\Assert;

class EmbedImage
{
    public const DEFAULT_IMAGE_CELL_HEIGHT = 100;

    private $url;
    private $title;
    private $style;

    public function __construct(string $url, Style $style =null, $title = null)
    {
        $this->url = $url;
        $this->style = $style ?? (new Style())->align('center', 'center');
        $this->title = $title ?? "View original image";
    }

    public static function fromImageUrl(string $imgUrl, Style $style =null, $title = null)
    {
        return new static($imgUrl, $style, $title);
    }

    public function render($writer, $rowIndex, $columnIndex, Column $column)
    {
        Assert::isInstanceOf($writer, XlsWriter::class);

        $writer->insertUrl(
            $rowIndex,
            $columnIndex,
            $this->url,
            $this->title,
            null,
            $writer->formatStyle((new Style())->align('center', 'bottom'))
        );

        try {
            $url = $this->url;
            $localPath = $this->downloadImageToTmpDir($url);
            $localPath = $this->convertTo100x100ThumbnailImageLocally($localPath);
        } catch(\Exception $e) {
            return null;
        }

        $height = self::DEFAULT_IMAGE_CELL_HEIGHT;
        if ($this->style && $height = $this->style->getWidth()) {
            $height =  $height;
        }

        $writer
            ->insertImage($rowIndex,  $columnIndex, $localPath)
            ->setRow(
                sprintf($column->getColumnIndex(). $rowIndex),
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

    private function convertTo100x100ThumbnailImageLocally($imagePath)
    {
        $outputFile = dirname($imagePath).DIRECTORY_SEPARATOR. md5($imagePath) . '.png';
        if (file_exists($outputFile)) {
            return $outputFile;
        }

        // TODO: Refactor this object according to flyweight pattern
        $image = new \Imagick();

        $image->readImage($imagePath);

        // 调整图像大小到100x100
        $image->resizeImage(100, 100, \Imagick::FILTER_LANCZOS, 1);
        $image->setImageFormat('png');

        $image->writeImage($outputFile);

        // 清理
        $image->clear();
        $image->destroy();

        return $outputFile;
    }
}
