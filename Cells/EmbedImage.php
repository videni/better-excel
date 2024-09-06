<?php

namespace Modules\BetterExcel\Cells;

use Modules\BetterExcel\CellInfo;
use Modules\BetterExcel\XlsWriter;
use Modules\BetterExcel\Style;
use GuzzleHttp\Client;

class EmbedImage extends BaseCell
{
    public const DEFAULT_IMAGE_SIZE = 100;

    protected $path;
    protected $style;
    protected $imageSize;
    protected $imagick;

    public function __construct(string|\Closure $path, Style $style =null, $imageSize = null)
    {
        $this->imagick = new \Imagick();

        if (is_callable($path)) {
            $path->bindTo($this);
        }

        $this->path = $path;
        $this->style = $style;
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
        $callback = function(EmbedImage $self, $writer,CellInfo $info) use($imgUrl, $title) {
            $writer->insertUrl(
                $info->rowIndex,
                $info->columnIndex,
                $imgUrl,
                $title ?? "View original image",
                null,
                // Warning: 如果每个 cell 都设置一次样式， 会消耗极大的内存，注意不要这样去设置样式。
                // 我保留下面注释的代码是为了提醒你。其它类似的用法也一样。
                $self->getFormattedStyle($writer, $info)
            );

            try {
                $localPath = $self->downloadImageToTmpDir($imgUrl);
                // XlsWriter仅支持 png, jpg, 所以需要转换。
                // If the image is not proper format,  then convert it to PNG.
                $localPath = $self->convertToPNGFormatLocally($localPath, self::DEFAULT_IMAGE_SIZE);
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
    public static function fromPath($path, $imageSize = null, Style $style =null)
    {
        return new static($path, $style, $imageSize);
    }

    public function render(XlsWriter $writer, CellInfo $info): void
    {
        $localPath = $this->path;
        if (\is_callable($localPath)) {
            $localPath =  $localPath($this, $writer, $info);
        }
        if ($localPath === null || file_exists($localPath) === false) {
            return;
        }

        $style = $this->style ?? $info->style;
        $height = self::DEFAULT_IMAGE_SIZE;
        if ($style && null !== $style->getHeight()) {
            $height = $style->getHeight();
        }

        $writer
            ->insertImage($info->rowIndex,  $info->columnIndex, $localPath)
            // Set the image cell height, the width is set by the header, that is why
            // I don't set the width here.
            ->setRow(
                //必须要加 1， 为什么设置行号要加1？ask the author of XlsWriter library.
                sprintf('%s', $info->columnLetter. $info->rowIndex + 1),
                $height,
                // $style ? $writer->formatStyle($style): null
            );
    }

     /**
     * @return string | \Exception
     */
    public function downloadImageToTmpDir($imgUrl)
    {
        $tempPath = sys_get_temp_dir() . DIRECTORY_SEPARATOR . md5($imgUrl);
        if (file_exists($tempPath)) {
            return $tempPath;
        }

        $client = new Client();
        $client->get($imgUrl, ["sink" => $tempPath]);

        return $tempPath;
    }

    public function convertToPNGFormatLocally($imagePath, $imageSize)
    {
        $outputFile = dirname($imagePath).DIRECTORY_SEPARATOR. md5($imagePath) . '.png';
        if (file_exists($outputFile)) {
            return $outputFile;
        }

        $imagick = $this->imagick;

        $imagick->readImage($imagePath);

        // Convert and resize
        //@TODO: php-ext-xlswriter 支持嵌入图片时，resize 可以移除掉。
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
