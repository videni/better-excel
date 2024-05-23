<?php

namespace Modules\BetterExcel\Cells;

use GuzzleHttp\Client;
use Modules\File\Models\File;
use Modules\File\Repositories\Interfaces\IFileRepository;
use Illuminate\Support\Facades\Log;
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
        $this->title = $title ?? "查看原图";
    }

    /**
     *  A factory method to create an ExcelEmbedImage instance from the File model
     *
     * @param File $file
     * @param int $row
     * @param int $column
     * @return self
     */
    public static function fromFileModel(
        File $file,
        Style $style =null,
        $title = null
    ): self {
        $url = !empty($file->path)? self::toAbsoluteOssUrl($file->path) : $file->raw_url;

        return new static($url, $style, $title);
    }

    public static function fromFullImageUrl(string $imgUrl, Style $style =null, $title = null)
    {
        return new static($imgUrl, $style, $title);
    }

    public static function fromFileRawUrl(string $imgUrl, Style $style =null, $title = null)
    {
        $file = app(IFileRepository::class)->getFileByRawUrl($imgUrl);

        return $file ? static::fromFileModel($file, $style, $title): null;
    }

    public static function fromFileId(int $fileId, Style $style =null, $title = null)
    {
        $file = app(IFileRepository::class)->find($fileId);

        return $file ? static::fromFileModel($file, $style, $title): null;
    }

    public function render($writer, $rowIndex, $columnIndex, Column $column)
    {
        Assert::isInstanceOf($writer, XlsWriter::class);

        $writer->insertUrl(
            $rowIndex,
            $columnIndex,
            $this->url,
            $this->title
        );

        try {
            $isOnCDN = $this->isImageOnAliOSS($this->url);
            $url = $this->url;
            if ($isOnCDN) {
                $url = $this->append100x100ThumbnailQuery($url);
            }

            $localPath = $this->downloadImageToTmpDir($url);
            if (!$isOnCDN) {
                $localPath = $this->convertTo100x100ThumbnailImageLocally($localPath);
            }

        } catch(\Exception $e) {
            Log::info("Failed to process image from $this->url: " . $e->getMessage());

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

    public function isImageOnAliOSS($url)
    {
        $domains = config("import.order.cdn_domains");

        foreach($domains as $domain) {
            if (strpos($url, $domain) !== false) {
                return true;
            }
        }

        return false;
    }

    public static function toAbsoluteOssUrl($path)
    {
        return config("filesystems.disks.oss")["ossHost"] . "/" . $path;
    }

    private function append100x100ThumbnailQuery($url)
    {
        return $url . "?x-oss-process=image/resize,h_100,w_100/format,png";
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
