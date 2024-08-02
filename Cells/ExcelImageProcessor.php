<?php

namespace Modules\BetterExcel\Cells;

use GuzzleHttp\Client;

class ExcelImageProcessor
{
    protected $imagick;

    protected $client;

    /**
     *
     * 按照最大尺寸计算图片的缩放比例
     *
     * @param string $localPath
     * @param integer $maxImageSize
     * @return [$scaleWidth, $scaleHeight]
     */
    public function calculateImageScaleFactor(string $localPath, int $maxImageSize)
    {
        $imagick = $this->getImagick();

        $imagick->readImage($localPath);

        $width = $imagick->getImageWidth();
        $height = $imagick->getImageHeight();

        $imagick->clear();

        // If the image is smaller than the imageSize, then don't scale it.
        if ($width < $maxImageSize && $height < $maxImageSize) {
            return [1, 1];
        }

        $scaleWidth = $maxImageSize / $width;
        $scaleHeight = $maxImageSize/ $height;

        // Make the image fit the imageSize
        $ratio = $height / $width;
        if ($ratio > 1) {
            $scaleWidth = $scaleHeight/$ratio;
        } else {
            $scaleHeight = $scaleWidth * $ratio;
        }

        return [$scaleWidth, $scaleHeight];
    }

    /**
     * 下载图片到系统tmp目录, 返回图片的本地路径
     *
     * @return string | \Exception
     */
    public function downloadImageToTmpDir($imgUrl)
    {
        $tempPath = sys_get_temp_dir() . DIRECTORY_SEPARATOR . basename($imgUrl);
        if (file_exists($tempPath)) {
            return $tempPath;
        }

        $client = $this->getClient();
        $client->get($imgUrl, ["sink" => $tempPath]);

        return $tempPath;
    }

    /**
     * 将图片转为PNG格式， 返回转化后的图片路径
     *
     * 注意，Excel只支持PNG， JGP 两种格式。
     *
     * @param string $imagePath
     * @return  $localPath
     */
    public function convertToPNGFormatLocally($imagePath)
    {
        // If the image is already proper format, then return the path.
        if (in_array(pathinfo($imagePath, PATHINFO_EXTENSION), ['png', 'jpg'])) {
            return $imagePath;
        }

        $outputFile = dirname($imagePath).DIRECTORY_SEPARATOR. md5($imagePath) . '.png';
        if (file_exists($outputFile)) {
            return $outputFile;
        }
        $imagick = $this->getImagick();

        $imagick->readImage($imagePath);
        $imagick->setImageFormat('png');
        $imagick->writeImage($outputFile);

        $this->imagick->clear();

        return $outputFile;
    }

    private function getImagick()
    {
        if ($this->imagick === null) {
            $this->imagick = new \Imagick();
        }

        return $this->imagick;
    }

    private function getClient()
    {
        if ($this->client === null) {
            $this->client = new Client();
        }

        return $this->client;
    }
}
