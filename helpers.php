<?php

use Modules\BetterExcel\Style;

if (!function_exists('style')) {

    function style($formats = [])
    {
        return Style::fromArray($formats);
    }
}
