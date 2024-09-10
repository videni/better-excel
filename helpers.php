<?php

use Videni\BetterExcel\Style;

if (!function_exists('style')) {
    function style($formats = [])
    {
        return Style::fromArray($formats);
    }
}
