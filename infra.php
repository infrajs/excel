<?php
namespace infrajs\excel;

use infrajs\event\Event;
use infrajs\path\Path;
use infrajs\infra\Infra;

$conf=&Infra::config('excel');
$conf=array_merge(Xlsx::$conf, $conf);
Xlsx::$conf=$conf;

Event::handler('oninstall', function () {
	Path::mkdir(Xlsx::$conf['cache']);
});