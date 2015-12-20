<?php
namespace infrajs\excel;

use infrajs\event\Event;
use infrajs\path\Path;
use infrajs\infra\Config;

$conf=&Config::get('excel');
$conf=array_merge(Xlsx::$conf, $conf);
Xlsx::$conf=$conf;

Event::handler('oninstall', function () {
	Path::mkdir(Xlsx::$conf['cache']);
});