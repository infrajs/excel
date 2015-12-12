<?php
namespace infrajs\doc;

use infrajs\ans\Ans;
use infrajs\path\Path;

$conf=&Infra::config('excel');
$conf=array_merge(Xlsx::$conf, $conf);
Xlsx::$conf=$conf;

Event::handler('oninstall', function () {
	Path::mkdir(Xlsx::$conf['cache']);
});