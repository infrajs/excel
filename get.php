<?php

/*
Copyright 2008-2010 ITLife, Ltd. Togliatti, Samara Oblast, Russian Federation. http://itlife-studio.ru

History
23.04.2010
Скрипт получает src без расширения и без цифры сортировки.... а возвращает html
25.04.2010
Добавлено кэширование modified

09.05.2010
Добавлена поддерж php файлов и возможность передачи get параметров запрашиваемому файлу
*/

//..'xls'=>'?*pages/xls/xls.php?src='

use infrajs\infra\ext\Ans;
use infrajs\excel;

$isrc = infra_toutf(urldecode($_SERVER['QUERY_STRING']));
infra_admin_modified();
$fdata = infra_srcinfo($isrc);
$src = infra_admin_cache('files_get_php', function ($isrc) {
	$src = infra_theme($isrc);
	if ($src) {
		return $src;
	}
	$fdata = infra_srcinfo($isrc);
	$folder = infra_theme($fdata['folder']);

	if (!infra_theme($folder)) {
		return false;
	}
	array_map(function ($file) use (&$result, $fdata) {

		if ($file{0} == '.') {
			return;
		}
		$file=infra_toutf($file);
		$fd = infra_nameinfo($file);
		
		if ($fdata['id'] && $fdata['id'] != $fd['id']) {
			return;
		}
		if ($fdata['name'] && $fdata['name'] != $fd['name']) {
			return;
		}
		
		if ($fdata['ext'] && $fdata['ext'] != $fd['ext']) {
			return;
		} elseif ($result) {
			//Расширение не указано и уже есть результат
			//Исключение.. расширение tpl самое авторитетное
			if ($fd['ext'] != 'tpl') {
				return;
			}
		}
		$result = $file;
	}, scandir(infra_theme($folder)));

	if (!$result) {
		return false;
	}

	return infra_theme($folder.$result);
}, array($fdata['path']), isset($_GET['re']));

$ans = array('src' => $isrc);
if (!$src) {
	if (!infra_isphp()) {
		header('HTTP/1.0 404 Not Found');
	}
	return;
}

$fdata = infra_srcinfo($src);


if (in_array($fdata['ext'], array('xls', 'xlsx', 'csv'))) {
	$ans = excel\Xlsx::get($src);
	return Ans::ans($ans);
}
if (!infra_isphp()) {
	header('HTTP/1.0 400 Bad Request');
}
