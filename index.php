<?php

use infrajs\ans\Ans;
use infrajs\load\Load;
use infrajs\excel\Xlsx;
use infrajs\path\Path;
use infrajs\access\Access;

if (!is_file('vendor/autoload.php')) {
	chdir('../../../');
	require_once('vendor/autoload.php');
}

$isrc = Ans::GET('src');
$fdata = Load::srcInfo($isrc);
$src = Access::cache('files_get_php', function ($isrc) {
	$src = Path::theme($isrc);
	if ($src) {
		return $src;
	}
	$fdata = Load::srcInfo($isrc);
	$folder = Path::theme($fdata['folder']);

	if (!Path::theme($folder)) {
		return false;
	}
	array_map(function ($file) use (&$result, $fdata) {

		if ($file{0} == '.') {
			return;
		}
		$file=Path::toutf($file);
		$fd = Load::nameInfo($file);
		
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
	}, scandir(Path::theme($folder)));

	if (!$result) {
		return false;
	}

	return Path::theme($folder.$result);
}, array($fdata['path']), isset($_GET['re']));

$ans = array('src' => $isrc);
if (!$src) {
	if (!Load::isphp()) {
		header('HTTP/1.0 404 Not Found');
	}
	return;
}

$fdata = Load::srcInfo($src);


if (in_array($fdata['ext'], array('xls', 'xlsx', 'csv'))) {
	$ans = Xlsx::get($src);
	return Ans::ans($ans);
}
if (!Load::isphp()) {
	header('HTTP/1.0 400 Bad Request');
}