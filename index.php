<?php

use infrajs\ans\Ans;
use infrajs\load\Load;
use infrajs\excel\Xlsx;
use infrajs\path\Path;
use infrajs\access\Access;
use infrajs\rest\Rest;


return Rest::get( function () {
	$isrc = Ans::GET('src');
	
	$fdata = Load::srcInfo($isrc);

	$src = Access::cache('files_get_php', function ($isrc) {
		$src = Path::theme($isrc);
		if ($src) {
			return $isrc;
		}

		$fdata = Load::srcInfo($isrc);
		$folder = Path::theme($fdata['folder']);

		if (!Path::theme($folder)) {
			return false;
		}
		array_map(function ($file) use (&$result, $fdata) {

			if ($file[0] == '.') {
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

		return $folder.$result;
	}, array($fdata['path']), isset($_GET['re']));

	$ans = array('src' => $isrc);
	if (!$src) {
		if (!Load::isphp()) {
			echo 'Требуется путь до файла можно без расширения ?src=! или /get/путь, /parse/путь, /init/путь. <br> Ещё так можно /-excel/get/group/Сферы/?src=~Параметры.xlsx';
			header('HTTP/1.0 404 Not Found');
		}
		return;
	}

	$fdata = Load::srcInfo($src);

	if (in_array($fdata['ext'], array('xls', 'xlsx', 'csv'))) {
		$ans = Xlsx::get($src);
		$ans['src'] = $isrc;
		return Ans::ans($ans);
	}
	if (!Load::isphp()) {
		header('HTTP/1.0 400 Bad Request');
	}

},'parse',[function() {
		echo 'Требуется путь до файла!';
	}, function (){
		$src = REST::getQuery();
		$r = explode('/',$src);
		array_shift($r);
		$src = implode('/',$r);
		$ans = array();

		$r = Path::isNest('~', $src);
		if (!$r) return Ans::err($ans, 'Передан небезопасный или некорректный путь');
		$file = Path::theme($src);
		$ext = Path::getExt($src);
		if (!in_array($ext, array('xlsx','xls'))) return Ans::err($ans, 'Не подходящее расширение файла');
		
		$data = Xlsx::parse($src);
		$ans['data'] = $data;
		return Ans::ret($ans);
	}
],'init',[function() {
		echo 'Требуется путь до файла!';
	}, 'group', function($a, $b, $group) {
		$src = Ans::GET('src');
		if (!$src) die('Требуется путь до файла!');
		//if (!Path::theme($src)) die('Файл не найден!');
		$data = Xlsx::init($src);
		$ans = array();
		$ans['group'] = $group;
		$ans['data'] = Xlsx::runGroups($data, function &($gr) use ($group) {
			if ($gr['title'] === $group) {
				return $gr;
			}
			$r = null;
			return $r;
		});
		return Ans::ret($ans);
	}, function (){
		$src = REST::getQuery();
		$r = explode('/',$src);
		array_shift($r);
		$src = implode('/',$r);

		$ans = array();
		$r = Path::isNest('~', $src);
		if (!$r) return Ans::err($ans, 'Передан небезопасный путь');
		$file = Path::theme($src);
		$ext = Path::getExt($src);
		if (!in_array($ext, array('xlsx','xls'))) return Ans::err($ans, 'Не подходящее расширение файла');
		
		$data = Xlsx::init($src);
		$ans['data'] = $data;
		return Ans::ret($ans);
	}
],'get',[function() {
		$src = Ans::GET('src');
		if (!$src) die('Требуется путь до файла!');
		$ans = array();
		$r = Path::isNest('~', $src);
		if (!$r) return Ans::err($ans, 'Передан небезопасный путь');
		$file = Path::theme($src);
		$ext = Path::getExt($src);
		if (!in_array($ext, array('xlsx','xls'))) return Ans::err($ans, 'Не подходящее расширение файла');
		
		
		$data = Xlsx::get($src);
		$ans['data'] = $data;
		return Ans::ret($ans);
	}, 'group', function($a, $b, $group) {
		$src = Ans::GET('src');
		if (!$src) die('Требуется путь до файла!');
		//if (!Path::theme($src)) die('Файл не найден!');
		$data = Xlsx::get($src);
		$ans = array();
		$ans['group'] = $group;
		$ans['data'] = Xlsx::runGroups($data, function &($gr) use ($group) {
			if ($gr['title'] === $group && $gr['type'] != 'book') {
				return $gr;
			}
			$r = null;
			return $r;
		});
		return Ans::ret($ans);
	}, function () {
		$src = REST::getQuery();
		$r = explode('/',$src);
		array_shift($r);
		$src = implode('/',$r);
		
		$ans = array();
		$r = Path::isNest('~', $src);
		if (!$r) return Ans::err($ans, 'Передан небезопасный путь');
		$file = Path::theme($src);
		$ext = Path::getExt($src);
		if (!in_array($ext, array('xlsx','xls'))) return Ans::err($ans, 'Не подходящее расширение файла');
		
		
		$data = Xlsx::get($src);
		$ans['data'] = $data;
		return Ans::ret($ans);
	}
]);
