<?php
use infrajs\excel\Xlsx;
use infrajs\router\Router;
use infrajs\ans\Ans;

if (!is_file('vendor/autoload.php')) {
	chdir('../../../');
	require_once('vendor/autoload.php');
	Router::init(); //Требуется автоматическая установка
}

$ans = array();
//Три уровня разбора данных

//1 Минимум требований к структуре Excel документа, данные "как есть"
$data = Xlsx::parse('vendor/infrajs/excel/test.xlsx');

if (sizeof($data) != 19) return Ans::err($ans, 'Некорректный результат Xlsx::parse '.sizeof($data));


//2 Простая структура - Распознаются заголовки таблицы, описание таблицы, структура групп. Можно применять Xlsx::runPoss и Xlsx::runGroups
$data = Xlsx::get('vendor/infrajs/excel/test.xlsx');
if (sizeof($data) != 7) return Ans::err($ans, 'Некорректный результат Xlsx::get '.sizeof($data));


//3 Оптимизировання структура. Обязательна колонка Артикул, объединение групп. Большой список опций вторым аргументом
$data = Xlsx::init('vendor/infrajs/excel/test.xlsx');
if (sizeof($data['childs']) != 1) return Ans::err($ans, 'Некорректный результат Xlsx::init '.sizeof($data));


return Ans::ret($ans);