<?php

infra_test(true);

use itlife\files\Xlsx;
use itlife\infra\ext\Ans;

$ans = array();

$data = Xlsx::init('*files/tests/resources/test.xlsx');

if (!$data) {
	return Ans::err($ans, 'Cant read test.xlsx');
}

$data = Xlsx::init('*files/tests/resources/test.csv');
if (!$data) {
	return Ans::err($ans, 'Cant read test.csv');
}
if (sizeof($data['childs']) != 1) {
	return Ans::err($ans, 'Cant read test.csv '.sizeof($data['childs']));
}


$num=ini_get('mbstring.func_overload');
if($num!=0){
	$ans['class']='bg-warning';
	return Ans::err($ans, 'mbstring.func_overload should be 0, not '.$num);
} else {
	$data = Xlsx::get('*files/tests/resources/test.xls');
	if (sizeof($data['childs'][0]['data']) != 30) {
		return Ans::err($ans, 'Cant read test.xls '.sizeof($data['childs'][0]['data']));
	}
}
return Ans::ret($ans, 'tpl, mht, docx, xls, xlsx read ok!');
