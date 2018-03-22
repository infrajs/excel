<?php
namespace infrajs\excel;

use infrajs\path\Path;
use infrajs\load\Load;
use infrajs\each\Each;
use infrajs\each\Fix;
use infrajs\cache\Cache as OldCache;
use infrajs\once\Once;
use infrajs\config\Config;
use akiyatkin\boo\Cache;
use akiyatkin\boo\MemCache;
use infrajs\sequence\Sequence;
/*
* xls методы для работы с xls документами. 
*
* Помимо получения данных в первозданном виде, 
* модуль также реализует определённый синтаксис в Excel для построения иерархичной структуры с данными.
*

* **Использование**

	//Получаем данные из Excel "как есть"
	$data=xls_parse('~Главное меню.xls');
	//или
	$data=xls_make('~Главное меню.xls');
	//Создаём объект с вложенными группами root->book->sheet данные на страницах ещё не изменялись, 
	//но сгрупировались
	//descr - всё что до head
	//head - первая строка в которой больше 2х заполненых ячеек
	//data - всё что после head
	xls_processDescr($data);//descr приводится к виду ключ значение
	xls_run($data,function($group){//Бежим по всем группам
		unset($group['parent']);//Удалили рекурсивное свойсто parent
		for ($i=0, $l=sizeof($group['data']);$i<$l;$i++){
			$pos=$group['data'][$i];
			unset($pos['group']);//Удалили рекурсивное свойсто group
		}
	});
	$data=xls_init(path,conf)
*/


/*var pathlib=require('path');
var util=require('util');
var csv=require('node-csv');
var crypto=require('crypto');
var fs=require('fs');
csv=csv.createParser(',','"','"');*/

function &xls_parseTable($path, $list)
{
	$data = xls_parse($path, $list);
}
function &xls_parseAll($path)
{
	return Xlsx::parseAll($path);
}
function &xls_parse($path, $list = false)
{
	$data = &xls_parseAll($path);
	if (!$list) {
		$list = Each::foro($data, function &(&$v, $k) {
			return $k;
		});
	}
	return $data[$list];
}

function &xls_make($path, $title = false)
{

	if (is_string($path)) {
		$datamain = xls_parseAll($path);	
		if (!$datamain) return $datamain;
		if (!$title) {
			$p = Load::srcInfo($path);
			$title = $p['name'];
			$title = Path::toutf($title);
		}
	} else {
		$datamain = $path;
	}


	$parent = false;

	$groups = &_xls_createGroup($title, $parent, 'book');
	
	foreach ($datamain as $title => $data) {
		if (!$data) continue;
		//Бежим по листам
		if ($title{0} === '.') continue; //Не применяем лист у которого точка в начале имени
		$argr = array();//Чтобы была возможность определить следующую группу и при этом работать со ссылкой и не переопределять предыдущую
		$argr[0] = &_xls_createGroup($title, $groups, 'list');
		if (!$argr[0]) continue;
		$groups['childs'][] = &$argr[0];

		$head = false;//Заголовки ещё не нашли
		$pgpy = false;//ПГПЯ Признак группы пустая ячейка в строке... а этом свойстве будет индекс ПГПЯ
		$wasdata = false;//Были ли до этого данные
		$wasgroup = false;
		//var empty=0;//Количество пустых строк
		$first_index = 0;

		foreach ($data as $i => $row) {
			//Бежим по строкам
			//Each::foro($data,function(&$row,$i) use(&$head,&$pgpy,&$wasdata,&$wasgroup,&$argr,&$first_index){
			$count = 0;
			//$group=&$argr[0];//Группа может появится среди данных в листах
			//echo $group['title'].'<br>';
			//if(!$row) continue;
			
			foreach ($row as $cell) {
				if ($cell) $count++;
			}

			if (!$head) {
				foreach ($row as $b => $rowcell) {
					$row[$b] = preg_replace('/\n/', '', $row[$b]);
					$row[$b] = preg_replace('/\s+$/', '', $row[$b]);
					$row[$b] = preg_replace('/^\s+/', '', $row[$b]);
				}
				$head = ($count > 2);//Больше 2х не пустых ячеек будет заголовком
				foreach ($row as $first_index => $first_value) {
					break;
				}
				if ($head) {
					//Текущий row и есть заголовок
					$argr[0]['head'] = $row;
				} else {
					if ($first_value == 'ПГПЯ') {
						//Признак группы пустая ячейка номер этой ячейки
						$pgpy = $row[$first_index + 1] - 1;//Индекс пустой ячейки
					} else {
						if ($first_value) {
							$argr[0]['descr'][] = $row;
						}
					}
				}
			} else {
				$isnewgroup = (isset($row[$first_index]) && ($count == 1) && mb_strlen($row[$first_index]) > 1);//Если есть только первая ячейка и та длинее одного символа

				if (!$isnewgroup && $pgpy && mb_strlen($row[$first_index]) !== 1) {
					//один символ в первой ячейке имеет специальное значение выхода на уровень вверх

					$roww = array_values($row);
					$isnewgroup = !$roww[$pgpy];
				}
				if ($isnewgroup) {
					if ($wasdata && !empty($argr[0]['parent']) && $argr[0]['parent']['type'] != 'book') {
						$argr = array(&$argr[0]['parent']);//Если уже были данные то поднимаемся наверх
					}
					$g = array();
					$g[0] = &_xls_createGroup($row[$first_index], $argr[0], 'row', $row);//Создаём новую группу
					if (!$g[0]) continue;
					$g[0]['parent']['childs'][] = &$g[0];
					$wasgroup = true;
					$wasdata = false;

					$pdescr = $g[0]['parent']['descr'];
					unset($pdescr['Наименование']);//Наименование родительской группы не наследуем
					$g[0]['descr'] = array_merge($pdescr, $g[0]['descr']);

					$g[0]['head'] = &$g[0]['parent']['head'];
					$argr = array(&$g[0]);

//$group=&$g;//Теперь ссылка на новую группу и следующие данные будут добавляться в неё
					//Новая ссылка забивает на старую, простое присвоение это новое место куда указывает ссылка
				} else {
					if (!empty($row[$first_index]) && $count === 1 && strlen($row[$first_index]) === 1) {
						//подъём на уровень выше
						if($argr[0]['parent']['type'] != 'book') {
							if (!empty($argr[0]['parent'])) {
								$argr = array(&$argr[0]['parent']);
								//echo '<b>'.$group['title'].'</b><br>';
							}
						}
					} else {
						$wasdata = true;
						$argr[0]['data'][] = $row;
					}
				}
			}
		}
	}
	return $groups;
}
function &xls_runPoss(&$data, $callback, $back=false)
{
	return xls_runGroups($data, function &(&$group) use ($back, &$callback) {
		$r = null;
		if(empty($group['data'])) return $r;
		foreach ($group['data'] as $i => &$pos){
			$r = $callback($pos, $i, $group);
			if (!is_null($r)) return $r;
		}
		return $r;
	});
}

function &_xls_createGroup($title, &$parent, $type, &$row = false)
{
	$tparam = '';
	$descr = array();
	$miss = false;
	$t = explode(':', $title);
	if (!$t[0] && $parent) {
		//Когда начинается с двоеточия
		array_shift($t);
		$title = implode(':', $t);
		$first_index = null;
		foreach ($parent['descr'] as $first_index => $first_value) {
			break;
		}
		$index = Each::forr($parent['descr'], function &(&$row, $i) use ($first_index, $title) {
			if ($row[$first_index] == 'Описание') {
				$row[$first_index + 1] .= '<br>'.$title;

				return $i;
			}
			$r = null;

			return $r;
		});
		if (!is_null($index)) {
			$parent['descr'][$index] = array('Описание',$title);
		} else {
			array_push($parent['descr'], array('Описание', $title));
		}
		$r = false;
		return $r;
	} else {
		if (sizeof($t) > 1) {
			$title = array_shift($t);
			if ($title == 'Производитель') {
				//Производитель:KUKA будет означать что у текущей группы указан производитель
				$title = implode(':', $t);
				$tparam = '';
				array_push($descr, array('Производитель', $title));
				$miss = true;
			} else {
				$tparam = implode(':', $t);
			}
		}
	}
	$title = preg_replace('/["+\']/', ' ', $title);
	$title = preg_replace('/[\\/\\\\]/', '', $title);
	$title = preg_replace('/^\s+/', '', $title);
	$title = preg_replace('/\s+$/', '', $title);
	$title = preg_replace('/\s+/', ' ', $title);


	if ($type == 'set') $pitch = 0;
	if ($type == 'book') $pitch = 1;
	if ($type == 'list') $pitch = 2;
	if ($type == 'row') $pitch = 3;
	// title=title.toUpperCase();
	//array_push($descr, array('Наименование', $title));
	$res = array(
		//'tparam'=>false,
		//'groups'=>false,//Количество групп вместе с текущей
		//'count'=>false,
		//'row' => $row,//Вся строка группы
		'pitch' => $pitch, //Шаг от верхнего уровня
		'miss' => $miss,//Группу надо расформировать, но мы не знаем ещё есть ли в ней позиции
		'type' => $type,
		'parent' => &$parent,
		'title' => (string) $title,
		'head' => array(),
		'descr' => &$descr,
		'data' => array(),
		'childs' => array(),
	);
	if ($tparam) {
		$res['tparam'] = $tparam;
	}//Параметр у группы Сварка:asdfasd что угодно
	return $res;
}

function xls_processPoss(&$data, $ishead = false)
{
//
	//используется data head


	xls_runGroups($data, function &(&$data) use ($ishead) {
		$r = null;
		if (!empty($data['head'])) {
			$head = &$data['head'];
		} else {
			return $r; //Значит и данных нет
		}

		Each::forr($data['data'], function &(&$pos, $i, &$group) use (&$head, &$data) {

			$p = array();
			
			foreach($pos as $k=>$propvalue) {
				if (empty($head[$k])) return $r;
				$propname = $head[$k];
				
				if ($propname{0} == '.') {
					continue;
				}//Колонки с точкой скрыты
				if ($propvalue == '') {
					continue;
				}
				if ($propvalue{0} == '.') {
					continue;
				}//Позиции у которых параметры начинаются с точки скрыты

				$propvalue = trim($propvalue);
				//$propvalue=preg_replace('/\s+$/','',$propvalue);
				//$propvalue=preg_replace('/^\s+/','',$propvalue);
				if (!$propname) {
					continue;
				}
				$p[$propname] = $propvalue;
			}
			//$p['parent'] = &$data;//Рекурсия
			$pos = $p;
			$group[$i] = &$pos;


			return $r;
		});
		if (!$ishead) {
			unset($data['head']);
		}
		$r = null; return $r;
	});
	
}
function xls_processPossFilter(&$data, $props)
{
	//Если Нет какого-то свойства не учитываем позицию
	xls_runGroups($data, function &(&$data) use (&$props) {
		$d = array();
		Each::forr($data['data'], function &(&$pos) use (&$props, &$d) {

			if (!Each::forr($props, function &($name) use ($pos) {
				$r = null;
				if (!$pos[$name]) {
					$r = true;
					return $r;
				}
				return $r;
			})) {
				$d[] = &$pos;
			}
			$r = null;

			return $r;
		});
		$data['data'] = $d;
		$r = null; return $r;
	});
}

function xls_processPossBe(&$data, $check1, $check2)
{
	//Если у позиции нет поля check1.. то оно будет равнятся полю check2
	//используется data
	xls_runPoss($data, function &(&$pos) use ($check1, $check2) {
		$r = null;
		if (is_null($pos[$check1])) {
			$pos[$check1] = $pos[$check2];
		}
		if (is_null($pos[$check2])) {
			$pos[$check2] = $pos[$check1];
		}
		return $r;
	});
}
function xls_processPossFS(&$data, $props)
{
	xls_runPoss($data, function &(&$pos) use (&$props) {
		return Each::foro($props, function &($name, $key) use (&$pos) {
			$r = null;
			if (isset($pos[$key])) {
				$pos[$name] = Path::encode($pos[$key]);
			}
			return $r;
		});
	});
};
function xls_processPossMore(&$data, $props)
{
	xls_runPoss($data, function &(&$pos, $i, &$group) use (&$props) {
		$r = null;
		$p = array();
		$more = array();

		$prop = array();
		Each::forr($props, function &($name) use (&$prop) {
			$prop[$name] = true;
			$r = null;

			return $r;
		});

		Each::foro($pos, function &(&$val, $name) use (&$p, &$prop, &$more) {
			if (!empty($prop[$name])) {
				$p[$name] = &$val;
			} else {
				$more[$name] = &$val;
			}
			$r = null;

			return $r;
		});
		if ($more) {
			$p['more'] = &$more;
		}
		$group['data'][$i] = &$p;
		return $r;
	});
}


function &xls_runGroups(&$data, $callback, $back = false, $i = 0, &$group = false)
{
	if (!$back) {
		$r = &$callback($data, $i, $group);
		if (!is_null($r)) return $r;
	}

	for ($i = 0; $i < sizeof($data['childs']); $i++) {
		$r = &xls_runGroups($data['childs'][$i], $callback, $back, $i, $data);
		if (!is_null($r)) return $r;
	}	

	if ($back) {
		$r = &$callback($data, $i, $group);
		if (!is_null($r)) return $r;
	}

	return $r;
}
function xls_processGroupFilter(&$data)
{
	/*$g1 = array(
		'title' => 'Группа',
		'childs' => [],
		'data' => [['article'=>'Тадам']]
	);
	$g2 = array(
		'title' => 'Группа',
		'childs' => [&$g1],
		'data' => [['article'=>'Парам']]
	);
	$data = array(
		'title' => 'Каталог',
		'childs' => [
			&$g2
		]
	);
	$g1['parent'] = &$g2;
	$g2['parent'] = &$data;*/
	
	

	
	$all = array();
	Xlsx::runGroups($data, function &(&$gr, $i, &$parent) use (&$all) {
		$title = mb_strtolower($gr['title']);
		if (!isset($all[$title])) {
			$all[$title] = array('orig' => &$gr, 'list' => array());
		} else {
			//Ну вот и нашли повторение
			$all[$title]['list'][] = array('gr' => &$gr, 'parent' => &$parent);
			//xls_merge($all[$title],$gr);
			//у некой прошлой группы появляются новые childs.. но мы всё ещё бежим по какому-то его childs и новые добавленные будут проигнорированны
			//return new Fix('del');
		}
		$r = null; return $r;
	});

	Each::foro($all, function &(&$des, $title) use(&$data, &$all){
		

		Each::forr($des['list'], function &($gr) use (&$des, $title, &$all) {
			xls_merge($des['orig'], $gr['gr']); //Переносим данные

			$gr['parent']['childs'] = array_values(array_filter($gr['parent']['childs'], function(&$g) use (&$gr){
				//Нужно удалить ссылки на старую группу в childs родителя
				if (Each::isEqual($g, $gr['gr'])) {
					return false;
				}
				return true;
			}));
			
			
			
			
			/*Each::forr($gr['parent']['childs'], function &(&$g) use (&$gr) {
				//Нужно удалить ссылки на старую группу в childs родителя
				if (Each::isEqual($g, $gr)) {
					$r = new Fix('del', true);
					return $r;
				}
				$r = null;
				return $r;
			});*/

			$r = null;
			return $r;
		});
		
		$r = null;
		return $r;
	});
	

	/*//$cat=$data['childs'][0];
	$cat=$data;
	unset($cat['parent']);
	Each::forr($cat['childs'], function &(&$g){
		//if(!is_string($g['parent']))
		$g['parent']=&$g['parent']['title'];
		//unset($g['parent']);
		$g['childs']=sizeof($g['childs']);
		$g['data']=sizeof($g['data']);
		$r = null; return $r;
	});
	echo '<pre>';
	print_r($cat);
	exit;
	/*
	xls_runGroups($data,function(&$gr,$i,&$group){//Удаляем пустые группы
		if(!$group) return;//Кроме верхней группы
		if(!sizeof($gr['childs'])&&!sizeof($gr['data'])){
			array_splice($group,$i,1);
		}
	},array(),true);
	*/
}
function xls_merge(&$gr, &$addgr)
{
	//Всё из группы addgr нужно перенести в gr

	//echo $addgr['type'];
	//$gr['miss']=0;
	//if ($gr['pitch'] < $addgr['pitch'] && Xlsx::isParent($addgr, $gr)) {
		$gr['childs'] = array_merge($addgr['childs'], $gr['childs']);
	//} else {
	//	$gr['childs'] = array_merge($gr['childs'], $addgr['childs']);
	//}
	//Each::forr($addgr['childs'], function &(&$val) use (&$gr, $addgr) {
	//	$val['parent'] = &$gr;
		//Объединения с вложенной группой добавляется до своих подгрупп
		//Сначало собираем все подгруппы для добавление в текущую и разом добавляем
		/*$r = null;
		if ($gr['type'] == 'set' && $addgr['type'] == 'book') {
			return $r;
		} else if($gr['type'] == 'set' && $addgr['type'] == 'list') {
			return $r;
		} else if($gr['type'] == 'set' && $addgr['type'] == 'row') {
			return $r;
		} else {
			//if (in_array($addgr['type'],array('row','list'))) {
			$gr['childs'][] = &$val;
		}*/
	//	$r = null;
	//	return $r;
	//});

	Each::foro($addgr['descr'], function &($des, $key) use (&$gr) {
		//if ($key=='Наименование') return;
		if (!isset($gr['descr'][$key])) {
			$gr['descr'][$key] = $des;
		};
		$r = null;

		return $r;
	});

	if (!empty($addgr['tparam'])) {
		$tparam = $addgr['tparam'];
	} else {
		$tparam = '';
	}
	if (!empty($gr['tparam'])) {
		$gr['tparam'] .= ','.$tparam;
	} else {
		$gr['tparam'] = $tparam;
	}
	for ($i = 0, $l = sizeof($addgr['data']); $i < $l; $i++) {
		$pos = &$addgr['data'][$i];
		//$pos['parent'] = &$gr;
		
		$gr['data'][] = &$pos;
	}
	return;
}
function xls_processDescr(&$data)
{
	//
	Xlsx::runGroups($data, function &(&$gr) {
		$descr = array();
		Each::forr($gr['descr'], function &($row) use (&$descr) {
			$row = array_values($row);
			if (empty($row[1])) $row[1] = '';
			if (empty($row[0])) $row[0] = '';
			$descr[$row[0]] = $row[1];
			$r = null;
			return $r;
		});
		$gr['descr'] = &$descr;
		$r = null;
		return $r;
	});
}
function xls_processGroupCalculate(&$data)
{
	xls_runGroups($data, function &(&$data) {
		$data['count'] = sizeof($data['data']);
		$data['groups'] = 1;
		Each::forr($data['childs'], function &(&$d) use (&$data) {
			$data['count'] += $d['count'];
			$data['groups'] += $d['groups'];
			$r = null;

			return $r;
		});
		$r = null; return $r;
	}, true);
};

function xls_processClassEmpty(&$data, $clsname)
{
	Xlsx::runGroups($data, function (&$gr) use ($clsname) {
		$poss = array();
		for ($i = 0, $l = sizeof($gr['data']); $i < $l; ++$i) {
			if (!isset($gr['data'][$i][$clsname]) || !$gr['data'][$i][$clsname]) {
				continue;
			}
			$poss[] = $gr['data'][$i];
		}
		$gr['data'] = $poss;
		$r = null; return $r;
	});
}
function xls_processClass(&$data, $clsname, $musthave = false)
{

	$run = function (&$data, $run, $clsname, $musthave, $clsvalue = '') {
		if ($data['type'] == 'book' && $musthave) {
			$data['miss'] = true;
			$clsvalue = Path::encode($data['title']);
		} elseif ($data['type'] == 'list' && !empty($data['descr'][$clsname])) {
			//Если в descr указан класс то имя листа игнорируется иначе это будет группой каталога, а классом будет считаться имя книги
			$data['miss'] = true;//Если у листа есть позиции без группы он не расформировывается
			$clsvalue = Path::encode($data['descr'][$clsname]);
		} elseif ($data['type'] == 'row' && !empty($data['descr'][$clsname])) {
			$clsvalue = Path::encode($data['descr'][$clsname]);
		}
		foreach ($data['data'] as $i => $pos) {
			if (!isset($data['data'][$i][$clsname])) {
				$data['data'][$i][$clsname] = $clsvalue;//У позиции будет установлен ближайший класс
			} else {
				$data['data'][$i][$clsname] = Path::encode($data['data'][$i][$clsname]);
			}
			$r = null;
		};	
		Each::forr($data['childs'], function &(&$data) use ($run, $clsvalue, $clsname, $musthave) {
			$run($data, $run, $clsname, $musthave, $clsvalue);
			$r = null;

			return $r;
		});
	};
	$run($data, $run, $clsname, $musthave);
}
function xls_processGroupMiss(&$data)
{
	Xlsx::runGroups($data, function &(&$gr, $i, &$parent) {

		if (!empty($gr['miss']) && $parent) {
			//Берём детей missгруппы и переносим их в родительскую
			/*Each::forr($gr['childs'], function &(&$g) use (&$gr) {
				$g['parent'] = &$gr['parent'];
				$r = null;

				return $r;
			});*/
			array_splice($parent['childs'], $i, 1, $gr['childs']);

			Each::forr($gr['data'], function &(&$p) use (&$gr, &$parent) {
				//$p['parent'] = &$gr['parent'];
				$parent['data'][] = $p;
				$r = null;

				return $r;
			});

			//Each::forr($gr['childs'],function(&$gr,&$childs, &$d){
		//		array_splice($childs,($i++)-1,0,array(&$d));
		//		$d['parent']=&$gr['parent'];
		//	},array(&$gr,&$childs));
		//	$arr[]=&$gr;
		}
		$r = null; return $r;
	}, true);//Если бежим вперёд повторы несколько раз находим, так как добавляем в конец// Если бежим сзади рушится порядок
}
function _xls_sort($a, $b)
{
	return ($a < $b) ? -1 : ($a > $b) ? 1 : 0;
}
function _xls_sortName($a, $b)
{
	$a = $a['Наименование'];
	$b = $b['Наименование'];

	return ($a < $b) ? -1 : ($a > $b) ? 1 : 0;
}
function xls_pageList(&$poss, $page, $count, $sort, $numbers)
{
	$all = sizeof($poss);
	$pages = ceil($all / $count);
	if ($page > $pages) {
		$page = $pages;
	}
	if ($page < 1) {
		$page = 1;
	}
	if ($numbers < 1) {
		$numbers = 1;
	}
	--$numbers;
	//page pages numbers first last
	$first = floor($numbers / 2);
	$tfirst = $first;
	$last = $numbers - $first;
	$show = array();

	while ($tfirst) {
		$p = $page - $tfirst;
		if ($p < 1) {
			++$last;
			--$first;
		}
		--$tfirst;
	}
	while ($last) {
		$p = $page + $last;
		if ($p <= $pages) {
			$show[] = $p;
		} else {
			++$first;
		}
		--$last;
	}
	while ($first) {
		$p = $page - $first;
		if ($p > 0) {
			$show[] = $p;
		}
		--$first;
	}
	$show[] = (int) $page;
	//usort($show,'_xls_sort');
	sort($show);

	if ($sort == 'name') {
		usort($poss, '_xls_sortName');
	}
	Each::forr($poss, function &(&$p, $i) {
		$p['num'] = $i + 1;
		$r = null;

		return $r;
	});
	$next = $page + 1;
	$prev = $page - 1;
	if ($prev < 1) {
		$prev = 1;
	}
	if ($next > $pages) {
		$next = $pages;
	}
	$r = array(
		'next' => $next,
		'prev' => $prev,
		'show' => $show,//Список страниц
		'page' => $page,//Текущая страница
		'sort' => $sort,//сортировка
		'list' => array(),//Список позиций на выбранной странице
		'pages' => $pages,//Всего страниц
	);

	$start = ($page * $count - $count);
	for ($i = $start, $l = $start + $count; $i < $l; ++$i) {
		if (!$poss[$i]) {
			break;
		}
		$r['list'][] = &$poss[$i];
	}

	return $r;
}
function xls_preparePosFiles(&$pos, $pth, $props = array())
{
	if (empty($pos['images'])) {
		$pos['images'] = array();
	}
	if (empty($pos['texts'])) {
		$pos['texts'] = array();
	}
	if (empty($pos['files'])) {
		$pos['files'] = array();
	}
	$dir = array();
	if (Each::forr($props, function &($name) use (&$dir, &$pos) {
		$rname = Sequence::right($name);
		$val = Sequence::get($pos, $rname);
		if (!$val) {
			return true;
		}
		$dir[] = $val;
		$r = null;

		return $r;
	})) {
		return;
	}

	if ($dir) {
		$dir = implode('/', $dir).'/';
		$dir = $pth.$dir;
	} else {
		$dir = $pth;
	}

	$dir = Path::theme($dir);
	if (!$dir) {
		return false;
	}

	if (is_dir($dir)) {
		$paths = glob($dir.'*');
	} elseif (is_file($dir)) {
		$paths = array($dir);
		$p = Load::srcInfo($dir);
		$dir = $p['folder'];
	}

	Each::forr($paths, function &($p) use (&$pos, $dir) {

		$d = explode('/', $p);
		$name = array_pop($d);
		$n = mb_strtolower($name);
		$fd = Load::nameInfo($n);
		$ext = $fd['ext'];

		//if(!$ext)return;
		if (!is_file($dir.$name)) {
			return;
		}
		//$name=preg_replace('/\.\w{0,4}$/','',$name);

		/*$p=pathinfo($p);
		$name=$p['basename'];
		$ext=strtolower($p['extension']);*/
		
		if ($name{0} == '.') return;
		$dir=Path::pretty($dir);
		$name = Path::toutf($dir.$name);
		
		$im = array('png', 'gif', 'jpg');
		$te = array('html', 'tpl', 'mht', 'docx');
		if (Each::forr($im, function ($e) use ($ext) {
			if ($ext == $e) {
				return true;
			}
		})) {
			$pos['images'][] = $name;
		} elseif (Each::forr($te, function ($e) use ($ext) {
			if ($ext == $e) {
				return true;
			}
		})) {
			$pos['texts'][] = $name;
		} else {
			if ($ext != 'db') {
				$pos['files'][] = $name;
			}
		}
		$r = null;

		return $r;
	});
	$pos['images'] = array_unique($pos['images']);
	$pos['texts'] = array_unique($pos['texts']);
	$pos['files'] = array_unique($pos['files']);
}
/*
 * Нет рекурсии, нет подсчёта количества.. .Какие нужны колонки, что подготовить к вставки в адрес передаются свойством
 * По умолчанию
*/

function &xls_init($path, $config = array())
{
	//Возвращает полностью гототовый массив
	//if(Each::isAssoc($path)===true)return $path;//Это если переданы уже готовые данные вместо адреса до файла данных

	

	$ar = array();
	$isonefile = true;
	Each::exec($path, function &($path) use (&$isonefile, &$ar) {
		$p = Path::theme($path);

		if ($p && !is_dir($p)) {
			if ($isonefile === true) {
				$isonefile = $p;
			} else {
				$isonefile = false;
			}
			$ar[] = $path;
		} elseif ($p) {
			$isonefile = false;
			array_map(function ($file) use (&$ar, $p, $path) {
				if ($file{0}=='.') {
					return;
				}
				$fd = Load::nameInfo($file);
				if (in_array($fd['ext'], array('xls', 'xlsx'))) {
					$ar[] = $path.Path::toutf($file);
				}
			}, scandir($p));
		}
		$r=null; return $r;
	});

	if (empty($config['root'])) {
		if ($isonefile) {
			$d = Load::srcInfo($isonefile);
			$config['root'] = Path::toutf($d['name']);
		} else {
			$config['root'] = 'Каталог';
		}
	}
	$list = array();
	Each::forr($ar, function &(&$path) use (&$data, &$list) {
		$r = null;
		$d = &xls_make($path);
		$list[] = &$d;
		return $r;
	});
	
	return Xlsx::initData($list, $config);
};
class Xlsx
{
	/**
	 * Функция считывает листы из Excle книги
	 */
	public static $conf=array(
		'cache'=>'!xlsx/'
	);
	public static function merge($ar) {
		$parent = false;
		$data = _xls_createGroup($ar[0]['title'], $parent, 'set');
		unset($data['parent']);
		$data['miss'] = true;//Если в группе будет только одна подгруппа она удалится... подгруппа поднимится на уровень выше
		Each::forr($ar, function &(&$d) use (&$data) {
			$r = null;
			if (!$d) return $r;
			//$d['parent'] = &$data;
			$data['childs'][] = &$d;
			return $r;
		});
		xls_processGroupFilter($data);//Объединяются группы с одинаковым именем, Удаляются пустые группы

		xls_processGroupMiss($data);//Группы miss(производители) расформировываются

		Xlsx::prepareMetaGroup($data);
		return $data;
	}
	public static function &initData($ar, $config = array()) {
		if (empty($config['root'])) {
			$config['root'] = 'Каталог';
		}

		$parent = false;
		$data = _xls_createGroup($config['root'], $parent, 'set');//Сделали группу в которую объединяются все остальные
		$data['miss'] = true;//Если в группе будет только одна подгруппа она удалится... подгруппа поднимится на уровень выше



		Each::forr($ar, function &(&$d) use (&$data) {
			$r = null;
			if (!$d) return $r;
			//$d['parent'] = &$data;
			$data['childs'][] = &$d;
			return $r;
		});
	
		//Реверс записей на листе
		if (!isset($config['listreverse'])) $config['listreverse'] = false;
		if ($config['listreverse']) {
			foreach($data['childs'] as $book => $v) {
				foreach($data['childs'][$book]['childs'] as $list => $vv) {
					$data['childs'][$book]['childs'][$list]['data'] = array_reverse($data['childs'][$book]['childs'][$list]['data']);
				}
			}
		}

		xls_processDescr($data);
		
		if (!isset($config['Сохранить head'])) $config['Сохранить head'] = true;
		
		xls_processPoss($data, $config['Сохранить head']);
		
		if (!isset($config['Переименовать колонки'])) $config['Переименовать колонки'] = array();	
		if (!isset($config['Удалить колонки'])) $config['Удалить колонки'] = array();
		if (!isset($config['more'])) $config['more'] = false;
		
		
		xls_runPoss($data, function &(&$pos) use (&$config) {
			$r = null;
			foreach ($config['Удалить колонки'] as $k) {
				if (isset($pos[$k])) unset($pos[$k]);
			}
			foreach ($config['Переименовать колонки'] as $k => $v) {
				if (isset($pos[$k])) {
					$pos[$v] = $pos[$k];
					unset($pos[$k]);
				}
			}
			return $r;
		});

		if (!isset($config['Имя файла'])) $config['Имя файла'] = 'Производитель'; //Группа остаётся, а производитель попадает в описание каждой позиции
		

		if ($config['Имя файла'] == 'Производитель') {
			xls_processClass($data, 'Производитель', true);
		}//Должен быть обязательно miss раставляется

		xls_runPoss($data, function &(&$pos, $i, &$group) {
			// пустая позиция
			$r = null;
			if (sizeof($pos) == 2) { //group_title Производитель
				unset($group['data'][$i]);
				return $r;
			}
			return $r;
		});

		
		xls_processGroupFilter($data);//Объединяются группы с одинаковым именем, Удаляются пустые группы
		
		Xlsx::runGroups($data, function &(&$g) {
			unset($g['parent']);
			$r = null; return $r;
		});
		Xlsx::runPoss($data, function &(&$g) {
			unset($g['parent']);
			$r = null; return $r;
		});
		xls_processGroupMiss($data);//Группы miss(производители) расформировываются

	//xls_processGroupCalculate($data);//Добавляются свойства count groups сколько позиций и групп группы должны быть уже определены... почищены...				

		xls_runGroups($data, function &(&$gr, $i, &$parent) {
			
			if (empty($gr['tparam'])) {
				$gr['tparam'] = $parent['tparam'];
			}//tparam наследуется Оборудование:что-то, что-то

			if (!empty($gr['descr']['Производитель'])) {
				for ($i = 0, $il = sizeof($gr['data']); $i < $il; ++$i) {
					if (!empty($gr['data'][$i]['Производитель'])) {
						continue;
					}
					$gr['data'][$i]['Производитель'] = $gr['descr']['Производитель'];
					$gr['data'][$i]['producer'] = Path::encode($gr['descr']['Производитель']);
				}
			}
			$r = null; return $r;
		});

		if (empty($config['Подготовить для адреса'])) {
			$config['Подготовить для адреса'] = array('Артикул' => 'article','Производитель' => 'producer');
		}

		xls_processPossFS($data, $config['Подготовить для адреса']);//Заменяем левые символы в свойстве


		if (empty($config['Обязательные колонки'])) {
			$config['Обязательные колонки'] = array('article','producer');
		}
		
		xls_runGroups($data, function &(&$group) use ($config) {
			$r = null; 
			if (empty($group['data'])) return $r;
			$group['data'] = array_values($group['data']);
			for ($i = sizeof($group['data']); $i >= 0 ; $i--) {
				foreach ($config['Обязательные колонки'] as $propneed) {
					if (empty($group['data'][$i][$propneed])) {
						unset($group['data'][$i]);
						break;
					}
				}
			}
			$group['data'] = array_values($group['data']);
			return $r;
		});
		if (empty($config['Известные колонки'])) {
			$config['Известные колонки'] = array('Производитель','Наименование','Описание','Артикул','positions');
		}
		$config['Известные колонки'][] = 'parent';
		$config['Известные колонки'][] = 'more';
		foreach ($config['Подготовить для адреса'] as $k => $v) {
			$config['Известные колонки'][] = $v;
			$config['Известные колонки'][] = $k;
		}
		if (!empty($config['more'])) {
			xls_processPossMore($data, $config['Известные колонки']);
			//позициям + more
		}
		
		Xlsx::prepareMetaGroup($data);
		

		if (empty($config['Ссылка parent'])) {
			xls_runGroups($data, function &(&$group) {
				$r = null;
				unset($group['parent']);
				return $r;
			});
			xls_runPoss($data, function &(&$pos, $i) {
				$r = null;
				unset($pos['parent']);
				return $r;
			});
		}
		return $data;
	}
	public static function prepareMetaGroup(&$data) {
		Xlsx::runGroups($data, function &(&$gr, $i, &$parent) {
			//Имя листа или файла короткое и настоящие имя группы прячется в descr. но имя листа или файла также остаётся в title
			$r = null;
			if(!empty($gr['descr']['Наименование'])){
				$gr['name'] = $gr['descr']['Наименование'];//name крутое правильное Наименование группы
			}
			if (empty($gr['name'])) {
				$gr['name'] = $gr['title'];
			}//title то как называется файл или какое имя используется в адресной строке
			return $r;
		});
		Xlsx::runGroups($data, function &(&$group, $i, $parent) {
			if ($parent) {
				$group['group'] = $parent['title'];
			} else {
				$group['group'] = false;
			}
			if (!empty($group['descr']['Наименование'])) {
				$group['Группа'] = $group['descr']['Наименование'];
			} else {
				$group['Группа'] = $group['title'];
			}
			$r = null; return $r;
		});
		xls_runPoss($data, function &(&$pos, $i, $group) {
			$r = null;
			$pos['group'] = $group['title'];
			$pos['Группа'] = $group['Группа'];
			return $r;
		});
		
		xls_runGroups($data, function &(&$data, $i, &$group) {
			//path
			$r = null;
			if (!$group) {
				$data['path'] = array();
			} else {
				$data['path'] = $group['path'];
				$data['path'][] = $data['title'];
			}
			return $r;
		});
		xls_runPoss($data, function &(&$pos, $i, &$group) {
			$r = null;
			$pos['path'] = $group['path'];
			return $r;
		});
	}
	public static function isParent(&$layer, &$parent)
	{
		while ($layer) {
			if (Each::isEqual($parent, $layer)) {
				return true;
			}
			$layer = &$layer['parent'];
		}

		return false;
	}
	/**
	 * Можно передавать путь или данные - двухмерный массив для обработки после parseAll
	 * почти Xlxs::make
	 **/
	public static function &get($src, $title = false)
	{

		$data = xls_make($src, $title);
		
		
		xls_processDescr($data);
		
		xls_processPoss($data, true);

		Xlsx::runGroups($data, function &(&$data, $i, &$group) {
			//path
			$r = null;
			if (!$group) {
				$data['path'] = array();
			} else {
				$data['path'] = $group['path'];
				$data['path'][] = $data['title'];
			}
			return $r;
		});

		return $data;
	}
	public static function &runGroups(&$data, $callback, $back = false)
	{
		return xls_runGroups($data, $callback, $back);
	}
	public static function isSpecified($val = null){
		if (is_null($val) || $val==='') return false;
		return true;
	}
	public static function &runPoss(&$data, $callback, $back = false)
	{
		return xls_runPoss($data, $callback, $back);
	}
	/**
	 * Функция считывает листы из Excle книги или папки с Excel книгами.
	 * Применяется сложная логика объединения групп и формирования новых групп.
	 */
	public static function &init($src, $config = array())
	{
		return xls_init($src, $config);
	}
	/**
	 * Просто считать данные первого листа в файле без каких бы то обработок
	 * Втоым параметром можно передать конкретный лист
	 */
	public static function parse($src, $list = false)
	{
		return xls_parse($src);
	}
	public static function make($path, $title = false)
	{
		return xls_make($path, $title);
	}
	public static function &parseAll($path)
	{
		$data = Cache::exec('Разбор табличных данных', function &($path) {

			$file = Path::theme($path);

			
			$data = array();
			if (!$file) {
				return $data;
			}

			$in = Load::srcInfo($path);

			Cache::setTitle($path);
			if ($in['ext'] == 'xls') {
				require_once __DIR__.'/excel_parser/oleread.php';
				require_once __DIR__.'/excel_parser/reader.php';

				if (!$file) {
					return $data;
				}
				
				$d = new \Spreadsheet_Excel_Reader();
				$d->setOutputEncoding('utf-8');
				//$d->setUTFEncoder('mb');
				$d->read($file);

				Each::forr($d->boundsheets, function &($v, $k) use (&$d, &$data) {
					$data[$v['name']] = &$d->sheets[$k]['cells'];
					$r = null;

					return $r;
				});
			} elseif ($in['ext'] == 'csv') {
				$handle = fopen('php://memory', 'w+');
				fwrite($handle, Path::toutf(file_get_contents($file)));
				rewind($handle);
				$data = array(); //Массив будет хранить данные из csv
				while (($line = fgetcsv($handle, 0, ";")) !== false) { //Проходим весь csv-файл, и читаем построчно. 3-ий параметр разделитель поля
					$data[] = $line; //Записываем строчки в массив
				}
				fclose($handle);
				foreach($data as $k=>$v){
					foreach($data[$k] as $kk=>$vv){
						$vv=trim($vv);
						if($vv==='')unset($data[$k][$kk]);
						else $data[$k][$kk]=$vv;
					}
					if(!$data[$k])unset($data[$k]);
				}
				$data=array('list'=>$data);
			} elseif ($in['ext'] == 'xlsx') {
				$cacheFolder = Path::resolve(Xlsx::$conf['cache']);
				//$cacheFolder .= Path::encode($path).'/';//кэш			
				$cacheFolder .= md5($path).'/';//кэш			
				OldCache::fullrmdir($cacheFolder, true);//удалить старый кэш

				$r = mkdir($cacheFolder);
				if(!$r) {
					echo '<pre>';
					throw new \Exception('Не удалось создать папку для кэша '.$cacheFolder);
				}

				//разархивировать
				$zip = new \ZipArchive();
				$pathfs = Path::theme($path);

				
				if ((int) phpversion() > 6) {
					$zipcacheFolder = Path::tofs($cacheFolder);
					$archiveFile = Path::toutf($pathfs);
					if (!empty($_SERVER['WINDIR'])) { //Только для Виндовс
						$archiveFile = Path::toutf($archiveFile);
						//$cacheFolder = Path::toutf($cacheFolder);
					}
				} else {
					$zipcacheFolder = Path::tofs($cacheFolder); //Без кирилицы
					$archiveFile = Path::tofs($pathfs);
				}
				$r = $zip->open($archiveFile);
				if ($r===true) {
					$zip->extractTo($zipcacheFolder);
					$zip->close();

					$contents = simplexml_load_file($cacheFolder.'xl/sharedStrings.xml');

					$contents = $contents->si;

					$workbook = simplexml_load_file($cacheFolder.'xl/workbook.xml');
					$sheets = $workbook->sheets->sheet;

					$handle = opendir($cacheFolder.'xl/worksheets/');
					$i = 0;
					$syms = array();
					while ($file = readdir($handle)) {
						if ($file{0} == '.') {
							continue;
						}
						$src = $cacheFolder.'xl/worksheets/'.$file;
						if (!is_file($src)) {
							continue;
						}
						$files[] = $file;
					}
					closedir($handle);
					natsort($files);

					foreach ($files as $file) {
						$src = $cacheFolder.'xl/worksheets/'.$file;

						$list = $sheets[$i];
						++$i;
						$list = $list->attributes();
						$list = (string) $list['name'];

						$data[$list] = array();

						$sheet = simplexml_load_file($cacheFolder.'xl/worksheets/'.$file);
						$rows = $sheet->sheetData->row;
						foreach ($rows as $row) {
							$attr = $row->attributes();
							$r = (string) $attr['r'];
							$data[$list][$r] = array();
							$cells = $row->c;

							foreach ($cells as $cell) {
								if (!$cell->v) {
									continue;
								}

								$attr = $cell->attributes();
								if ($attr['t'] == 's') {
									$place = (integer) $cell->v;

									if (isset($contents[$place]->r)) {
										$value = '';
										foreach ($contents[$place]->r as $con) {
											$value .= $con->t;
										}
									} else {
										$value = $contents[$place]->t;
									}
								} else {
									$value = $cell->v;
									$value = (double) $value;
								}

								$attr = $cell->attributes();
								$c = (string) $attr['r'];//FA232
								preg_match("/\D+/", $c, $c);
								$c = $c[0];
								$syms[$c] = true;
								$data[$list][$r][$c] = (string) $value;
							}
						}
					}

					$syms = array_keys($syms);
					natsort($syms);
					/*usort($syms,function($a,$b){
						$la=strlen($a);
						$lb=strlen($b);
						if($la>$lb)return 1;
						if($la<$lb)return -1;
						if($a>$b)return 1;
						if($a<$b)return -1;
						return 0;
					});*/
					$symbols = array();
					foreach ($syms as $i => $s) {
						$symbols[$s] = $i + 1;
					}

					foreach ($data as $list => $listdata) {
						foreach ($listdata as $row => $rowdata) {
							$data[$list][$row] = array();
							foreach ($rowdata as $cell => $celldata) {
								$data[$list][$row][$symbols[$cell]] = $celldata;
							}
							if (!$data[$list][$row]) {
								unset($data[$list][$row]);
							}//Пустые строки нам не нужны
						}
					}
				}
				// Если что-то пошло не так, возвращаем пустую строку
				//return "";
				//собрать данные
			}
			return $data;
		}, array($path), ['akiyatkin\boo\Cache','getModifiedTime'], array($path));
		
		return $data;
	}
	public static function getFiles($src) {
		return Once::func( function ($src){
			$res = [
				'images' => array(),
				'texts' => array(),
				'files' => array()
			];
			$dir = Path::theme($src);
			if (!$dir) return $res;
			
			if (is_dir($dir)) {
				$paths = scandir($dir);
				foreach($paths as $k=>$v) {
					$paths[$k] = $dir.$v;
				}
			} else {
				$paths = array($dir);
				$p = Load::srcInfo($src);
				$src = $p['folder'];
			}
			
			Each::forr($paths, function &($p) use (&$res, $src) {

				$d = explode('/', $p);
				$name = array_pop($d);
				$n = mb_strtolower($name);
				$fd = Load::nameInfo($n);
				$ext = $fd['ext'];
				$r = null;
				if (!Path::theme($src.$name)) return $r;
				if ($name{0} == '.') return $r;
				$path = $src.Path::toutf($name);
				
				
				$im = array('png', 'gif', 'jpg');
				$te = array('html', 'tpl', 'mht', 'docx');
				$ignore = array('db', 'json');

				if (in_array($ext, $im)) {
					$res['images'][] = $path;
				} else if (in_array($ext, $te)) {
					$res['texts'][] = $path;
				} else {
					if (!in_array($ext, $ignore)) {
						$res['files'][] = Load::srcInfo($path);
					}
				}
				return $r;
			});
			return $res;
		}, array($src), ['akiyatkin\\boo\\Cache','getModifiedTime'], [$src]);
	}
	public static function addFiles($root, &$pos, $dir = false)
	{
		
		if (!$dir) {
			$props = array('producer','article');
			$dir = array();
			$pth = Path::resolve($root);
			Each::forr($props, function &($name) use (&$dir, $pos) {
				$rname = Sequence::right($name);
				$val = Sequence::get($pos, $rname);
				$dir[] = $val;

				$r = null;
				return $r;
			});

			if ($dir) {
				$dir = implode('/', $dir).'/';
				$dir = $root.$dir;
			} else {
				$dir = $root;
			}

		} else {
			$dir = $root.$dir;
		}

		$res = Xlsx::getFiles($dir);

		if (!isset($pos['images'])) {
            $pos['images'] = array();
        }
        if (!isset($pos['texts'])) {
        	$pos['texts'] = array();
        }
        if (!isset($pos['files'])) {
        	$pos['files'] = array();
        }
        
		$pos['images'] = array_merge($res['images'], $pos['images']);
		$pos['files'] = array_merge($res['files'], $pos['files']);
		$pos['texts'] = array_merge($res['texts'], $pos['texts']);
		
		//$pos['images'] = array_unique($pos['images']);
		//$pos['texts'] = array_unique($pos['texts']);
		//$pos['files'] = array_unique($pos['files']);

	}
}
