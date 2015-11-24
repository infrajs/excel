<?php

$dirs = infra_dirs();
if (!is_dir($dirs['cache'].'xlsx/')) {
	mkdir($dirs['cache'].'xlsx/');
}