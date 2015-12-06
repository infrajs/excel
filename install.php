<?php
namespace infrajs\autoedit;
use infrajs\path\Path;
require_once(__DIR__.'/../../../vendor/autoload.php');
require_once(__DIR__.'/../path/install.php');

Path::mkdir($dirs['cache'].'.xlsx/');