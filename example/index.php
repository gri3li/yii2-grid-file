<?php

error_reporting(E_ALL);
ini_set('display_errors', 1);
define('YII_ENABLE_ERROR_HANDLER', false);
define('YII_DEBUG', true);

require __DIR__ . '/../../../autoload.php';
require __DIR__ . '/../../../yiisoft/yii2/Yii.php';
require __DIR__ . '/controllers/SiteController.php';

(new yii\web\Application([
    'id' => 'exampleApp',
    'basePath' => __DIR__,
    'vendorPath' => __DIR__ . '/../../../',
    'controllerNamespace' => 'app\\example\\controllers',
]))->run();
