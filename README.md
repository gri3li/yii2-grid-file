
Data Export extension for Yii2 based on PhpSpreadsheet
===

This extension provides ability to export data form data provider to format supported by PhpSpreadsheet

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
php composer require --prefer-dist gri3li/grid-file
```

or add

```json
"yii2tech/csv-grid": "*"
```

to the require section of your composer.json.


Usage
-----

 ```php
$export = new GridFile([
    'dataProvider' => new \yii\data\ArrayDataProvider([
        'allModels' => [
            [
                'name' => 'some name',
                'date' => 1538571363,
            ],
            [
                'name' => 'name 2',
                'date' => 1538571363,
            ],
        ],
    ]),
    'columns' => [
        'name',
        'date:datetime',
    ],
]);
$export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Xls::class, '/path/to/file.xls');

// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Xlsx::class, '/path/to/file.xlsx');
// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Ods::class, '/path/to/file.ods');
// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Html::class, '/path/to/file.html');
// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Csv::class, '/path/to/file.csv');
 ```