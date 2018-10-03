
Data Export extension for Yii2 based on PhpSpreadsheet
===

This Yii2 extension provides ability to export data form data provider to format supported by PhpSpreadsheet

Installation
------------

The preferred way to install this extension is through [composer](http://getcomposer.org/download/).

Either run

```
composer require --prefer-dist gri3li/yii2-grid-file
```

or add

```json
"gri3li/yii2-grid-file": "*"
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
    'headerCellStyle' => [
        'font' => ['bold' => true],
        'fill' => [
            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
            'startColor' => ['argb' => 'FFA0A0A0'],
        ],
    ],
]);
$export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Xls::class, '/path/to/file.xls');

// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Xlsx::class, '/path/to/file.xlsx');
// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Ods::class, '/path/to/file.ods');
// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Html::class, '/path/to/file.html');
// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Csv::class, '/path/to/file.csv');
 ```
 
 More info about phpspreadsheet style [https://phpspreadsheet.readthedocs.io/en/develop/topics/recipes/#styles](https://phpspreadsheet.readthedocs.io/en/develop/topics/recipes/#styles)