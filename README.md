Data Export extension for Yii2
==============================

This Yii2 extension provides ability to export data form instances of `yii\data\DataProviderInterface` to format supported by PhpSpreadsheet

Installation
------------

Install the package via Composer:

```bash
composer require gri3li/yii2-grid-file
```

Usage
-----

 ```php
$export = new \gri3li\yii2gridfile\GridFile([
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
            'startColor' => ['rgb' => 'CCCCCC'],
        ],
    ],
]);
$export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Xls::class, '/path/to/file.xls');

// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Xlsx::class, '/path/to/file.xlsx');
// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Ods::class, '/path/to/file.ods');
// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Html::class, '/path/to/file.html');
// $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Csv::class, '/path/to/file.csv');
 ```

More info about phpspreadsheet style [https://phpspreadsheet.readthedocs.io/en/latest/topics/recipes/#styles](https://phpspreadsheet.readthedocs.io/en/latest/topics/recipes/#styles)

Use case [https://github.com/gri3li/yii2-grid-file/tree/master/example](https://github.com/gri3li/yii2-grid-file/tree/master/example)

For run use case:

```
cd vendor/gri3li/yii2-grid-file/example/
php -S 127.0.0.1:8877
```

open [http://127.0.0.1:8877](http://127.0.0.1:8877)