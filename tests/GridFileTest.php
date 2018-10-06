<?php

namespace gri3li\yii2gridfile\tests;

use Yii;
use gri3li\yii2gridfile\GridFile;
use yii\base\InvalidConfigException;
use yii\data\ActiveDataProvider;
use yii\data\ArrayDataProvider;
use yii\data\SqlDataProvider;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Ods;
use PhpOffice\PhpSpreadsheet\Writer\Html;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use yii\db\Query;

class GridFileTest extends TestCase
{
    public function testRequired()
    {
        $e = null;
        try {
            new GridFile();
        } catch (InvalidConfigException $e) {}
        $this->assertInstanceOf(InvalidConfigException::class, $e);
    }

    public function testFormats()
    {
        $export = new GridFile([
            'dataProvider' => new ArrayDataProvider([
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
        $formats = [
            ['file' => $this->data . '/tmp.xls', 'writerClass' => Xls::class],
            ['file' => $this->data . '/tmp.xlsx', 'writerClass' => Xlsx::class],
            ['file' => $this->data . '/tmp.ods', 'writerClass' => Ods::class],
            ['file' => $this->data . '/tmp.html', 'writerClass' => Html::class],
            ['file' => $this->data . '/tmp.csv', 'writerClass' => Csv::class],
        ];
        foreach ($formats as $format) {
            $export->saveAs($format['writerClass'], $format['file']);
            $this->assertTrue(file_exists($format['file']));
        }
    }

    public function testDataProviders()
    {
        $this->setupTestDbData();
        $check = 'check';
        $providers = [];
        $providers[] = new ArrayDataProvider([
            'allModels' => [
                ['id' => 1, 'name' => $check],
            ],
        ]);
        $providers[] = new SqlDataProvider([
            'sql' => 'SELECT * FROM item',
        ]);
        $providers[] = new ActiveDataProvider([
            'query' => (new Query())->from('item'),
        ]);
        foreach ($providers as $provider) {
            $export = new GridFile(['dataProvider' => $provider]);
            $models = $export->dataProvider->getModels();
            $this->assertEquals($check, $models[0]['name']);
        }
    }

    protected function setupTestDbData()
    {
        Yii::$app->db->createCommand()->createTable('item', [
            'id' => 'pk',
            'name' => 'string',
        ])->execute();
        Yii::$app->db->createCommand()->batchInsert('item', ['name'], [
            ['check'],
        ])->execute();
    }
}
