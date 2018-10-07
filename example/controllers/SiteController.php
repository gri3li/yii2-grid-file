<?php

namespace app\example\controllers;

use Yii;
use yii\web\Controller;
use yii\base\DynamicModel;
use yii\data\ArrayDataProvider;
use gri3li\yii2gridfile\GridFile;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\IOFactory;

class SiteController extends Controller
{
    /**
     * example stub
     * @var array
     */
    private $models = [
        [
            'name' => 'one',
            'date' => 1538571363,
        ],
        [
            'name' => 'two',
            'date' => 1538571363,
        ],
    ];

    /**
     * @var \yii\base\Model
     */
    private $searchModel;

    /**
     * @var \yii\data\BaseDataProvider
     */
    private $dataProvider;


    public function init()
    {
        $this->searchModel = new DynamicModel(['name']);
        $this->searchModel->addRule(['name'], 'safe');
        $this->searchModel->load(Yii::$app->request->get());
        $models = array_filter($this->models, function ($model) {
            if (!empty($this->searchModel->name)) {
                return stripos($model['name'], $this->searchModel->name) !== false;
            }
            return true;
        });
        $this->dataProvider = new ArrayDataProvider(['allModels' => $models]);
    }


    public function actionIndex()
    {
        return $this->renderPartial('index', [
            'searchModel' => $this->searchModel,
            'dataProvider' => $this->dataProvider,
        ]);
    }


    public function actionExport()
    {
        // Spreadsheet can be created in advance
        $template = __DIR__ . '/../export-template/export.xls';
        $spreadsheet = IOFactory::load($template);

        // no pagination !!!
        $this->dataProvider->setPagination(false);

        $tmpFile = tempnam(sys_get_temp_dir(), 'contract-xls-tmp-');

        $export = new GridFile([
            'dataProvider' => $this->dataProvider,
            'spreadsheet' => $spreadsheet,
            'startTopLeftCell' => 'A:3',
            'columns' => [
                'name',
                'date:datetime',
            ],
            'headerCellStyle' => [
                'font' => ['bold' => true],
                'fill' => [
                    'fillType' => Fill::FILL_SOLID,
                    'startColor' => ['rgb' => 'CCCCCC'],
                ],
            ],
        ]);

        $export->saveAs(Xls::class, $tmpFile);
        Yii::$app->response->sendFile($tmpFile, 'export.xls');
        unlink($tmpFile);
    }
}
