<?php

namespace gri3li\yii2gridfile;

use Yii;
use yii\di\Instance;
use yii\grid\DataColumn;
use yii\i18n\Formatter;
use yii\base\Component;
use yii\base\InvalidConfigException;
use yii\data\DataProviderInterface;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

/**
 * Data Export extension based on PhpSpreadsheet
 *
 * Example:
 *
 * ```php
 * $export = new \gri3li\yii2gridfile\GridFile([
 *     'dataProvider' => new \yii\data\ArrayDataProvider([
 *         'allModels' => [
 *             [
 *                 'name' => 'some name',
 *                 'date' => 1538571363,
 *             ],
 *             [
 *                 'name' => 'name 2',
 *                 'date' => 1538571363,
 *             ],
 *         ],
 *     ]),
 *     'columns' => [
 *         'name',
 *         'date:datetime',
 *     ],
 *     'headerCellStyle' => [
 *         'font' => ['bold' => true],
 *         'fill' => [
 *              'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
 *              'startColor' => ['rgb' => 'CCCCCC'],
 *          ],
 *     ],
 * ]);
 * $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Xls::class, '/path/to/file.xls');
 *
 * // $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Xlsx::class, '/path/to/file.xlsx');
 * // $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Ods::class, '/path/to/file.ods');
 * // $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Html::class, '/path/to/file.html');
 * // $export->saveAs(\PhpOffice\PhpSpreadsheet\Writer\Csv::class, '/path/to/file.csv');
 * ```
 *
 * @author Mikhail Gerasimov <migerasimoff@gmail.com>
 * @since 1.0
 */
class GridFile extends Component
{
    /**
     * @var \yii\data\DataProviderInterface
     */
    public $dataProvider;

    /**
     * @var string the default data column class if the class name is not explicitly specified when configuring a data column.
     */
    public $dataColumnClass;

    /**
     * @var array|\yii\grid\Column[] grid column configuration
     */
    public $columns = [];

    /**
     * @var array PhpSpreadsheet style in array format for all cells
     */
    public $cellStyle = [];

    /**
     * @var array PhpSpreadsheet style in array format for header cells
     */
    public $headerCellStyle = [];

    /**
     * @var array PhpSpreadsheet style in array format for body cells
     */
    public $bodyCellStyle = [];

    /**
     * @var array PhpSpreadsheet style in array format for body cells
     */
    public $footerCellStyle = [];

    /**
     * @var null dummy for \yii\grid\DataColumn (he use the grid)
     */
    public $filterModel = null;

    /**
     * @var string
     */
    public $emptyCell = '';

    /**
     * @var boolean whether to show the header section of the grid table.
     */
    public $showHeader = true;

    /**
     * @var boolean whether to show the footer section of the grid table.
     */
    public $showFooter = false;

    /**
     * @var array|Formatter the formatter used to format model attribute values into displayable texts.
     * This can be either an instance of [[Formatter]] or an configuration array for creating the [[Formatter]]
     * instance. If this property is not set, the "formatter" application component will be used.
     */
    public $formatter;

    /**
     * @var Spreadsheet
     */
    public $spreadsheet;

    /**
     * @var string grid first cell position.
     */
    public $startTopLeftCell = 'A:1';

    /**
     * @var int
     */
    private $startColumnIndex;

    /**
     * @var int
     */
    private $startRowIndex;

    /**
     * @throws InvalidConfigException
     */
    public function init()
    {
        parent::init();
        if (!($this->dataProvider instanceof DataProviderInterface)) {
            throw new InvalidConfigException('dataProvider should implement the DataProviderInterface');
        }
        if ($this->formatter === null) {
            $this->formatter = Yii::$app->getFormatter();
        } else {
            $this->formatter = Instance::ensure($this->formatter, Formatter::class);
        }
        if ($this->spreadsheet === null) {
            $this->spreadsheet = new Spreadsheet();
        } elseif (!($this->spreadsheet instanceof Spreadsheet)) {
            throw new InvalidConfigException('spreadsheet class must inherit from PhpOffice\PhpSpreadsheet\Spreadsheet');
        }
        $this->startColumnIndex = Coordinate::columnIndexFromString(explode(':', $this->startTopLeftCell)[0]);
        $this->startRowIndex = explode(':', $this->startTopLeftCell)[1];
    }

    /**
     * This function tries to guess the columns to show from the given data
     * if [[columns]] are not explicitly specified.
     */
    protected function guessColumns()
    {
        $models = $this->dataProvider->getModels();
        $model = reset($models);
        if (is_array($model) || is_object($model)) {
            foreach ($model as $name => $value) {
                if ($value === null || is_scalar($value) || is_callable([$value, '__toString'])) {
                    $this->columns[] = (string) $name;
                }
            }
        }
    }

    /**
     * Creates a [[DataColumn]] object based on a string in the format of "attribute:format:label".
     *
     * @param string $text the column specification string
     * @return DataColumn the column instance
     * @throws InvalidConfigException if the column specification is invalid
     */
    protected function createDataColumn($text)
    {
        if (!preg_match('/^([^:]+)(:(\w*))?(:(.*))?$/', $text, $matches)) {
            throw new InvalidConfigException('The column must be specified in the format of "attribute", "attribute:format" or "attribute:format:label"');
        }

        return Yii::createObject([
            'class' => $this->dataColumnClass ? : DataColumn::class,
            'grid' => $this,
            'attribute' => $matches[1],
            'format' => isset($matches[3]) ? $matches[3] : 'text',
            'label' => isset($matches[5]) ? $matches[5] : null,
        ]);
    }

    /**
     * Creates column objects and initializes them
     *
     * @throws InvalidConfigException
     */
    protected function initColumns()
    {
        if (empty($this->columns)) {
            $this->guessColumns();
        }
        foreach ($this->columns as $i => $column) {
            if (is_string($column)) {
                $column = $this->createDataColumn($column);
            } else {
                $column = Yii::createObject(array_merge([
                    'class' => $this->dataColumnClass ? : DataColumn::class,
                    'grid' => $this,
                ], $column));
            }
            if (!$column->visible) {
                unset($this->columns[$i]);
                continue;
            }
            $this->columns[$i] = $column;
        }
    }

    /**
     * Renders the table header
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function renderTableHeader()
    {
        $columnIndex = $this->startColumnIndex;
        /** @var \yii\grid\Column $column */
        foreach ($this->columns as $column) {
            $cell = $this->spreadsheet->getActiveSheet()->getCellByColumnAndRow($columnIndex++, $this->startRowIndex);
            $value = $column->renderHeaderCell();
            $cell->setValue(html_entity_decode(strip_tags($value)));
            $cell->getStyle()->applyFromArray(array_merge($this->cellStyle, $this->headerCellStyle));
        }
        $this->startRowIndex++;
    }

    /**
     * Renders the table body
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function renderTableBody()
    {
        foreach ($this->dataProvider->getModels() as $model) {
            $columnIndex = $this->startColumnIndex;
            /** @var \yii\grid\Column $column */
            foreach ($this->columns as $column) {
                $cell = $this->spreadsheet->getActiveSheet()->getCellByColumnAndRow($columnIndex, $this->startRowIndex);
                $value = $column->renderDataCell($model, null, $columnIndex - $this->startColumnIndex);
                $cell->setValue(html_entity_decode(strip_tags($value)));
                $cell->getStyle()->applyFromArray(array_merge($this->cellStyle, $this->bodyCellStyle));
                $columnIndex++;
            }
            $this->startRowIndex++;
        }
    }

    /**
     * Renders the table footer
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function renderTableFooter()
    {
        $columnIndex = $this->startColumnIndex;
        /* @var \yii\grid\Column $column */
        foreach ($this->columns as $column) {
            $cell = $this->spreadsheet->getActiveSheet()->getCellByColumnAndRow($columnIndex++, $this->startRowIndex);
            $value = $column->renderFooterCell();
            $cell->setValue(html_entity_decode(strip_tags($value)));
            $cell->getStyle()->applyFromArray(array_merge($this->cellStyle, $this->footerCellStyle));
        }
    }

    /**
     * Make export file
     *
     * @param string $writerClass
     * @param string $path
     * @throws InvalidConfigException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function saveAs($writerClass, $path)
    {
        if (!in_array(IWriter::class, class_implements($writerClass))) {
            throw new InvalidConfigException('writerClass should implement the IWriter interface');
        }
        $this->initColumns();
        $this->autoSizeColumns();
        if ($this->showHeader) {
            $this->renderTableHeader();
        }
        $this->renderTableBody();
        if ($this->showFooter) {
            $this->renderTableFooter();
        }
        /** @var \PhpOffice\PhpSpreadsheet\Writer\IWriter $writer */
        $writer = new $writerClass($this->spreadsheet);
        $writer->save($path);
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function autoSizeColumns()
    {
        for ($i = $this->startColumnIndex, $end = $this->startColumnIndex + count($this->columns); $i < $end; $i++) {
            $this->spreadsheet->getActiveSheet()->getColumnDimensionByColumn($i)->setAutoSize(true);
        }
    }
}
