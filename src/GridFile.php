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
 *              'startColor' => ['argb' => 'FFA0A0A0'],
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
 * @property array|Formatter $formatter the formatter used to format model attribute values into displayable texts.
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
    public $dataColumnClass = DataColumn::class;

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
     * @var null dummy for \yii\grid\DataColumn (he use the grid)
     */
    public $filterModel = null;

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
            'class' => $this->dataColumnClass,
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
                    'class' => $this->dataColumnClass,
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
        $headerPosition = 1;
        $i = 1;
        /** @var \yii\grid\Column $col */
        foreach ($this->columns as $col) {
            $value = strip_tags($col->renderHeaderCell());
            $cell = $this->spreadsheet->getActiveSheet()->getCellByColumnAndRow($i, $headerPosition);
            $cell->setValue($value);
            $cell->getStyle()->applyFromArray(array_merge($this->cellStyle, $this->headerCellStyle));
            $i++;
        }
    }

    /**
     * Renders the table body
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function renderTableBody()
    {
        $headerOffset = 1;
        $rowIndex = 1;
        foreach ($this->dataProvider->getModels() as $model) {
            $columnIndex = 1;
            /** @var \yii\grid\Column $col */
            foreach ($this->columns as $col) {
                $value = strip_tags($col->renderDataCell($model, null, $rowIndex));
                $cell = $this->spreadsheet->getActiveSheet()->getCellByColumnAndRow($columnIndex, $rowIndex + $headerOffset);
                $cell->setValue($value);
                $cell->getStyle()->applyFromArray(array_merge($this->cellStyle, $this->bodyCellStyle));
                $columnIndex++;
            }
            $rowIndex++;
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
        $this->renderTableHeader();
        $this->renderTableBody();
        /** @var \PhpOffice\PhpSpreadsheet\Writer\IWriter $writer */
        $writer = new $writerClass($this->spreadsheet);
        $writer->save($path);
    }
}
