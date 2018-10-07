<?php
/* @var $this \yii\web\View */
?>
<?php $this->beginPage() ?>
<!DOCTYPE html>
<html lang="<?= Yii::$app->language ?>">
<head>
    <meta charset="<?= Yii::$app->charset ?>">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>example</title>
    <?php $this->head() ?>
</head>
<body>
<?php $this->beginBody() ?>

<div class="wrap">
    <div class="container">
        <?= \yii\helpers\Html::a('Экспорт', \yii\helpers\ArrayHelper::merge(['export'], Yii::$app->request->get())) ?>
        <br>
        <br>
        <?= \yii\grid\GridView::widget([
            'filterModel' => $searchModel,
            'dataProvider' => $dataProvider,
            'summary' => false,
            'columns' => [
                'name',
                'date:datetime'
            ],
        ]) ?>
    </div>
</div>

<?php $this->endBody() ?>
</body>
</html>
<?php $this->endPage() ?>
