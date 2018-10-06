<?php

namespace gri3li\yii2gridfile\tests;

use Yii;
use yii\helpers\FileHelper;
use PHPUnit\Framework\TestCase as BaseTestCase;
use yii\console\Application;

/**
 * Base class for the test cases.
 */
class TestCase extends BaseTestCase
{
    protected $data;
    /**
     * @throws \yii\base\Exception
     * @throws \yii\base\InvalidConfigException
     */
    protected function setUp()
    {
        parent::setUp();
        new Application([
            'id' => 'testApp',
            'basePath' => __DIR__,
            'vendorPath' => dirname(__DIR__) . '/vendor',
            'components' => [
                'db' => [
                    'class' => 'yii\db\Connection',
                    'dsn' => 'sqlite::memory:',
                ],
            ],
        ]);
        $this->data = Yii::getAlias('@runtime/test');
        FileHelper::createDirectory($this->data);
    }

    protected function tearDown()
    {
        parent::tearDown();
        FileHelper::removeDirectory($this->data);
        Yii::$app = null;
    }

    /**
     * Invokes object method, even if it is private or protected.
     * @param object $object object.
     * @param string $method method name.
     * @param array $args method arguments
     * @return mixed method result
     * @throws \ReflectionException
     */
    protected function invoke($object, $method, array $args = [])
    {
        $classReflection = new \ReflectionClass(get_class($object));
        $methodReflection = $classReflection->getMethod($method);
        $methodReflection->setAccessible(true);
        $result = $methodReflection->invokeArgs($object, $args);
        $methodReflection->setAccessible(false);
        return $result;
    }
}
