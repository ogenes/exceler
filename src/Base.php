<?php
/**
 * User: ogenes<ogenes.yi@gmail.com>
 * Date: 2022/6/17
 */

namespace Ogenes\Exceler;


abstract class Base
{
    /**
     * @var array
     */
    protected static $instance = [];
    
    /**
     * @return static
     */
    public static function getInstance(): Base
    {
        $className = static::class;
        if (isset(self::$instance[$className])) {
            return self::$instance[$className];
        }
        self::$instance[$className] = new static();
        return self::$instance[$className];
    }
}