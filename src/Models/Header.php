<?php
/**
 * Created by PhpStorm.
 * User: hugh.li
 * Date: 2024/4/5
 * Time: 23:26
 */

namespace HughCube\Spreadsheet\Models;

class Header
{
    protected $index;

    protected $title;

    protected $format;

    public function __construct($index, $title, $format = null)
    {
        $this->index = $index;
        $this->title = $title;
        $this->format = $format;
    }

    public function getIndex()
    {
        return $this->index;
    }

    public function formatValue($value)
    {
        return null === $this->format ? $value : call_user_func($this->format, $value);
    }
}
