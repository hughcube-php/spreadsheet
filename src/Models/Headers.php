<?php
/**
 * Created by PhpStorm.
 * User: hugh.li
 * Date: 2024/4/5
 * Time: 23:26
 */

namespace HughCube\Spreadsheet\Models;

use HughCube\Spreadsheet\ParseSheet;
use Throwable;

class Headers
{
    /**
     * @var int
     */
    protected $index;

    /**
     * @var array
     */
    protected $headers;

    public function __construct($index, array $headers)
    {
        $this->index = $index;
        $this->headers = $headers;
    }

    public function getHeaders(): array
    {
        return $this->headers;
    }

    public function getIndex(): int
    {
        return $this->index;
    }

    public function getMaxHeaderIndex(): string
    {
        $index = 'A';
        foreach ($this->getHeaders() as $header) {
            $index = max($index, $header->getIndex());
        }
        return $index;
    }

    /**
     * @return null|static
     */
    public static function parse(ParseSheet $parse, $patterns)
    {
        foreach ($parse->getsheet()->getRowIterator() as $row) {

            /** 获取整个行 */
            $cells = [];
            foreach ($row->getCellIterator() as $index => $cell) {
                try {
                    $cells[$index] = $cell->getFormattedValue();
                } catch (Throwable $exception) {
                    $cells[$index] = $cell->getValue();
                }
            }

            /** 尝试解析Header */
            $headers = [];
            foreach ($cells as $index => $cell) {
                foreach ($patterns as $key => $pattern) {
                    if (call_user_func($pattern['is'], $cell)) {
                        $headers[$key] = new Header($index, $cell, $pattern['format'] ?? null);
                        break;
                    }
                }
            }

            if (count($headers) === count($patterns)) {
                /** @phpstan-ignore-next-line */
                return new static($row->getRowIndex(), $headers);
            }
        }

        return null;
    }
}
