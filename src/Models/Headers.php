<?php
/**
 * Created by PhpStorm.
 * User: hugh.li
 * Date: 2024/4/5
 * Time: 23:26
 */

namespace HughCube\Spreadsheet\Models;

use HughCube\Spreadsheet\SheetParser;
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
     * @return array<int, null|static>
     */
    public static function parse(SheetParser $parse, $patterns): array
    {
        $requiredKeys = [];
        foreach ($patterns as $key => $pattern) {
            if ($pattern['required'] ?? true) {
                $requiredKeys[] = $key;
            }
        }

        $match = false;
        $closestHeaders = null;
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

            if (!$closestHeaders instanceof static || count($headers) >= count($closestHeaders->getHeaders())) {
                /** @phpstan-ignore-next-line */
                $closestHeaders = new static($row->getRowIndex(), $headers);
            }

            if (empty(array_diff($requiredKeys, array_keys($headers)))) {
                $match = true;
                break;
            }
        }

        return [($match ? $closestHeaders : null), $closestHeaders];
    }
}
