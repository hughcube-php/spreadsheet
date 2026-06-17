<?php

namespace HughCube\Spreadsheet;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * 表格安全加载器
 *
 * 统一替代裸 IOFactory::load(): 读取时跳过空单元格, 规避「真实数据只有几十行,
 * 却被 Excel 等工具写入延伸到 XFD 列(第 16384 列)的上百万个空占位单元格」的表格 ——
 * 这类表格会让 PhpSpreadsheet 为每个空格都建 Cell 对象而内存耗尽(实测一份 66 行的
 * 表, 默认加载峰值 490MB/68s, 跳过空单元格后 32MB/6.5s)。
 *
 * 为兼容 1.x/2.x/3.x 三条版本线(最低 phpspreadsheet ^1.0、PHP 7.1):
 *  - 用 identify()+createReader() 而非 createReaderForFile()(后者自 phpspreadsheet 1.5 才有);
 *  - setReadEmptyCells(bool) 自 1.0 起即存在, 各版本(有/无类型声明)均兼容;
 *  - 仅用 PHP 7.0+ 语法。
 */
class SpreadsheetReader
{
    /**
     * 安全加载表格文件.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function load(string $file): Spreadsheet
    {
        $reader = IOFactory::createReader(IOFactory::identify($file));
        $reader->setReadEmptyCells(false);

        return $reader->load($file);
    }
}
