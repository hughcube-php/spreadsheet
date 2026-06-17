<?php

namespace HughCube\Spreadsheet\Tests;

use HughCube\Spreadsheet\SpreadsheetReader;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use ZipArchive;

class SpreadsheetReaderTest extends TestCase
{
    /** @var string */
    private $file = '';

    protected function tearDown(): void
    {
        if ('' !== $this->file && file_exists($this->file)) {
            @unlink($this->file);
        }
        parent::tearDown();
    }

    /**
     * 真实数据只占 A:C, 但 sheet1.xml 被写入延伸到 ZZ 列的空单元格(模拟 Excel
     * 产出的膨胀表): 默认加载会把空单元格读进来(最高列被撑到 ZZ), 安全加载则跳过
     * (最高列回到真实数据的 C), 数据保持完好。
     */
    public function testLoadSkipsTrailingEmptyCells(): void
    {
        $this->file = $this->makeBloatedFile();

        $plain = IOFactory::load($this->file);
        self::assertSame('ZZ', $plain->getActiveSheet()->getHighestColumn());
        $plain->disconnectWorksheets();

        $safe = SpreadsheetReader::load($this->file);
        self::assertSame('C', $safe->getActiveSheet()->getHighestColumn());
        self::assertSame('张三', $safe->getActiveSheet()->getCell('A2')->getValue());
        $safe->disconnectWorksheets();
    }

    /**
     * 造一个 A:C 有数据、第一行末尾被注入一个延伸到 ZZ 列空单元格的 xlsx,
     * 模拟 Excel/外部工具产出的「使用范围被撑大」的表格。
     */
    private function makeBloatedFile(): string
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->getActiveSheet()->fromArray(
            [['姓名', '年龄', '城市'], ['张三', 18, '上海']],
            null,
            'A1'
        );

        $file = sys_get_temp_dir() . '/sr_' . uniqid() . '.xlsx';
        (new Xlsx($spreadsheet))->save($file);
        $spreadsheet->disconnectWorksheets();

        $zip = new ZipArchive();
        $zip->open($file);
        $xml = $zip->getFromName('xl/worksheets/sheet1.xml');
        $xml = preg_replace('#</row>#', '<c r="ZZ1"/></row>', $xml, 1);
        $zip->addFromString('xl/worksheets/sheet1.xml', $xml);
        $zip->close();

        return $file;
    }
}
