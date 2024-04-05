<?php
/**
 * Created by PhpStorm.
 * User: hugh.li
 * Date: 2024/4/6
 * Time: 01:22
 */

namespace HughCube\Spreadsheet\Tests;

use HughCube\Spreadsheet\ParseSheet;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ParseSheetTest extends TestCase
{
    ///**
    // * @throws Exception
    // */
    //public function testHandle()
    //{
    //    $file = __DIR__.'/resources/test.xlsx';
    //    $excel = IOFactory::load($file);
    //
    //    foreach ($excel->getAllSheets() as $index => $sheet) {
    //        $parse = ParseSheet::parse($sheet, [
    //            'index' => [
    //                'is' => function ($value) {
    //                    return '序号' === $value;
    //                },
    //                'format' => function ($value) {
    //                    return sprintf('系列号: %s', $value);
    //                },
    //            ],
    //            'name' => [
    //                'is' => function ($value) {
    //                    return '姓名' === $value;
    //                },
    //            ],
    //        ]);
    //
    //        $parse->eachWithCheck(function ($fields, $index) {
    //            return ['index' => 'error'];
    //        });
    //
    //        $parse->dumpErrors();
    //    }
    //
    //    /** 保存错误文件 */
    //    $writer = new Xlsx($excel);
    //    $writer->save(sprintf('%s.error.%s.xlsx', $file, __FUNCTION__));
    //}

    /**
     * @throws Exception
     */
    public function testHandle1()
    {
        $file = __DIR__.'/resources/test.xlsx';
        $excel = IOFactory::load($file);

        foreach ($excel->getAllSheets() as $index => $sheet) {
            $parse = ParseSheet::parse($sheet, [
                'index' => [
                    'is' => function ($value) {
                        return '序号' === $value;
                    },
                    'format' => function ($value) {
                        return sprintf('系列号: %s', $value);
                    },
                ],

                'id_code' => [
                    'is' => function ($value) {
                        return '身份证号' === $value;
                    },
                    'format' => function ($value) {
                        return trim(strtoupper($value), '\'');
                    },
                ],

                'work' => [
                    'is' => function ($value) {
                        return 0 < preg_match("/岗位名称\s+（工种）/", $value);
                    },
                    'format' => function ($value) {
                        return $value;
                    },
                ],

                'code' => [
                    'is' => function ($value) {
                        return '原证书编号' === $value;
                    },
                ],
            ]);

            $parse->eachWithCheck(function ($fields, $index) {
                return [];
            });

            $parse->dumpErrors();

            /** 保存错误文件 */
            $writer = new Xlsx($excel);
            $writer->save(sprintf('%s.error.%s.xlsx', $file, __FUNCTION__));
        }
    }
}
