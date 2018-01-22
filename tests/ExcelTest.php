<?php

use DPRMC\Excel;
use PHPUnit\Framework\TestCase;

use org\bovigo\vfs\vfsStreamDirectory;
use org\bovigo\vfs\vfsStream;

class ExcelTest extends TestCase {

    public function testToArrayCreatesHeader() {

        $vfsRootDirObject = vfsStream::setup( 'root', null, [
            'test' => [
                'test.txt' => 'test content',
            ],
        ] );
        file_put_contents( $vfsRootDirObject->url() . '/test.txt', 'test' );

        $rows      = [
            'CUSIP'  => '123456789',
            'DATE'   => '2018-01-01',
            'ACTION' => 'BUY',
        ];
        $totals    = [];
        $sheetName = 'testOutput.xlsx';
        $path      = $vfsRootDirObject->url();
        $options   = [];

        $pathToFile = Excel::simple( $rows, $totals, $sheetName, $path, $options );

        echo $pathToFile;


    }
}
