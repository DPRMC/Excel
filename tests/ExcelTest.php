<?php

use DPRMC\Excel;
use PHPUnit\Framework\TestCase;

use org\bovigo\vfs\vfsStreamDirectory;
use org\bovigo\vfs\vfsStream;
use org\bovigo\vfs\visitor\vfsStreamStructureVisitor;

class ExcelTest extends TestCase {

    public function testToArrayCreatesHeader() {

        vfsStream::setup( 'root', 0777, [
            'test.xlsx' => '',
        ] );

        print_r( vfsStream::inspect( new vfsStreamStructureVisitor() )->getStructure() );

        $rows[]    = [
            'CUSIP'  => '123456789',
            'DATE'   => '2018-01-01',
            'ACTION' => 'BUY',
        ];
        $totals    = [];
        $sheetName = 'testOutput.xlsx';
        //$path      = $vfsRootDirObject->url();
        $path    = vfsStream::url( 'root' ) . '/test.xlsx';
        $options = [];

        $pathToFile = Excel::simple( $rows, $totals, $sheetName, $path, $options );

        echo $pathToFile;


    }
}
