<?php
namespace DPRMC\Excel\Tests;
use DPRMC\Excel;
use PHPUnit\Framework\TestCase;

class ExcelTest extends TestCase {

    protected $pathToOutputDirectory = './tests/test_files/output/';
    protected $pathToOutputFile = './tests/test_files/output/testOutput.xlsx';

    public function tearDown() {
        $files = scandir( $this->pathToOutputDirectory );
        array_shift( $files ); // .
        array_shift( $files ); // ..
        foreach ( $files as $file ):
            unlink( $this->pathToOutputDirectory . $file );
        endforeach;
    }

    public function testToArrayCreatesHeader() {
        $rows[]    = [
            'CUSIP'  => '123456789',
            'DATE'   => '2018-01-01',
            'ACTION' => 'BUY',
        ];
        $totals    = [
            'CUSIP'  => '1',
            'DATE'   => '2',
            'ACTION' => '3',
        ];
        $options   = [];
        $sheetName = 'testOutput.xlsx';

        $pathToFile = Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options );
        $sheetAsArray = Excel::sheetToArray( $pathToFile );

        $this->assertEquals( 'CUSIP', $sheetAsArray[ 0 ][ 0 ] );

        Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options );
        $files      = scandir( $this->pathToOutputDirectory );
        array_shift( $files ); // .
        array_shift( $files ); // ..

        $this->assertCount( 3, $files ); // My two test files and the gitignore and blank file.
    }
}
