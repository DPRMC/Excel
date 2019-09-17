<?php

namespace DPRMC\Excel\Tests;

use DPRMC\Excel;
use org\bovigo\vfs\vfsStreamDirectory;
use org\bovigo\vfs\vfsStream;
use PHPUnit\Framework\TestCase;

class ExcelTest extends TestCase {

    protected $pathToOutputDirectory = './tests/test_files/output/';
    protected $pathToOutputFile      = './tests/test_files/output/testOutput.xlsx';

    const VFS_ROOT_DIR         = 'vfsRootDir';
    const UNWRITEABLE_DIR_NAME = 'unwriteableDir';


    /**
     * @var  vfsStreamDirectory $vfsRootDirObject The root VFS object created in the setUp() method.
     */
    protected static $vfsRootDirObject;

    /**
     * @var string $unreadableSourceFilePath The path on the VFS to a source file that is unreadable.
     */
    protected static $unreadableSourceFilePath;


    public function setUp() {
        self::$vfsRootDirObject         = vfsStream::setup( self::VFS_ROOT_DIR );
        self::$unreadableSourceFilePath = vfsStream::url( self::VFS_ROOT_DIR . DIRECTORY_SEPARATOR . self::UNWRITEABLE_DIR_NAME );
        chmod( self::$unreadableSourceFilePath, '0000' );
    }


    public function tearDown() {
        $files = scandir( $this->pathToOutputDirectory );
        array_shift( $files ); // .
        array_shift( $files ); // ..
        foreach ( $files as $file ):
            if ( '.gitignore' != $file ):
                unlink( $this->pathToOutputDirectory . $file );
            endif;
        endforeach;
    }


    /**
     * @test
     */
    public function toArrayCreatesHeader() {
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

        $pathToFile   = Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options );
        $sheetAsArray = Excel::sheetToArray( $pathToFile, $sheetName );

        $this->assertEquals( 'CUSIP', $sheetAsArray[ 0 ][ 0 ] );

        Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options );
        $files = scandir( $this->pathToOutputDirectory );

        array_shift( $files ); // .
        array_shift( $files ); // ..

        $this->assertCount( 3, $files ); // My two test files and the .gitignore file.
    }


    /**
     * @test
     */
    public function unableToInitializeFileShouldThrowException() {
        $this->expectException( \Exception::class );

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

        Excel::simple( $rows, $totals, $sheetName, self::$unreadableSourceFilePath, $options );
    }


    /**
     * @test
     */
    public function toArrayWithArrayInTotalsMakesSecondRowOfFooter() {
        $rows[]    = [
            'CUSIP'  => '123456789',
            'DATE'   => '2018-01-01',
            'ACTION' => 'BUY',
        ];
        $totals    = [
            'CUSIP'  => '1',
            'DATE'   => '2',
            'ACTION' => [ 'A', 'B' ],
        ];
        $options   = [];
        $sheetName = 'testOutput.xlsx';

        $pathToFile   = Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options );
        $sheetAsArray = Excel::sheetToArray( $pathToFile, $sheetName );

        $this->assertEquals( 'B', $sheetAsArray[ 3 ][ 2 ] );
    }


    /**
     * @test
     */
    public function toArrayWithArrayInTotalsThrowsExceptions() {
        $this->expectException( \Exception::class );
        $rows[]    = [
            'CUSIP'  => '123456789',
            'DATE'   => '2018-01-01',
            'ACTION' => 'BUY',
        ];
        $totals    = [
            'CUSIP'                     => '1',
            'DATE'                      => '2',
            'NOT_PRESENT_IN_HEADER_ROW' => '3',
        ];
        $options   = [];
        $sheetName = 'testOutput.xlsx';

        $pathToFile = Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options );
        Excel::sheetToArray( $pathToFile );
    }


    /**
     * @test
     * @group split
     */
    public function splitSheetShouldReturnArrayOfFilePaths() {
        $rows = [];
        for ( $i = 0; $i < 10; $i++ ):
            $rows[] = [
                'CUSIP'  => '123456789',
                'DATE'   => '2018-01-01',
                'ACTION' => 'BUY',
            ];
        endfor;
        $sourceSheetPath = Excel::simple( $rows, [], 'test', $this->pathToOutputFile, [] );
        $filePaths       = Excel::splitSheet( $sourceSheetPath, 0, 6 );

        $this->assertCount( 2, $filePaths );
    }


    /**
     * @test
     */
    public function numLinesInSheetShouldReturnTheNumberOfLines() {
        $rows = [];
        for ( $i = 0; $i < 10; $i++ ):
            $rows[] = [
                'CUSIP'  => '123456789',
                'DATE'   => '2018-01-01',
                'ACTION' => 'BUY',
            ];
        endfor;
        $sourceSheetPath = Excel::simple( $rows, [], 'test', $this->pathToOutputFile, [] );
        $numLinesInSheet = Excel::numLinesInSheet( $sourceSheetPath, 0 );
        $this->assertEquals( 11, $numLinesInSheet );
    }

    /**
     * @test
     */
    public function creatingEmptySpreadsheetShouldNotThrowException() {
        $sourceSheetPath = Excel::simple( [], [], 'test', $this->pathToOutputFile, [] );
        $numLinesInSheet = Excel::numLinesInSheet( $sourceSheetPath, 0 );
        $array           = Excel::sheetToArray( $sourceSheetPath, 'test' );
        $this->assertEquals( 0, $numLinesInSheet );
    }


    /**
     * @test
     * @group list
     */
    public function getSheetNameShouldReturnString() {
        $sourceSheetPath = Excel::simple( [], [], 'test', $this->pathToOutputFile, [] );
        $sheetName       = Excel::getSheetName( $sourceSheetPath, 0 );
        $this->assertEquals( 'test', $sheetName );
    }


    /**
     * @test
     * @group num
     */
    public function setColumnAsNumericShouldSetNumeric() {
        $rows[]    = [
            'CUSIP'  => '123456789',
            'DATE'   => '2018-01-01',
            'ACTION' => 'BUY',
            'PRICE'  => '123.456',
        ];
        $rows[]    = [
            'CUSIP'  => 'ABC123789',
            'DATE'   => '2019-01-01',
            'ACTION' => 'BUY',
            'PRICE'  => '998.342',
        ];
        $totals    = [
            'CUSIP'  => '1',
            'DATE'   => '2',
            'ACTION' => '3',
            'PRICE'  => '987.654',
        ];
        $options   = [];
        $sheetName = 'testOutput.xlsx';

        $numberTypeColumns = [
            'PRICE',
        ];

        $pathToFile   = Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options, $numberTypeColumns );
        $sheetAsArray = Excel::sheetToArray( $pathToFile, $sheetName );

        $this->assertTrue( gettype( $sheetAsArray[ 1 ][ 3 ] ) === 'double' );
    }


}