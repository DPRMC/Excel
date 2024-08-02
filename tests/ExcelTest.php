<?php

namespace DPRMC\Excel\Tests;

use DPRMC\Excel\Excel;
use DPRMC\Excel\Exceptions\UnableToInitializeOutputFile;
use org\bovigo\vfs\vfsStreamDirectory;
use org\bovigo\vfs\vfsStream;
use org\bovigo\vfs\vfsStreamWrapperAlreadyRegisteredTestCase;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PHPUnit\Framework\TestCase;
use Exception;


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

    protected static $unwritableSourceFilePath = 'tests' . DIRECTORY_SEPARATOR . 'unwritable';


    public function setUp(): void {
        self::$vfsRootDirObject         = vfsStream::setup( self::VFS_ROOT_DIR );
        self::$unreadableSourceFilePath = vfsStream::url( self::VFS_ROOT_DIR . DIRECTORY_SEPARATOR . self::UNWRITEABLE_DIR_NAME );

        chmod( $this->pathToOutputDirectory, 0777 );

        chmod( self::$unwritableSourceFilePath, 0400 );
    }


    public function tearDown(): void {
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
     * @group toarray
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
        $sheetName = 'Test';

        $pathToFile   = Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options );
        $sheetAsArray = Excel::sheetToArray( $pathToFile, $sheetName );

        $this->assertEquals( 'CUSIP', $sheetAsArray[ 0 ][ 0 ] );

        Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options );
        $files = scandir( $this->pathToOutputDirectory );


        array_shift( $files ); // .
        array_shift( $files ); // ..

        $this->assertCount( 2, $files ); // My two test files and the .gitignore file.
    }


    /**
     * @test
     * @group ex
     */
    public function unableToInitializeFileShouldThrowException() {
        //$this->markTestIncomplete(
        //    'This test has not been implemented yet.'
        //);

        $this->expectException( UnableToInitializeOutputFile::class );

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

        //Excel::simple( $rows, $totals, $sheetName, self::$unreadableSourceFilePath, $options );
        $asdf = Excel::simple( $rows, $totals, $sheetName, 'tests/unwritable/shouldNeverExist.xlxs', $options );
        var_dump( $asdf );
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
     * @test\
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
     * @group meta
     */
    public function getDescriptionShouldReturnMetaDescription() {
        $sheetName = 'metaDescription';
        $options = ['description' => 'Meta Description'];
        $pathToFile = Excel::advanced( [], [], $sheetName, $this->pathToOutputFile, $options );
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setLoadSheetsOnly( $sheetName );
        $spreadsheet = $reader->load( $pathToFile );
        $retrievedMetaDescription = $spreadsheet->getProperties()->getDescription();
        $this->assertEquals( $options['description'], $retrievedMetaDescription );
    }


    /**
     * @test
     * @group num1
     */
    public function setColumnAsNumericShouldSetNumeric() {
        $rows[]    = [
            'CUSIP'  => '123456789',
            'DATE'   => '2018-01-01',
            'ACTION' => 'BUY',
            'PRICE'  => '123.456',
            'FORMAT' => '123.456'
        ];
        $rows[]    = [
            'CUSIP'  => 'ABC123789',
            'DATE'   => '2019-01-01',
            'ACTION' => 'BUY',
            'PRICE'  => '998.342',
            'FORMAT' => '998.342'
        ];
        $totals    = [
            'CUSIP'  => '1',
            'DATE'   => '2',
            'ACTION' => '3',
            'PRICE'  => '987.654',
            'FORMAT' => '987.654'
        ];
        $options   = [];
        $sheetName = 'num1';

        $numberTypeColumns = [
            'PRICE',
            'FORMAT'
        ];

        $numberTypeColumnsWithCustomNumericFormats = [
            'PRICE'  => NumberFormat::FORMAT_NUMBER,
            'FORMAT' => Excel::FORMAT_NUMERIC,
        ];

        $pathToFile   = Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options, $numberTypeColumns,$numberTypeColumnsWithCustomNumericFormats  );
        $sheetAsArray = Excel::sheetToArray( $pathToFile, $sheetName );
        $this->assertTrue( gettype( $sheetAsArray[ 1 ][ 3 ] ) === 'double' );
    }


    /**
     * @test
     * @group advanced
     */
    public function advancedCreatesSheet() {
        $rows[] = [
            'CUSIP'     => '123456789',
            'DATE'      => '2018-01-01',
            'ACTION'    => 'BUY',
            'PRICE'     => '123.456',
            'NEW PRICE' => '150',
            'FORM'      => '=IFERROR(((E2-D2)/D2),"")'


        ];
        $rows[] = [
            'CUSIP'     => 'ABC123789',
            'DATE'      => '2019-01-01',
            'ACTION'    => 'BUY',
            'PRICE'     => '998.342',
            'NEW PRICE' => "1000.05",
            'FORM'      => '=IFERROR(((E3-D3)/D3),"")'
        ];
        $totals = [
            'CUSIP'     => '1',
            'DATE'      => '2',
            'ACTION'    => '3',
            'PRICE'     => '1121.798',
            'NEW PRICE' => '1150.05',
            'FORM'      => '=IFERROR(((E4-D4)/D4),"")'
        ];


        $testStyle1 =
            [ 'font'      => [
                'bold' => TRUE,
            ],
              'alignment' => [
                  'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
              ],
              'borders'   => [
                  'top' => [
                      'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                  ],
              ],
              'fill'      => [
                  'fillType'   => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
                  'rotation'   => 90,
                  'startColor' => [
                      'argb' => 'A0A0A0',
                  ],
                  'endColor'   => [
                      'argb' => 'FFFFFF',
                  ],
              ],
            ];

        $testStyle2 =
            [ 'font'      => [
                'bold' => TRUE,
            ],
              'alignment' => [
                  'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
              ],
              'borders'   => [
                  'top' => [
                      'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOT,
                  ],
              ],
              'fill'      => [
                  'fillType'   => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
                  'rotation'   => 45,
                  'startColor' => [
                      'argb' => 'CCCCCC',
                  ],
                  'endColor'   => [
                      'argb' => 'FF0000',
                  ],
              ],
            ];

        $testStyle3 =
            [ 'font'      => [
                'bold' => TRUE,
            ],
              'alignment' => [
                  'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
              ],
              'borders'   => [
                  'top' => [
                      'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                  ],
              ],
              'fill'      => [
                  'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,

                  'color' => [
                      'argb' => 'FFFF00',
                  ],

              ],
            ];


        $sheetName       = 'advanced';
        $options         = [];
        $columnDataTypes = [
            'CUSIP'     => DataType::TYPE_STRING,
            'PRICE'     => DataType::TYPE_NUMERIC,
            'NEW PRICE' => DataType::TYPE_NUMERIC,
            'FORM'      => DataType::TYPE_FORMULA
        ];
        $styles          = [
            'CUSIP'   => $testStyle1,
            'ACTION'  => $testStyle2,
            'CUSIP:*' => $testStyle3,
            'DATE:4'  => $testStyle1,
        ];
        $customNumberFormats = [
            'PRICE'     => Excel::FORMAT_NUMERIC,
            'NEW PRICE' => Excel::FORMAT_NUMERIC,
            'FORM'      => Excel::FORMAT_NUMERIC
        ];

        $freezeHeader = true;


        $columnsWithCustomWidths = [
            'PRICE' => 25,
            'NEW PRICE' => 50,
            'FORM' => 100
        ];

        $pathToFile = Excel::advanced( $rows,
                                       $totals,
                                       $sheetName,
                                       $this->pathToOutputFile,
                                       $options,
                                       $columnDataTypes,
                                       $customNumberFormats,
                                       $columnsWithCustomWidths,
                                       $styles,
                                       $freezeHeader);

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setLoadSheetsOnly( $sheetName );


        $spreadsheet = $reader->load( $pathToFile );
        $RGB         = $spreadsheet->getSheet( 0 )->getStyle( 'A1' )->getFill()->getStartColor()->getRGB();

        $this->assertTrue( $RGB === 'A0A0A0' );
    }

    /**
     * @test
     * @group formula
     */
    public function formulaCellShouldCalculateValue() {
        $rows[] = [
            'CUSIP'     => '123456789',
            'DATE'      => '2018-01-01',
            'ACTION'    => 'BUY',
            'PRICE'     => '75',
            'NEW PRICE' => '150',
            'FORM'      => '=IFERROR(((E2-D2)/D2),"")'


        ];
        $rows[] = [
            'CUSIP'     => 'ABC123789',
            'DATE'      => '2019-01-01',
            'ACTION'    => 'BUY',
            'PRICE'     => '200',
            'NEW PRICE' => "150",
            'FORM'      => '=IFERROR(((E3-D3)/D3),"")'
        ];
        $totals = [
            'CUSIP'     => '1',
            'DATE'      => '2',
            'ACTION'    => '3',
            'PRICE'     => '275',
            'NEW PRICE' => '300',
            'FORM'      => '=IFERROR(((E4-D4)/D4),"")'
        ];

        $sheetName       = 'advanced';
        $options         = [];
        $columnDataTypes = [
            'CUSIP'     => DataType::TYPE_STRING,
            'PRICE'     => DataType::TYPE_NUMERIC,
            'NEW PRICE' => DataType::TYPE_NUMERIC,
            'FORM'      => DataType::TYPE_FORMULA
        ];

        $customNumberFormats = [
            'PRICE'     => Excel::FORMAT_NUMERIC,
            'NEW PRICE' => Excel::FORMAT_NUMERIC,
            'FORM'      => Excel::FORMAT_NUMERIC
        ];



        $pathToFile = Excel::advanced( $rows,
            $totals,
            $sheetName,
            $this->pathToOutputFile,
            $options,
            $columnDataTypes,
            $customNumberFormats);

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setLoadSheetsOnly( $sheetName );


        $spreadsheet = $reader->load( $pathToFile );
        $value = $spreadsheet->getActiveSheet()->getCell('F3')->getCalculatedValue();
        $this->assertTrue(  $value === -.25 );

    }

    /**
     * @test
     * @group text
     */
    public function dataTypeStringShouldAcceptNumberFormats()
    {
        $rows[] = [
            'CUSIP'     => '123456789',
        ];
        $totals = [];
        $sheetName       = 'advanced';
        $options         = [];
        $columnDataTypes = [
            'CUSIP'     => DataType::TYPE_STRING
        ];
        $customNumberFormats = [
            'CUSIP'     => NumberFormat::FORMAT_GENERAL
        ];
        $pathToFile = Excel::advanced( $rows,
            $totals,
            $sheetName,
            $this->pathToOutputFile,
            $options,
            $columnDataTypes,
            $customNumberFormats);

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setLoadSheetsOnly( $sheetName );


        $spreadsheet = $reader->load( $pathToFile );
        $numberFormat = $spreadsheet->getActiveSheet()->getStyle('A2')->getNumberFormat();
        $formatCode = $numberFormat->getFormatCode();
        $this->assertTrue(  $formatCode === 'General' );
    }

    /**
     * @test
     * @group exception
     */
    public function invalidFormulaColumnNameShouldThrowException()
    {
        $this->expectException( Exception::class );
        $rows[] = [
            'CUSIP'     => '123456789',
        ];
        $totals = [];
        $sheetName       = 'advanced';
        $options         = [];
        $columnDataTypes = [
            'InvalidName'     => DataType::TYPE_FORMULA
        ];
        $customNumberFormats = [];

        $pathToFile = Excel::advanced( $rows,
            $totals,
            $sheetName,
            $this->pathToOutputFile,
            $options,
            $columnDataTypes,
            $customNumberFormats);


    }

    /**
     * @test
     * @group exception
     */
    public function invalidNumericColumnNameShouldThrowException()
    {
        $this->expectException( Exception::class );
        $rows[] = [
            'CUSIP'     => '123456789',
        ];
        $totals = [];
        $sheetName       = 'advanced';
        $options         = [];
        $columnDataTypes = [];
        $customNumberFormats = [
            'InvalidName'     => NumberFormat::FORMAT_GENERAL
        ];

        $pathToFile = Excel::advanced( $rows,
            $totals,
            $sheetName,
            $this->pathToOutputFile,
            $options,
            $columnDataTypes,
            $customNumberFormats);


    }

    /**
     * @test
     * @group exception
     */
    public function invalidNumberTypeColumnNameShouldThrowException() {
        $this->expectException( Exception::class );
        $rows[]    = [
            'CUSIP'  => '123456789'
        ];
        $totals    = [];
        $options   = [];
        $sheetName = 'num1';

        $numberTypeColumns = [
            'InvalidNumberColumnName'
        ];

        $numberTypeColumnsWithCustomNumericFormats = [];

        $pathToFile   = Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options, $numberTypeColumns,$numberTypeColumnsWithCustomNumericFormats  );

    }

    /**
     * @test
     * @group exception
     */
    public function blankSheetNameShouldThrowException() {
        $this->expectException( Exception::class );
        $rows[]    = [
            'CUSIP'  => '123456789'
        ];
        $totals    = [];
        $options   = [];
        $sheetName = '';

        $numberTypeColumns = [];

        $numberTypeColumnsWithCustomNumericFormats = [];

        $pathToFile   = Excel::simple( $rows, $totals, $sheetName, $this->pathToOutputFile, $options, $numberTypeColumns,$numberTypeColumnsWithCustomNumericFormats  );

    }

    /**
     * @test
     * @group ord
     */
    public function shouldReturnPhpArrayIndexFromExcelColumnLetters() {

        $arrayIndex = Excel::getPhpArrayIndexFromExcelColumn( 'a' );
        $this->assertEquals( 0, $arrayIndex );

        $arrayIndex = Excel::getPhpArrayIndexFromExcelColumn( 'Z' );
        $this->assertEquals( 25, $arrayIndex );

        $arrayIndex = Excel::getPhpArrayIndexFromExcelColumn( 'AA' );
        $this->assertEquals( 26, $arrayIndex );

        $arrayIndex = Excel::getPhpArrayIndexFromExcelColumn( 'GX' );
        $this->assertEquals( 205, $arrayIndex );

        $arrayIndex = Excel::getPhpArrayIndexFromExcelColumn( 'AAA' );
        $this->assertEquals( 702, $arrayIndex );

        $arrayIndex = Excel::getPhpArrayIndexFromExcelColumn( 'AAB' );
        $this->assertEquals( 703, $arrayIndex );

        $arrayIndex = Excel::getPhpArrayIndexFromExcelColumn( 'AAZ' );
        $this->assertEquals( 727, $arrayIndex );

        $arrayIndex = Excel::getPhpArrayIndexFromExcelColumn( 'BAA' );
        $this->assertEquals( 1378, $arrayIndex );

        // The biggest column in Excel is XFD
        $arrayIndex = Excel::getPhpArrayIndexFromExcelColumn( 'XFD' );
        $this->assertEquals( 16383, $arrayIndex );
    }


    /**
     * @test
     * @group split_small
     */
    public function shouldSplitFileWithoutFormattingSmallNumbers() {
        $sourceSheetPath = './tests/test_files/test_split.xlsx';
        $filePaths       = Excel::splitSheet( $sourceSheetPath,
                                              0,
                                              2,
                                              NULL,
                                              NULL,
                                              TRUE,
                                              FALSE,
                                              FALSE,
                                              './tests/test_files/output/',
                                              [ 'price' ] );

        foreach ( $filePaths as $filePath ):
            $parts               = pathinfo( $filePath );
            $destinationFilename = './tests/test_files/output/' . $parts[ 'basename' ] . '.xlsx';
            copy( $filePath, $destinationFilename );
            unset( $filePath );
            $array = Excel::sheetToArray( $destinationFilename,
                                          'Security_Pricing_Update',
                                          NULL,
                                          NULL,
                                          FALSE,
                                          FALSE,
                                          FALSE );

            foreach ( $array as $i => $row ):
                $value = Excel::decimalNotation( $row[ 5 ] );

                if ( is_numeric( $value ) ):
                    $this->assertEquals('0.00001', $value);
                endif;
            endforeach;
        endforeach;
    }
}