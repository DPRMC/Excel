<?php

namespace DPRMC\Excel\Tests;

use DPRMC\Excel\Excel;
use DPRMC\Excel\Exceptions\UnableToInitializeOutputFile;
use DPRMC\Excel\Markup;
use org\bovigo\vfs\vfsStreamDirectory;
use org\bovigo\vfs\vfsStream;
use org\bovigo\vfs\vfsStreamWrapperAlreadyRegisteredTestCase;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PHPUnit\Framework\TestCase;
use Exception;


class MarkupTest extends TestCase {

    protected $pathToOutputDirectory = './tests/test_files/output/';
    protected $pathToOutputFile      = './tests/test_files/output/testOutput.xlsx';

    const VFS_ROOT_DIR         = 'vfsRootDir';
    const UNWRITEABLE_DIR_NAME = 'unwriteableDir';


    /**
     * @var  vfsStreamDirectory $vfsRootDirObject The root VFS object created in the setUp() method.
     */
    protected static vfsStreamDirectory $vfsRootDirObject;

    /**
     * @var string $unreadableSourceFilePath The path on the VFS to a source file that is unreadable.
     */
    protected static string $unreadableSourceFilePath;

    protected static string $unwritableSourceFilePath = 'tests' . DIRECTORY_SEPARATOR . 'unwritable';


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
     * @group markup
     */
    public function testTranslateMarkup() {

        $styles = [
            'ins' => [
                Markup::BOLD => true,
                Markup::ITALIC => true,
                Markup::COLOR => Color::COLOR_GREEN
            ],
            'del' => [Markup::BOLD => true,
                      Markup::ITALIC => true,
                      Markup::COLOR => Color::COLOR_RED],
        ];

        $filePath = './tests/test_files/test.xlsx';
        $newFilePath = './tests/test_files/test_colorized.xlsx';

        Markup::translateMarkupOfFile( $filePath,
                                       $newFilePath,
                                       $styles );
    }


    /**
     * @test
     * @group markup
     */
    public function testTranslateMarkupInCell() {
        $this->markTestSkipped();
        $this->markTestIncomplete("This test has not been implemented yet.");

    }


    /**
     * @test
     * @group mucell
     */
    public function testSplitStringInCellByMarkup() {
        $this->markTestSkipped();
// Example usage:
        $html  = "This <ins>adds</ins> and this <del>removes</del> ok?";
        $parts = Markup::splitHTMLString( $html );
        $this->assertNotEmpty( $parts );

        $this->assertEquals( 'This', $parts[ 0 ] );
    }


    /**
     * @test
     * @group mucell
     */
    public function testGetTagFromString() {
        $this->markTestSkipped();
        $html  = "This <ins>adds</ins> and this <del>removes</del> ok?";
        $parts = Markup::splitHTMLString( $html );


        $addsText = $parts[1];
        $tag = Markup::getTagFromString( $addsText );
        $this->assertEquals( 'ins', $tag );
    }

}