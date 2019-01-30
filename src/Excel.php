<?php

namespace DPRMC;

use DPRMC\Excel\Exceptions\UnableToInitializeOutputFile;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;

/**
 * Class Excel
 * @package DPRMC
 */
class Excel {

    static $title = 'Default Title';

    static $subject = 'Default Subject';

    static $creator = 'DPRMC Labs';

    static $description = 'Default description.';

    static $keywords = 'keywords';

    static $category = 'category';

    static $headerStyleArray = [
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'color'    => [
                'argb' => Color::COLOR_DARKBLUE,
            ],
        ],
        'font' => [
            'bold'  => TRUE,
            'color' => [
                'argb' => Color::COLOR_WHITE,
            ],
        ],
    ];

    /**
     * A wrapper around the PhpSpreadsheet library to make consistently formatted spreadsheets.
     *
     * @param array $rows
     * @param array $totals
     * @param string $sheetName
     * @param string $path
     * @param array $options
     *
     * @return string
     * @throws \Exception
     */
    public static function simple( array $rows = [], array $totals = [], string $sheetName = 'worksheet', string $path = '', array $options = [] ) {
        try {

            /**
             * @var Spreadsheet $spreadsheet
             */
            $spreadsheet = new Spreadsheet();
            $path        = self::getUniqueFilePath( $path );
            self::initializeFile( $path );
            self::setOptions( $spreadsheet, $options );
            self::setOrientationLandscape( $spreadsheet );
            self::setHeaderRow( $spreadsheet, $rows );
            self::setRows( $spreadsheet, $rows );
            self::setFooterTotals( $spreadsheet, $totals );
            self::setWorksheetTitle( $spreadsheet, $sheetName );
            self::writeSpreadsheet( $spreadsheet, $path );
        } catch ( \Exception $e ) {
            throw $e;
        }

        return $path;
    }

    /**
     * @param string $path
     * @param int $sheetIndex
     *
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function sheetToArray( $path, $sheetIndex = 0 ) {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $reader->setLoadSheetsOnly( [ 0 ] );
        $spreadsheet = $reader->load( $path );
        return $spreadsheet->setActiveSheetIndex( $sheetIndex )->toArray();
    }

    /**
     * @param $path
     * @param int $sheetIndex
     * @return string
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function getSheetName($path, $sheetIndex = 0 ): string {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $sheetNames = $reader->listWorksheetNames($path);
        return (string)$sheetNames[$sheetIndex];
    }

    /**
     * Returns the number of lines in the sheet.
     * @param $path
     * @param int $sheetIndex
     * @return int
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function numLinesInSheet( $path, $sheetIndex = 0 ): int {
        $emptySheetAsArray = [
            0 => [
                0 => NULL,
            ],
        ];
        $sheetAsArray      = self::sheetToArray( $path, $sheetIndex );
        if ( $sheetAsArray == $emptySheetAsArray ):
            return 0;
        endif;
        return count( $sheetAsArray );
    }


    /**
     * @param string $path
     * @param int $sheetIndex
     * @param int $maxLinesPerFile
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function splitSheet( string $path, int $sheetIndex = 0, int $maxLinesPerFile = 100 ): array {
        $sheetName = Excel::getSheetName($path,$sheetIndex);
        $pathsToSplitFiles = [];
        $sheetAsArray      = self::sheetToArray( $path, $sheetIndex );
        $header            = array_shift( $sheetAsArray );
        $chunks            = array_chunk( $sheetAsArray, $maxLinesPerFile );
        foreach ( $chunks as $i => $chunk ):
            $chunk               = self::setHeadersAsIndexes( $chunk, $header );
            $pathsToSplitFiles[] = self::simple( $chunk, [], $sheetName, tempnam( NULL, 'split_' . $i ), [] );
        endforeach;
        return $pathsToSplitFiles;
    }

    /**
     * @param array $rows
     * @param array $headers
     * @return array
     */
    protected static function setHeadersAsIndexes( array $rows, array $headers ): array {
        $modifiedRows = [];
        foreach ( $rows as $i => $row ):
            foreach ( $row as $j => $value ):
                $modifiedRows[ $i ][ $headers[ $j ] ] = $value;
            endforeach;
        endforeach;
        return $modifiedRows;
    }


    /**
     * Given desired startingPath, this method will check if a file already exists at that location.
     * If a file with the same file name exists, this method will return a path with a timestamp
     * appended to the end of the intended file name.
     *
     * @param string $startingPath
     *
     * @return string
     * @throws \Exception
     */
    protected static function getUniqueFilePath( $startingPath = '' ) {
        if ( file_exists( $startingPath ) ) {
            $filename_ext = pathinfo( $startingPath, PATHINFO_EXTENSION );
            $startingPath = preg_replace( '/^(.*)\.' . $filename_ext . '$/', '$1_' . date( 'YmdHis' ) . '.' . $filename_ext, $startingPath );

            if ( is_null( $startingPath ) ) {
                throw new \Exception( "The php function preg_replace (called in Excel::getUniqueFilePath()) returned null, indicating an error." );
            }
        }

        return $startingPath;
    }

    /**
     * @param string $path The destination file path.
     *
     * @throws UnableToInitializeOutputFile
     */
    protected static function initializeFile( $path ) {
        $bytes_written = file_put_contents( $path, '' );
        if ( $bytes_written === FALSE ):
            throw new UnableToInitializeOutputFile( "Unable to write to the file at " . $path );
        endif;
    }

    /**
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet
     * @param                                       $options
     */
    protected static function setOptions( &$spreadsheet, $options = [] ) {
        self::$title       = $options[ 'title' ] ?? self::$title;
        self::$subject     = $options[ 'subject' ] ?? self::$subject;
        self::$creator     = $options[ 'creator' ] ?? self::$creator;
        self::$description = $options[ 'description' ] ?? self::$description;
        self::$keywords    = $options[ 'keywords' ] ?? self::$keywords;
        self::$category    = $options[ 'category' ] ?? self::$category;

        $spreadsheet->getProperties()
                    ->setCreator( self::$creator )
                    ->setLastModifiedBy( self::$creator )
                    ->setTitle( self::$title )
                    ->setSubject( self::$subject )
                    ->setDescription( self::$description )
                    ->setKeywords( self::$keywords )
                    ->setCategory( self::$category );
    }

    /**
     * @param Spreadsheet $spreadsheet
     *
     * @throws \Exception
     */
    protected static function setOrientationLandscape( &$spreadsheet ) {
        $spreadsheet->getActiveSheet()
                    ->getPageSetup()
                    ->setOrientation( PageSetup::ORIENTATION_LANDSCAPE );
    }


    protected static function setHeaderRow( &$spreadsheet, $rows = [] ) {

        // I guess you want to create a blank spreadsheet. Go right ahead.
        if ( empty( $rows ) ):
            return;
        endif;

        // Set header row
        $startChar = 'A';
        foreach ( $rows[ 0 ] as $field => $value ) {
            $spreadsheet->setActiveSheetIndex( 0 )
                        ->setCellValueExplicit( $startChar . '1', $field, DataType::TYPE_STRING );

            $spreadsheet->setActiveSheetIndex( 0 )
                        ->getStyle( $startChar . '1' )
                        ->applyFromArray( self::$headerStyleArray );

            $spreadsheet->setActiveSheetIndex( 0 )
                        ->getColumnDimension( $startChar )
                        ->setAutoSize( TRUE );

            $startChar++;
        }
    }


    protected static function setRows( &$spreadsheet, $rows ) {
        if ( empty( $rows ) ):
            return;
        endif;
        for ( $i = 0; $i < count( $rows ); $i++ ):
            $startChar = 'A';
            foreach ( $rows[ $i ] as $j => $value ):
                $iProperIndex   = $i + 2; // The data should start one row below the header.
                $cellCoordinate = $startChar . $iProperIndex;
                $spreadsheet->setActiveSheetIndex( 0 )
                            ->setCellValueExplicit( $cellCoordinate, $value, DataType::TYPE_STRING );
                $startChar++;
            endforeach;
        endfor;
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @param array $totals
     *
     * @throws \Exception
     */
    protected static function setFooterTotals( &$spreadsheet, $totals ) {
        // Create a map array by iterating through the headers
        $aHeaderMap = [];
        $lastColumn = $spreadsheet->setActiveSheetIndex( 0 )
                                  ->getHighestColumn();
        $lastColumn++; //Because of the != iterator below, we need to tell it to stop one AFTER the last column. Or we could change it to a doWhile loop... This was easier.

        $startColumn = 'A';

        for ( $c = $startColumn; $c != $lastColumn; $c++ ):
            $cellValue        = $spreadsheet->setActiveSheetIndex( 0 )
                                            ->getCell( $c . 1 )
                                            ->getValue();
            $aHeaderMap[ $c ] = $cellValue;
        endfor;


        $lastRow        = $spreadsheet->setActiveSheetIndex( 0 )
                                      ->getHighestRow();
        $footerRowStart = $lastRow + 1;

        foreach ( $totals as $field => $value ):
            // GET THE FIELD LETTER
            // Find the matching value in row one...
            $columnLetter = array_search( $field, $aHeaderMap );

            if ( $columnLetter === FALSE ):
                throw new \Exception( "EXCEPTION: " . $field . " was not found in " . print_r( $aHeaderMap, TRUE ) );
            endif;


            // If the value is not scalar, then
            if ( is_array( $value ) ):
                $multiDimensionalFooterRow = $footerRowStart;
                foreach ( $value as $name => $childValue ):
                    $spreadsheet->setActiveSheetIndex( 0 )
                                ->setCellValueExplicit( $columnLetter . $multiDimensionalFooterRow, $childValue, DataType::TYPE_STRING );
                    $multiDimensionalFooterRow++;
                endforeach;
            else:
                $spreadsheet->setActiveSheetIndex( 0 )
                            ->setCellValueExplicit( $columnLetter . $footerRowStart, $value, DataType::TYPE_STRING );
            endif;

        endforeach;
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @param string $worksheetName
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception;
     */
    protected static function setWorksheetTitle( &$spreadsheet, $worksheetName = 'worksheet' ) {
        $spreadsheet->getActiveSheet()
                    ->setTitle( $worksheetName );
    }

    /**
     * @param $spreadsheet
     * @param $path
     * @throws \Exception
     */
    protected static function writeSpreadsheet( $spreadsheet, $path ) {
        try {
            $writer = new Xlsx( $spreadsheet );
            $writer->save( $path );
        } catch ( \Exception $exception ) {
            throw new \Exception( $exception->getMessage() );
        }

    }
}