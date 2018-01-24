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
            'bold'  => true,
            'color' => [
                'argb' => Color::COLOR_WHITE,
            ],
        ],
    ];

    /**
     * A wrapper around the PhpSpreadsheet library to make consistently formatted spreadsheets.
     *
     * @param array  $rows
     * @param array  $totals
     * @param string $sheetName
     * @param string $path
     * @param array  $options
     *
     * @return string
     * @throws \Exception
     */
    public static function simple( $rows = [], $totals = [], $sheetName = 'worksheet', $path = '', $options = [] ) {
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
     * @param int    $sheetIndex
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
        if ( $bytes_written === false ):
            throw new UnableToInitializeOutputFile( "Unable to write to the file at " . $path );
        endif;
    }

    /**
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet
     * @param                                       $options
     */
    protected static function setOptions( &$spreadsheet, $options = [] ) {
        self::$title       = isset( $options[ 'title' ] ) ? $options[ 'title' ] : self::$title;
        self::$subject     = isset( $options[ 'subject' ] ) ? $options[ 'subject' ] : self::$subject;
        self::$creator     = isset( $options[ 'creator' ] ) ? $options[ 'creator' ] : self::$creator;
        self::$description = isset( $options[ 'description' ] ) ? $options[ 'description' ] : self::$description;
        self::$keywords    = isset( $options[ 'keywords' ] ) ? $options[ 'keywords' ] : self::$keywords;
        self::$category    = isset( $options[ 'category' ] ) ? $options[ 'category' ] : self::$category;

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
                        ->setAutoSize( true );

            $startChar++;
        }
    }


    protected static function setRows( &$spreadsheet, $rows ) {
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
     * @param array       $totals
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

            if ( $columnLetter === false ):
                throw new \Exception( "EXCEPTION: " . $field . " was not found in " . print_r( $aHeaderMap, true ) );
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
     * @param string      $worksheetName
     *
     * @throws \PhpOffice\PhpSpreadsheet\Exception;
     */
    protected static function setWorksheetTitle( &$spreadsheet, $worksheetName = 'worksheet' ) {
        $spreadsheet->getActiveSheet()
                    ->setTitle( $worksheetName );
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @param string      $path
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected static function writeSpreadsheet( $spreadsheet, $path ) {
        $writer = new Xlsx( $spreadsheet );
        $writer->save( $path );
    }
}