<?php

namespace DPRMC\Excel;

use Carbon\Carbon;
use DPRMC\Excel\Exceptions\UnableToInitializeOutputFile;
use PhpOffice\PhpSpreadsheet\Reader\IReadFilter;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use Exception;

/**
 * Class Excel
 * @package DPRMC
 */
class Excel {

    const FORMAT_NUMERIC                = '0.000000####;[=0]0';
    const CELL_ADDRESS_TYPE_HEADER_CELL = 'cell_address_type_header';
    const CELL_ADDRESS_TYPE_ALL_ROWS    = 'cell_address_type_all_rows';
    const CELL_ADDRESS_TYPE_SINGLE_CELL = 'cell_address_type_single_cell';

    /**
     * @var string
     */
    static $title = 'Default Title';

    /**
     * @var string
     */
    static $subject = 'Default Subject';

    /**
     * @var string
     */
    static $creator = 'DPRMC Labs';

    /**
     * @var string
     */
    static $description = 'Default description.';

    /**
     * @var string
     */
    static $keywords = 'keywords';

    /**
     * @var string
     */
    static $category = 'category';

    /**
     * @var array
     */
    static $columnsThatShouldBeNumbers = [];

    /**
     * @var array
     */
    static $columnsThatShouldBeFormulas = [];

    /**
     * @var array
     */
    static $columnsWithCustomNumberFormats = [];

    /**
     * @var array[]
     */
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
     * @param array $rows
     * @param array $totals
     * @param string $sheetName
     * @param string $path
     * @param array $options
     * @param array $columnDataTypes
     * @param array $columnsWithCustomNumberFormats
     * @param array $styles
     * @param array $columnsWithCustomWidths
     * @param bool $freezeHeader
     * @return string
     * @throws UnableToInitializeOutputFile
     */
    public static function advanced( array  $rows = [],
                                     array  $totals = [],
                                     string $sheetName = 'worksheet',
                                     string $path = '',
                                     array  $options = [],
                                     array  $columnDataTypes = [],
                                     array  $columnsWithCustomNumberFormats = [],
                                     array  $columnsWithCustomWidths = [],
                                     array  $styles = [],
                                     bool   $freezeHeader = TRUE ) {
        try {
            $numeric_columns   = [];
            $formulaic_columns = [];
            foreach ( $columnDataTypes as $column_name => $data_type ) :
                if ( $data_type === DataType::TYPE_NUMERIC ) :
                    $numeric_columns[] = $column_name;
                endif;
                if ( $data_type === DataType::TYPE_FORMULA ) :
                    $formulaic_columns[] = $column_name;
                endif;
            endforeach;
            $spreadsheet = new Spreadsheet();
            $path        = self::getUniqueFilePath( $path );
            self::initializeFile( $path );
            self::setOptions( $spreadsheet, $options );
            self::setOrientationLandscape( $spreadsheet );
            self::setHeaderRow( $spreadsheet, $rows, $columnsWithCustomWidths );
            self::setColumnsThatShouldBeNumbers( $numeric_columns, $rows );
            self::setColumnsThatShouldBeFormulas( $formulaic_columns, $rows );
            self::setColumnsWithCustomNumberFormats( $columnsWithCustomNumberFormats, $rows );
            self::setRows( $spreadsheet, $rows );
            self::setFooterTotals( $spreadsheet, $totals );
            self::setWorksheetTitle( $spreadsheet, $sheetName );
            self::setStyles( $spreadsheet, $rows, $styles );
            self::freezeHeader( $spreadsheet, $freezeHeader );
            self::writeSpreadsheet( $spreadsheet, $path );
        } catch ( Exception $e ) {
            throw $e;
        }

        return $path;
    }


    /**
     * @param $spreadsheet
     * @param array $rows
     * @param array $styles
     */
    protected static function setStyles( &$spreadsheet, array $rows = [], array $styles = [] ) {
        // I guess you want to create a blank spreadsheet. Go right ahead.
        if ( empty( $rows ) ):
            return;
        endif;

        $columns = array_keys( $rows[ 0 ] );

        foreach ( $styles as $cellAddress => $styleArray ):
            $cellAddressType = self::getCellAddressType( $cellAddress );
            switch ( $cellAddressType ):
                case self::CELL_ADDRESS_TYPE_HEADER_CELL:
                    $cell_index = array_search( $cellAddress, $columns );
                    if ( FALSE === $cell_index ) :
                        // The column name in the style array does not exist...so skip it
                        break;
                    endif;
                    $excel_column = self::getExcelColumnFromIndex( $cell_index );
                    $spreadsheet->setActiveSheetIndex( 0 )
                                ->getStyle( $excel_column . '1' )
                                ->applyFromArray( $styleArray );
                    break;

                case self::CELL_ADDRESS_TYPE_ALL_ROWS:
                    $cell_address_parts = explode( ':', $cellAddress );
                    $cell_index         = array_search( $cell_address_parts[ 0 ], $columns );
                    if ( FALSE !== $cell_index ) :
                        foreach ( $rows as $i => $row ) :
                            $excel_column = self::getExcelColumnFromIndex( $cell_index );
                            // Apply to all rows excluding header
                            $excel_address = $excel_column . ( $i + 2 );
                            $spreadsheet->setActiveSheetIndex( 0 )
                                        ->getStyle( $excel_address )
                                        ->applyFromArray( $styleArray );
                        endforeach;
                    endif;
                    break;

                case self::CELL_ADDRESS_TYPE_SINGLE_CELL:
                    $cell_address_parts = explode( ':', $cellAddress );
                    $cell_index         = array_search( $cell_address_parts[ 0 ], $columns );
                    if ( FALSE !== $cell_index ) :
                        $excel_column = self::getExcelColumnFromIndex( $cell_index );
                        $spreadsheet->setActiveSheetIndex( 0 )
                                    ->getStyle( $excel_column . $cell_address_parts[ 1 ] )
                                    ->applyFromArray( $styleArray );
                    endif;
                    break;

            endswitch;
        endforeach;
    }

    protected static function freezeHeader( &$spreadsheet, bool $freezeHeader = TRUE ) {

        if ( $freezeHeader ) :
            $spreadsheet->getActiveSheet()->freezePane( 'A2' );
        endif;
    }


    /**
     * @param string $cellAddress
     * @return string
     * @throws Exception
     */
    private static function getCellAddressType( string $cellAddress ): string {
        $cellAddressParts = explode( ':', $cellAddress );
        if ( 1 == sizeof( $cellAddressParts ) ):
            return self::CELL_ADDRESS_TYPE_HEADER_CELL;
        endif;

        if ( '*' == $cellAddressParts[ 1 ] ):
            return self::CELL_ADDRESS_TYPE_ALL_ROWS;
        endif;

        if ( is_numeric( $cellAddressParts[ 1 ] ) ):
            return self::CELL_ADDRESS_TYPE_SINGLE_CELL;
        endif;

        throw new Exception( "I'm unable to determine the cell address type of: " . $cellAddress );
    }


    /**
     * @param string $headerLabel
     * @param array $rows
     * @return string Ex: D1, Z1, EE1, etc
     * @throws Exception
     */
    private static function getHeaderCellAddressFromLabel( string $headerLabel,
                                                           array  $rows ): string {
        // Blank spreadsheet? Ok...
        if ( empty( $rows ) ):
            return '';
        endif;

        $firstRow  = array_shift( $rows );
        $startChar = 'A';
        foreach ( $firstRow as $label => $value ):
            if ( $headerLabel == $value ):
                return $startChar . 1;
            endif;
            $startChar++;
        endforeach;
        throw new Exception( "I could not find a header named: " . $headerLabel );
    }


    /**
     * A wrapper around the PhpSpreadsheet library to make consistently formatted spreadsheets.
     * @param array $rows
     * @param array $totals
     * @param string $sheetName
     * @param string $path
     * @param array $options
     * @param array $columnsThatShouldBeNumbers
     * @param array $columnsWithCustomNumberFormats
     * @return string
     * @throws UnableToInitializeOutputFile
     */
    public static function simple( array  $rows = [],
                                   array  $totals = [],
                                   string $sheetName = 'worksheet',
                                   string $path = '',
                                   array  $options = [],
                                   array  $columnsThatShouldBeNumbers = [],
                                   array  $columnsWithCustomNumberFormats = [] ) {
        try {

            $spreadsheet = new Spreadsheet();
            $path        = self::getUniqueFilePath( $path );
            self::initializeFile( $path );
            self::setOptions( $spreadsheet, $options );
            self::setOrientationLandscape( $spreadsheet );
            self::setHeaderRow( $spreadsheet, $rows );
            self::setColumnsThatShouldBeNumbers( $columnsThatShouldBeNumbers, $rows );
            self::setColumnsWithCustomNumberFormats( $columnsWithCustomNumberFormats, $rows );
            self::setRows( $spreadsheet, $rows );
            self::setFooterTotals( $spreadsheet, $totals );
            self::setWorksheetTitle( $spreadsheet, $sheetName );
            self::writeSpreadsheet( $spreadsheet, $path );
        } catch ( Exception $e ) {
            throw $e;
        }

        return $path;
    }


    /**
     * @param string $path
     * @param null $sheetName This should be a string containing a single worksheet name.
     * @param IReadFilter|null $readFilter Only want specific columns, use this parameter.
     * @param null $nullValue
     * @param bool $calculateFormulas
     * @param bool $formatData Set to false if you want total precision of numbers, and not formatted.
     * @param bool $returnCellRef
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function sheetToArray( string      $path,
                                                     $sheetName = NULL,
                                         IReadFilter $readFilter = NULL,
                                                     $nullValue = NULL,
                                         bool        $calculateFormulas = TRUE,
                                         bool        $formatData = TRUE,
                                         bool        $returnCellRef = FALSE ): array {
        $path_parts    = pathinfo( $path );
        $fileExtension = $path_parts[ 'extension' ];
        $fileExtension = strtolower($fileExtension);

        switch ( $fileExtension ):
            case 'xls':
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
                break;
            case 'xlxs':
            default:
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
                break;
        endswitch;
        $reader->setLoadSheetsOnly( $sheetName );

        // 2023-02-09:mdd
        $reader->setReadDataOnly( TRUE );

        if ( $readFilter ):
            $reader->setReadFilter( $readFilter );
        endif;

        $spreadsheet = $reader->load( $path );
        return $spreadsheet->setActiveSheetIndexByName( $sheetName )->toArray( $nullValue,
                                                                               $calculateFormulas,
                                                                               $formatData,
                                                                               $returnCellRef );
    }


    /**
     * Work in progress...
     * @param string $path
     * @param $sheetName
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function sheetHeaderToArray( string $path, $sheetName = NULL ): array {
        $headers       = [];
        $path_parts    = pathinfo( $path );
        $fileExtension = $path_parts[ 'extension' ];
        $fileExtension = strtolower( $fileExtension );

        $sheetName = 'LL_Res_LOC';
        switch ( $fileExtension ):
            case 'xls':
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
                break;
            case 'xlxs':
            default:
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
                break;
        endswitch;
        $reader->setLoadSheetsOnly( $sheetName );

        // 2023-02-09:mdd
        $reader->setReadDataOnly( TRUE );

//        if ( $readFilter ):
//            $reader->setReadFilter( $readFilter );
//        endif;
        $spreadsheet  = $reader->load( $path );
        $headerFooter = $spreadsheet->setActiveSheetIndexByName( $sheetName )->getHeaderFooter();
//        var_dump( $headerFooter->getEvenHeader() );
//        var_dump( $headerFooter->getOddHeader() );
//        var_dump( $headerFooter->getFirstHeader() );
//        var_dump( $headerFooter->getDifferentFirst() );
//        var_dump( $headerFooter->getImages() );
//        var_dump( $headerFooter->getScaleWithDocument() );
//        var_dump( Carbon::createFromTimestamp($spreadsheet->getProperties()->getCreated(), 'America/New_York')->toDateString() );
//        var_dump( Carbon::createFromTimestamp($spreadsheet->getProperties()->getModified(), 'America/New_York')->toDateString() );
//
//        dump($spreadsheet->getProperties()->getSubject());
//        dump($spreadsheet->getProperties()->getCategory());
//        dump($spreadsheet->getProperties()->getDescription());


//        $spreadsheet->setActiveSheetIndexByName( $sheetName )->

//        dump($sheetName);
//        dd('done');
        return $headers;
    }


    /**
     * @param string $path
     * @param int|null $index This should be the index of the sheet.
     * @param IReadFilter|null $readFilter // Only want specific columns, use this parameter.
     * @param null $nullValue
     * @param bool $calculateFormulas
     * @param bool $formatData // Set to false if you want total precision of numbers, and not formatted.
     * @param bool $returnCellRef
     * @return array
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function sheetByIndexToArray( string $path, int $index = NULL, IReadFilter $readFilter = NULL,
                                                       $nullValue = NULL,
                                                bool   $calculateFormulas = TRUE,
                                                bool   $formatData = TRUE,
                                                bool   $returnCellRef = FALSE ): array {
        $path_parts    = pathinfo( $path );
        $fileExtension = $path_parts[ 'extension' ];
        $fileExtension = strtolower( $fileExtension );

        switch ( $fileExtension ):
            case 'xls':
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
                break;
            case 'xlxs':
            default:
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
                break;
        endswitch;
//        $reader->setLoadSheetsOnly( $sheetName );
        $reader->setLoadAllSheets();

        if ( $readFilter ):
            $reader->setReadFilter( $readFilter );
        endif;

        $spreadsheet = $reader->load( $path );

        return $spreadsheet->setActiveSheetIndex( $index )->toArray( $nullValue,
                                                                     $calculateFormulas,
                                                                     $formatData,
                                                                     $returnCellRef );
    }


    /**
     * Given a zero based index from a php array, this method will return the Excel column equivalent.
     * @param $index
     * @return string
     */
    public static function getExcelColumnFromIndex( $index ) {
        $numeric = $index % 26;
        $letter  = chr( 65 + $numeric );
        $index2  = intval( $index / 26 );
        if ( $index2 > 0 ):
            return self::getExcelColumnFromIndex( $index2 - 1 ) . $letter;
        else:
            return $letter;
        endif;
    }


    /**
     * @param string $excelColumnLetters XFD
     * @return int 16383
     */
    public static function getPhpArrayIndexFromExcelColumn( string $excelColumnLetters ): int {
        $excelColumnLetters = strtoupper( $excelColumnLetters );
        $array              = str_split( $excelColumnLetters );
        $placeCounter       = 0;

        // Reverse only the keys in the array.
        // Leave the letters (values) in the order they were.
        // The key will serve as the exponent.
        $array = array_combine( array_reverse( array_keys( $array ) ), $array );

        //
        foreach ( $array as $count => $letter ):
            $singleLetterValue = ord( $letter ) - 64;
            $baseOfPlace       = pow( 26, $count );
            $valueOfLetter     = $baseOfPlace * $singleLetterValue;
            $placeCounter      += $valueOfLetter;
        endforeach;
        $phpIndex = $placeCounter - 1;
        return $phpIndex;
    }


    /**
     * @param $path
     * @param int $sheetIndex
     * @return string
     */
    public static function getSheetName( $path, $sheetIndex = 0 ): string {

        $path_parts    = pathinfo( $path );
        $fileExtension = $path_parts[ 'extension' ];
        $fileExtension = strtolower( $fileExtension );

        switch ( $fileExtension ):
            case 'csv':
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
                break;
            case 'xls':
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
                break;
            case 'xlxs':
            default:
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
                break;
        endswitch;

        $sheetNames = $reader->listWorksheetNames( $path );
        return (string)$sheetNames[ $sheetIndex ];
    }


    /**
     * Some other methods here require the sheet index, instead of the name.
     * Why not have a little helper to get that for you.
     * @param string $path The absolute path to the spreadsheet file.
     * @param string $sheetName The sheet (tab) name that you want the index of.
     * @return int The numeric index of the sheet in question.
     * @throws Exception Thrown if unable to locate a sheet with $sheetName
     */
    public static function getSheetIndexByName( string $path, string $sheetName ): int {
        $sheetNames = self::getSheetNames( $path );
        $index      = array_search( $sheetName, $sheetNames );

        if ( FALSE === $index ):
            throw new \Exception( "Unable to find sheet named: " . $sheetName );
        endif;

        return (int)$index;
    }


    /**
     * @param $path
     * @return array
     */
    public static function getSheetNames( $path ): array {

        $path_parts    = pathinfo( $path );
        $fileExtension = $path_parts[ 'extension' ];
        $fileExtension = strtolower( $fileExtension );

        switch ( $fileExtension ):
            case 'csv':
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
                break;
            case 'xls':
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
                break;
            case 'xlxs':
            default:
                $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
                break;
        endswitch;

        return $reader->listWorksheetNames( $path );
    }


    /**
     * Returns the number of lines in the sheet.
     * @param $path
     * @param int $sheetIndex
     * @return int
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function numLinesInSheet( $path, $sheetIndex = 0 ): int {
        $emptySheetAsArray = [
            0 => [
                0 => NULL,
            ],
        ];
        $sheetName         = Excel::getSheetName( $path, $sheetIndex );
        $sheetAsArray      = self::sheetToArray( $path, $sheetName );
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
     * @throws Exception
     */
    public static function splitSheet( string $path, int $sheetIndex = 0, int $maxLinesPerFile = 100 ): array {
        $sheetName         = Excel::getSheetName( $path, $sheetIndex );
        $pathsToSplitFiles = [];
        $sheetAsArray      = self::sheetToArray( $path, $sheetName );
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
     * @throws Exception
     */
    protected static function getUniqueFilePath( $startingPath = '' ) {
        if ( file_exists( $startingPath ) ) {
            $filename_ext = pathinfo( $startingPath, PATHINFO_EXTENSION );
            $startingPath =
                preg_replace( '/^(.*)\.' . $filename_ext . '$/', '$1_' . date( 'YmdHis' ) . '.' . $filename_ext, $startingPath );

            if ( is_null( $startingPath ) ) {
                throw new Exception( "The php function preg_replace (called in Excel::getUniqueFilePath()) returned null, indicating an error." );
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
        if ( FALSE === $bytes_written ):
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
     * @throws Exception
     */
    protected static function setOrientationLandscape( &$spreadsheet ) {
        $spreadsheet->getActiveSheet()
                    ->getPageSetup()
                    ->setOrientation( PageSetup::ORIENTATION_LANDSCAPE );
    }


    /**
     * @param $spreadsheet
     * @param array $rows
     */
    protected static function setHeaderRow( &$spreadsheet, $rows = [], $columnsWithCustomWidths = [] ) {

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

            if ( array_key_exists( $field, $columnsWithCustomWidths ) ) :
                $spreadsheet->getActiveSheet()
                            ->getColumnDimension( $startChar )
                            ->setWidth( $columnsWithCustomWidths[ $field ] );
            else :
                $spreadsheet->setActiveSheetIndex( 0 )->getColumnDimension( $startChar )->setAutoSize( TRUE );
            endif;

            $spreadsheet->setActiveSheetIndex( 0 )->getStyle( $startChar . '1' )->getAlignment()->setWrapText( TRUE );
            $startChar++;
        }
    }


    /**
     * @param $spreadsheet
     * @param $rows
     */
    protected static function setRows( &$spreadsheet, $rows ) {
        if ( empty( $rows ) ):
            return;
        endif;

        for ( $i = 0; $i < count( $rows ); $i++ ):
            $startChar = 'A';
            foreach ( $rows[ $i ] as $j => $value ):
                $iProperIndex   = $i + 2; // The data should start one row below the header.
                $cellCoordinate = $startChar . $iProperIndex;

                if ( self::shouldBeNumeric( $startChar ) ):
                    self::setNumericCell( $spreadsheet, $cellCoordinate, $value, self::hasCustomNumberFormat( $startChar ) ?
                        self::$columnsWithCustomNumberFormats[ $startChar ] : self::FORMAT_NUMERIC );

                elseif ( self::shouldBeFormulaic( $startChar ) ):
                    self::setFormulaicCell( $spreadsheet, $cellCoordinate, $value, self::hasCustomNumberFormat( $startChar ) ?
                        self::$columnsWithCustomNumberFormats[ $startChar ] : '' );
                else :
                    self::setTextCell( $spreadsheet, $cellCoordinate, $value, self::hasCustomNumberFormat( $startChar ) ?
                        self::$columnsWithCustomNumberFormats[ $startChar ] : '' );
                endif;


                $startChar++;
            endforeach;
        endfor;
    }


    /**
     *
     * @param string $startChar
     * @return bool
     */
    protected static function shouldBeNumeric( string $startChar ): bool {
        if ( array_key_exists( $startChar, self::$columnsThatShouldBeNumbers ) ):
            return TRUE;
        endif;
        return FALSE;
    }

    /**
     * @param string $startChar
     * @return bool
     */
    protected static function hasCustomNumberFormat( string $startChar ) {
        if ( array_key_exists( $startChar, self::$columnsWithCustomNumberFormats ) ):
            return TRUE;
        endif;
        return FALSE;
    }

    /**
     *
     * @param string $startChar
     * @return bool
     */
    protected static function shouldBeFormulaic( string $startChar ): bool {
        if ( array_key_exists( $startChar, self::$columnsThatShouldBeFormulas ) ):
            return TRUE;
        endif;
        return FALSE;
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @param array $totals
     *
     * @throws Exception
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
                throw new Exception( "EXCEPTION: " . $field . " was not found in " . print_r( $aHeaderMap, TRUE ) );
            endif;


            // If the value is not scalar, then
            if ( is_array( $value ) ):
                $multiDimensionalFooterRow = $footerRowStart;
                foreach ( $value as $name => $childValue ):
                    $cell_coordinate = $columnLetter . $multiDimensionalFooterRow;
                    if ( self::shouldBeNumeric( $columnLetter ) ):
                        self::setNumericCell( $spreadsheet, $cell_coordinate, $childValue, self::hasCustomNumberFormat( $columnLetter ) ?
                            self::$columnsWithCustomNumberFormats[ $columnLetter ] : '' );

                    elseif ( self::shouldBeFormulaic( $columnLetter ) ):
                        self::setFormulaicCell( $spreadsheet, $cell_coordinate, $childValue, self::hasCustomNumberFormat( $columnLetter ) ?
                            self::$columnsWithCustomNumberFormats[ $columnLetter ] : '' );

                    else:
                        self::setTextCell( $spreadsheet, $cell_coordinate, $childValue );
                    endif;

                    $multiDimensionalFooterRow++;
                endforeach;
            else:
                $cell_coordinate = $columnLetter . $footerRowStart;
                if ( self::shouldBeNumeric( $columnLetter ) ) :
                    self::setNumericCell( $spreadsheet, $cell_coordinate, $value, self::hasCustomNumberFormat( $columnLetter ) ?
                        self::$columnsWithCustomNumberFormats[ $columnLetter ] : '' );

                elseif ( self::shouldBeFormulaic( $columnLetter ) ) :
                    self::setFormulaicCell( $spreadsheet, $cell_coordinate, $value, self::hasCustomNumberFormat( $columnLetter ) ?
                        self::$columnsWithCustomNumberFormats[ $columnLetter ] : '' );


                else:
                    self::setTextCell( $spreadsheet, $cell_coordinate, $value );
                endif;

            endif;

        endforeach;
    }

    /**
     * @param $spreadsheet
     * @param string $worksheetName
     * @throws Exception
     */
    protected static function setWorksheetTitle( &$spreadsheet, $worksheetName = 'worksheet' ) {

        if ( empty( $worksheetName ) ):
            throw new Exception( "The work sheet name is empty. You need to supply a name to create a spread sheet." );
        endif;

        $spreadsheet->getActiveSheet()
                    ->setTitle( $worksheetName );


    }


    /**
     * Send an array of column columns that should be treated as numeric
     * @param array $columnsThatShouldBeNumbers
     * @param array $rows
     * @throws Exception
     */
    protected static function setColumnsThatShouldBeNumbers( array $columnsThatShouldBeNumbers, array $rows ) {
        if ( empty( $rows ) ):
            return;
        endif;

        $columnsWithExcelIndexes = [];

        $firstRow = $rows[ 0 ];
        $keys     = array_keys( $firstRow );
        foreach ( $columnsThatShouldBeNumbers as $i => $columnName ):
            $indexFromFirstRow = array_search( $columnName, $keys );

            if ( FALSE === $indexFromFirstRow ):
                throw new Exception( "Unable to find the column header named $columnName. Check your list of columns that should be numeric." );
            endif;

            $excelColumnLetter                             = self::getExcelColumnFromIndex( $indexFromFirstRow );
            $columnsWithExcelIndexes[ $excelColumnLetter ] = $columnName;
        endforeach;

        self::$columnsThatShouldBeNumbers = $columnsWithExcelIndexes;
    }

    /**
     * @param array $columnsWithCustomNumberFormats
     * @param array $rows
     * @throws Exception
     */
    protected static function setColumnsWithCustomNumberFormats( array $columnsWithCustomNumberFormats, array $rows ) {
        if ( empty( $rows ) ):
            return;
        endif;

        $columnsWithExcelIndexes = [];

        $firstRow = $rows[ 0 ];
        $keys     = array_keys( $firstRow );
        foreach ( $columnsWithCustomNumberFormats as $columnName => $customNumberFormat ):
            $indexFromFirstRow = array_search( $columnName, $keys );

            if ( FALSE === $indexFromFirstRow ):
                throw new Exception( "Unable to find the column named $columnName to apply custom number formats to. " );
            endif;

            $excelColumnLetter                             = self::getExcelColumnFromIndex( $indexFromFirstRow );
            $columnsWithExcelIndexes[ $excelColumnLetter ] = $customNumberFormat;
        endforeach;

        self::$columnsWithCustomNumberFormats = $columnsWithExcelIndexes;

    }

    /**
     * Send an array of column columns that should be treated as formulas
     * @param array $columnsThatShouldBeFormulas
     * @param array $rows
     * @throws Exception
     */
    protected static function setColumnsThatShouldBeFormulas( array $columnsThatShouldBeFormulas, array $rows ) {
        if ( empty( $rows ) ):
            return;
        endif;

        $columnsWithExcelIndexes = [];

        $firstRow = $rows[ 0 ];
        $keys     = array_keys( $firstRow );
        foreach ( $columnsThatShouldBeFormulas as $i => $columnName ):
            $indexFromFirstRow = array_search( $columnName, $keys );

            if ( FALSE === $indexFromFirstRow ):
                throw new Exception( "Unable to find the column header named $columnName. Check your list of columns that should be formulas." );
            endif;

            $excelColumnLetter                             = self::getExcelColumnFromIndex( $indexFromFirstRow );
            $columnsWithExcelIndexes[ $excelColumnLetter ] = $columnName;
        endforeach;

        self::$columnsThatShouldBeFormulas = $columnsWithExcelIndexes;
    }


    /**
     * @param $spreadsheet
     * @param $path
     * @throws Exception
     */
    protected static function writeSpreadsheet( $spreadsheet, $path ) {
        try {
            $writer = new Xlsx( $spreadsheet );
            $writer->save( $path );
        } catch ( Exception $exception ) {
            throw new Exception( $exception->getMessage() );
        }
    }

    /**
     * @param $spreadsheet
     * @param $cellCoordinate
     * @param $value
     * @param string $customNumberFormat
     * @param int $activeSheetIndex
     */
    protected static function setNumericCell( &$spreadsheet, $cellCoordinate, $value, $customNumberFormat = '', $activeSheetIndex = 0 ) {
        $spreadsheet->setActiveSheetIndex( $activeSheetIndex )
                    ->setCellValueExplicit( $cellCoordinate, $value, is_null( $value ) ? DataType::TYPE_NULL :
                        DataType::TYPE_NUMERIC );
        if ( $customNumberFormat ) :
            $spreadsheet->getActiveSheet()
                        ->getStyle( $cellCoordinate )
                        ->getNumberFormat()
                        ->setFormatCode( $customNumberFormat );
        endif;
    }

    /**
     * @param Spreadsheet $spreadsheet
     * @param $cellCoordinate
     * @param $value
     * @param string $customNumberFormat
     * @param int $activeSheetIndex
     */
    protected static function setFormulaicCell( Spreadsheet &$spreadsheet,
                                                            $cellCoordinate,
                                                            $value,
                                                string      $customNumberFormat = '',
                                                int         $activeSheetIndex = 0 ) {
        $spreadsheet->setActiveSheetIndex( $activeSheetIndex )
            // ->setCellValue( $cellCoordinate, $value, DataType::TYPE_FORMULA );
                    ->setCellValue( $cellCoordinate, $value ); // 2023-03-10:mdd
        if ( $customNumberFormat ) :
            $spreadsheet->getActiveSheet()
                        ->getStyle( $cellCoordinate )
                        ->getNumberFormat()
                        ->setFormatCode( $customNumberFormat );
        endif;
    }

    /**
     * @param $spreadsheet
     * @param $cellCoordinate
     * @param $value
     * @param int $activeSheetIndex
     */
    protected static function setTextCell( &$spreadsheet, $cellCoordinate, $value, $customNumberFormat = '', $activeSheetIndex = 0 ) {
        $spreadsheet->setActiveSheetIndex( $activeSheetIndex )
                    ->setCellValueExplicit( $cellCoordinate, $value, DataType::TYPE_STRING );
        $spreadsheet->getActiveSheet()
                    ->getStyle( $cellCoordinate )
                    ->getNumberFormat()
                    ->setFormatCode( NumberFormat::FORMAT_TEXT );
        if ( $customNumberFormat ) :
            $spreadsheet->getActiveSheet()
                        ->getStyle( $cellCoordinate )
                        ->getNumberFormat()
                        ->setFormatCode( $customNumberFormat );
        endif;
    }


}