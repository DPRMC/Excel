<?php

namespace DPRMC\Excel;

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


class Style {

    const CELL_ADDRESS_TYPE_HEADER_CELL = 'cell_address_type_header';
    const CELL_ADDRESS_TYPE_ALL_ROWS    = 'cell_address_type_all_rows';
    const CELL_ADDRESS_TYPE_SINGLE_CELL = 'cell_address_type_single_cell';


    /**
     * @param $spreadsheet
     * @param array $rows
     * @param array $styles
     */
    public static function setStyles( &$spreadsheet, array $rows = [], array $styles = [] ) {
        // I guess you want to create a blank spreadsheet. Go right ahead.
        if ( empty( $rows ) ):
            return;
        endif;

        //$spreadsheet->getActiveSheet()->getStyle('B3:B7')->applyFromArray($styleArray);

        foreach ( $styles as $cellAddress => $styleArray ):
            $cellAddressType = self::getCellAddressType( $cellAddress );
            switch ( $cellAddressType ):
                case self::CELL_ADDRESS_TYPE_HEADER_CELL:
                    break;

                case self::CELL_ADDRESS_TYPE_ALL_ROWS:
                    break;

                case self::CELL_ADDRESS_TYPE_SINGLE_CELL:
                    break;

            endswitch;
        endforeach;


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
                                                           array $rows ): string {
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

    private static function setStyleForHeaderCell(&$spreadsheet){

    }



}