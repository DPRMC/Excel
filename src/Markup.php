<?php

namespace DPRMC\Excel;


use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\RichText\Run;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;


class Markup {


    const BOLD   = 'bold';
    const ITALIC = 'italic';
    const COLOR  = 'color';

    const BACKGROUND_COLOR = 'background_color';
    //const UNDERLINE = 'underline';
    //const STRIKETHROUGH = 'strikethrough';
    //const SUPERSCRIPT = 'superscript';
    //const SUBSCRIPT = 'subscript';
    //const FONT_SIZE = 'font_size';
    //const FONT_FAMILY = 'font_family';
    //const FONT_COLOR = 'font_color';
    //const BACKGROUND_COLOR = 'background_color';


    /**
     * @param \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet
     * @param array                                 $styles
     *
     * @return void
     * @throws \Exception
     */
    public static function translateMarkupOfSpreadsheet( Spreadsheet &$spreadsheet, array $styles = [] ): void {
        foreach ( $spreadsheet->getWorksheetIterator() as $i => $worksheet ) {
            //echo "Worksheet: " . $worksheet->getTitle() . "\n";

            $highestRow    = $worksheet->getHighestRow();    // Get the highest row number
            $highestColumn = $worksheet->getHighestColumn(); // Get the highest column number (e.g., 'Z', 'AA', etc.)

            // Method 1: Looping by row and column (more common and generally efficient)
            for ( $row = 1; $row <= $highestRow; ++$row ) {
                for ( $col = 'A'; $col <= $highestColumn; ++$col ) {
                    $cell  = $worksheet->getCell( $col . $row );
                    $value = $cell->getValue();

                    if ( empty( $value ) ):
                        continue;
                    endif;


                    self::translateMarkupInCell( $cell, $styles );

                    //echo "Cell " . $col . $row . ": " . $value . "\n";

                    //$spreadsheet->getSheet( $i )->setCellValueExplicit( $col . $row, $cell->getValue(), DataType::TYPE_STRING );
                    $spreadsheet->getSheet( $i )->getCell( $col . $row )->setValue( $cell );
                }
            }
        }
    }


    /**
     * @param string $filePath
     * @param array  $styles
     *
     * @return void
     * @throws \Exception
     */
    public static function translateMarkupOfFile( string $filePath, string $newFilePath, array $styles = [] ): void {
        $spreadsheet = IOFactory::load( $filePath ); // Replace with your file path

        self::translateMarkupOfSpreadsheet( $spreadsheet, $styles );

        // TODO Write the file back to filepath
        IOFactory::createWriter( $spreadsheet, 'Xlsx' )->save( $newFilePath );
    }


    /**
     * @param \PhpOffice\PhpSpreadsheet\Cell\Cell $cell
     * @param array                               $styles
     *
     * @return void
     * @throws \Exception
     */
    public static function translateMarkupInCell( Cell &$cell, array $styles = [] ): void {
        $value = $cell->getValue();
        if ( empty( $value ) ):
            return;
        endif;

        $parts = self::splitHTMLString( $value );

        $objRichText = new RichText();

        foreach ( $parts as $i => $string ):

            if ( $i < count( $parts ) - 1  ):
                $string .= " ";
            endif;

            if ( Markup::stringStartsWithTag( $string ) ):
                $tag = Markup::getTagFromString( $string );

                // If there is a style rule associated with this tag...
                if ( array_key_exists( $tag, $styles ) ):
                    $string = strip_tags( $string );
                    $objTextWithStyle = $objRichText->createTextRun( $string );
                    //$objTextWithStyle = Markup::applyStylesToTextRun( $objTextWithStyle, $styles[ $tag ] );
                    foreach ( $styles as $tag => $stylesToApply ):
                        foreach ( $stylesToApply as $style => $value ):
                            switch ( $style ):
                                case self::BOLD:
                                    $objTextWithStyle->getFont()->setName("Times New Roman");
                                    $objTextWithStyle->getFont()->setBold( TRUE );
                                    break;

                                case self::ITALIC:
                                    $objTextWithStyle->getFont()->setItalic( TRUE );
                                    break;

                                case self::COLOR:
                                    // Color::COLOR_DARKGREEN
                                    $objTextWithStyle->getFont()->setColor( new Color( $value ) );
                                    break;

                                default:
                                    throw new Exception( "I'm unable to apply the style: " . $style . " The styles are " . print_r( $styles, TRUE ) );

                            endswitch;
                        endforeach;

                    endforeach;


                else:
                    $objRichText->createText( $string );
                endif;
            else:
                $objRichText->createText( $string );
            endif;
        endforeach;

        $cell->setValue( $objRichText );
    }


    /**
     * @param string $string "This <ins>adds</ins> and this <del>removes</del> ok?"
     *
     * @return array An array of "This", "<ins>adds</ins>", "and this", ...
     */
    public static function splitHTMLString( string $string ): array {
        $pattern     = '/\s/';
        $stringParts = preg_split( $pattern, $string );
        $pieces      = [];


        $stringPartsWrappedInTags = [];
        $tagIsOpen                = FALSE;
        foreach ( $stringParts as $stringPart ):

            $trimmedStringPart = trim( $stringPart );

            if ( self::stringStartsWithTag( $stringPart ) && self::stringEndsWithTag( $stringPart ) ):
                $pieces[] = $trimmedStringPart;

            elseif ( self::stringStartsWithTag( $stringPart ) ):

                $tagIsOpen                  = TRUE;
                $stringPartsWrappedInTags   = [];
                $stringPartsWrappedInTags[] = $trimmedStringPart;


            elseif ( self::stringEndsWithTag( $stringPart ) ):
                $stringPartsWrappedInTags[] = $trimmedStringPart;
                $pieces[]                   = implode( ' ', $stringPartsWrappedInTags );
                $tagIsOpen                  = FALSE;

            elseif ( $tagIsOpen ):
                $stringPartsWrappedInTags[] = $trimmedStringPart;

            else:
                $pieces[] = $trimmedStringPart;
            endif;

        endforeach;

        return $pieces;
    }


    /**
     * @param string $string
     *
     * @return bool True if string starts with <X>
     */
    public static function stringStartsWithTag( string $string ): bool {
        $pattern = '/^<[^>]+>/';
        return preg_match( $pattern, $string );
    }


    /**
     * @param string $string
     *
     * @return bool True if string ends with </X>
     */
    public static function stringEndsWithTag( string $string ): bool {
        $pattern = '/<[^>]+>$/';
        return preg_match( $pattern, $string );
    }


    /**
     * @param string $string "<ins>adds</ins>"
     *
     * @return string "ins"
     */
    public static function getTagFromString( string $string ): string {
        $pattern = '/^<([^>]*)>/';
        preg_match( $pattern, $string, $matches );
        return $matches[ 1 ];
    }


    /**
     * @param \PhpOffice\PhpSpreadsheet\RichText\Run $objTextWithStyle
     * @param array                                  $styles
     *
     * @return \PhpOffice\PhpSpreadsheet\RichText\Run
     * @throws \Exception
     */
    public static function applyStylesToTextRun( Run $objTextWithStyle, array $styles ): Run {
        foreach ( $styles as $tag => $stylesToApply ):
            foreach ( $stylesToApply as $style => $value ):
                switch ( $style ):
                    case self::BOLD:
                        $objTextWithStyle->getFont()->setBold( TRUE );
                        break;

                    case self::ITALIC:
                        $objTextWithStyle->getFont()->setItalic( TRUE );
                        break;

                    case self::COLOR:
                        // Color::COLOR_DARKGREEN
                        $objTextWithStyle->getFont()->setColor( new Color( $value ) );
                        break;

                    default:
                        throw new Exception( "I'm unable to apply the style: " . $style );

                endswitch;
            endforeach;
        endforeach;

        return $objTextWithStyle;
    }


}