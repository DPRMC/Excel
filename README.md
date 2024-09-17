# Excel v2
[![Latest Stable Version](https://poser.pugx.org/dprmc/excel/version)](https://packagist.org/packages/dprmc/excel) 
[![codecov](https://codecov.io/gh/DPRMC/Excel/branch/master/graph/badge.svg)](https://codecov.io/gh/DPRMC/Excel)
[![Build Status](https://travis-ci.org/DPRMC/Excel.svg?branch=master)](https://travis-ci.org/DPRMC/Excel) 
[![Total Downloads](https://poser.pugx.org/dprmc/excel/downloads)](https://packagist.org/packages/dprmc/excel) 
[![License](https://poser.pugx.org/dprmc/excel/license)](https://packagist.org/packages/dprmc/excel) 
  

A php library that is a wrapper around the PhpSpreadsheet library. 

<code>composer require dprmc/excel</code>

## Usage: Create a Simple Spreadsheet
Below is an example showing usage of this class.

You can see we create a couple of associative arrays:
- $rows
- $totals

This library will take the keys from the $rows array, and make those the column headers.

If the output file already exists, this method will append a timestamp at the end to try to make a unique filename.

```php
$rows[]     = [
    'CUSIP'  => '123456789',
    'DATE'   => '2018-01-01',
    'ACTION' => 'BUY',
];
$totals     = [
    'CUSIP'  => '1',
    'DATE'   => '2',
    'ACTION' => '3',
];

$options = [
    'title'    => "Sample Title",
    'subject'  => "CUSIP List",
    'category' => "Back Office",
];

$pathToFile = Excel::simple( $rows, $totals, "Tab Label", '/outputFile.xlsx', $options );

```

## Usage: Create an Advanced Spreadsheet
Below is an example showing usage of this class.

You can see we create multiple associative arrays:
- $rows
- $totals
- $options
- $columnDataTypes
- $columnsWithCustomNumberFormats
- $columnsWithCustomWidths
- $styles

The '$columnDataTypes' optional associative array parameter will apply the value of the array as the Data Type to the column cells corresponding to the array key.

The '$columnsWithCustomNumberFormats' optional associative array parameter will apply the number format value of the array to the column cells corresponding to the array key.

The '$columnsWithCustomWidths' optional associative array parameter will apply the width value of the array to the column cells corresponding to the array key.

The '$styles' optional associative array parameter will apply the style values of the $styles array to the corresponding column headers, non-header cells, or to a specific cell.

Additionally, an optional boolean parameter '$freezeHeader' will determine if the header row will be frozen.  Defaults to 'TRUE'. 

If the output file already exists, this method will append a timestamp at the end to try to make a unique filename.

```php
$rows[]     = [
    'CUSIP'  => '123456789',
    'DATE'   => '2018-01-01',
    'PRICE'  => '123.45',   
    'ACTION' => 'BUY',
    'FORM'   => '=IFERROR(((E2-D2)/D2),"")'
];
$totals     = [
    'CUSIP'  => '1',
    'DATE'   => '2',
    'PRICE'  => '3',
    'ACTION' => '4',
    'FORM'   => '5'
];

$sheetName = 'Sheet Name';
$pathToFile = '/outputFile.xlsx';
$options = [];

$columnDataTypes = [
    'CUSIP' => DataType::TYPE_STRING,
    'DATE'  => DataType::TYPE_STRING,
    'PRICE' => DataType::TYPE_NUMERIC,
    'FORM'  => DataType::TYPE_FORMULA
];
$columnsWithCustomNumberFormats = [
    'PRICE' => Excel::FORMAT_NUMERIC,
    'FORM'  => NumberFormat::FORMAT_NUMBER
];
$columnsWithCustomWidths = [
    'CUSIO' => 50,
    'PRICE' => 75,
    'FORM' => 100
];
$styles = [
    'CUSIP'   => [ 'font' => [ 'bold' => TRUE ] ], // Apply style to column header
    'CUSIP:*' => [ 'borders' => [ 'top' => [ 'borderStyle' => 'thin'] ] ], // Apply style to all column rows except header row
    'DATE:4'  => [ 'fill' => [ 'fillType' => 'linear', 'rotation' => 90 ] ] // Apply style to cell in column and specified row 
];

$freezeHeader = TRUE;
$pathToFile = Excel::advanced( $rows, $totals, $sheetName, $pathToFile, $options, $columnDataTypes, $columnsWithCustomNumberFormats, $columnsWithCustomWidths, $styles, $freezeHeader );
```

## Usage: Create a Spreadsheet with multiple sheets
A multidimensional associative array is used to create a workbook with multiple sheets.  Each key of the multidimensional array will represent a new sheet within the Workbook. Each value of the multidimensional array follows the formatting of the advanced sheet shown in the example above.

```php
$pathToFile = '/outputFile.xlsx';
$options = [];

$workbook['first sheet'] = [
    'rows'                           => [], // A multidimensional array with each item representing a row on the sheet
    'totals'                         => [], 
    'columnDataTypes'                => [],
    'columnsWithCustomNumberFormats' => [],
    'columnsWithCustomWidths'        => [],
    'styles'                         => [],
    'freezeHeader'                   => TRUE // A boolean value  
];

$workbook['first sheet']['rows'][0] = [
    'CUSIP'     => '123456789',
    'DATE'      => '2024-01-01',
    'ACTION'    => 'BUY',
    'PRICE'     => '123.456',
    'QUANTITY'  => '1'
];

$workbook['first sheet']['rows'][1] = [
    'CUSIP'     => '123456789',
    'DATE'      => '2024-09-01',
    'ACTION'    => 'SELL',
    'PRICE'     => '123.456',
    'QUANTITY'  => '1'
];

$workbook['first sheet']['totals'] = [
    'CUSIP'     => '123456789',
    'DATE'      => '2024-09-17',
    'ACTION'    => '',
    'PRICE'     => '123.456',
    'QUANTITY'  => '0'
];

$workbook['first sheet']['columnDataTypes'] = [
    'CUSIP'     => DataType::TYPE_STRING,
    'ACTION'    => DataType::TYPE_STRING,
    'PRICE'     => DataType::TYPE_NUMERIC,
    'QUANTITY'  => DataType::TYPE_NUMERIC
];

$workbook['first sheet']['columnsWithCustomNumberFormats'] = [
    'PRICE'     => Excel::FORMAT_NUMERIC,
    'QUANTITY'  => Excel::FORMAT_NUMERIC
];

$workbook['first sheet']['columnsWithCustomWidths'] = [
    'CUSIP'    => 50,
    'PRICE'    => 50,
    'ACTION'   => 25,
    'QUANTITY' => 25
];

$workbook['first sheet']['styles'] = [
    'CUSIP' => [
        'font' => ['bold' => TRUE]
    ]            
];

$workbook['second sheet'] = [];
$workbook['second sheet']['rows'][0] = [
    'CUSIP' => '987654321',
    'NAV'   => '1234.56'
];
$workbook['second sheet']['rows'][1] = [
    'CUSIP' => 'ABCDEFGHI',
    'NAV'   => '6543.21'
];

$workbook['second sheet']['totals'] = [];
$workbook['second sheet']['columnDataTypes'] = [
    'CUSIP' => DataType::TYPE_STRING,
    'NAV'   => DataType::TYPE_NUMERIC
];
$workbook['second sheet']['columnsWithCustomNumberFormats'] = ['NAV' => Excel::FORMAT_NUMERIC];
$workbook['second sheet']['columnsWithCustomWidths'] = [];
$workbook['second sheet']['styles'] = [
    'CUSIP' => [
        'font' => ['bold' => TRUE]
    ],
    'NAV' => [
        'font' => ['italic' => TRUE]  
    ]
];                   
$workbook['second sheet']['freezeHeader'] = FALSE;

$pathToFile = Excel::multiSheet( $pathToFile, $options, $workbook );
```
## Usage: Reading a Spreadsheet into a PHP Array
Pass in the path to an XLSX spreadsheet and a sheet name, and this method will return an associative array.
```php
/**  Define a Read Filter class implementing \PhpOffice\PhpSpreadsheet\Reader\IReadFilter  */
class MyReadFilter implements \PhpOffice\PhpSpreadsheet\Reader\IReadFilter
{
    public function readCell($column, $row, $worksheetName = '') {
        //  Read rows 1 to 7 and columns A to E only
        if ($row >= 1 && $row <= 7) {
            if (in_array($column,range('A','E'))) {
                return true;
            }
        }
        return false;
    }
}
/**  Create an Instance of our Read Filter  **/
$filterSubset = new MyReadFilter();


$pathToWorkbook = '/outputFile.xlsx';
$sheetName = 'Security_Pricing_Update';
$array = Excel::sheetToArray($pathToWorkbook, $sheetName, $filterSubset);
print_r($array);
```