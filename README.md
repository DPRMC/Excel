# Excel v2

[![Latest Stable Version](https://poser.pugx.org/dprmc/excel/version)](https://packagist.org/packages/dprmc/excel) [![Build Status](https://travis-ci.org/DPRMC/Excel.svg?branch=master)](https://travis-ci.org/DPRMC/Excel) [![Total Downloads](https://poser.pugx.org/dprmc/excel/downloads)](https://packagist.org/packages/dprmc/excel) [![License](https://poser.pugx.org/dprmc/excel/license)](https://packagist.org/packages/dprmc/excel) [![Build Status](https://travis-ci.org/DPRMC/Excel.svg?branch=master)](https://travis-ci.org/DPRMC/Excel) [![Coverage Status](https://coveralls.io/repos/github/DPRMC/Excel/badge.svg?branch=master)](https://coveralls.io/github/DPRMC/Excel?branch=master)  

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
    'title'    => "Sample Title,
    'subject'  => "CUSIP List",
    'category' => "Back Office",
];

$pathToFile = Excel::simple( $rows, $totals, "Tab Label", '/outputFile.xlsx', $options );

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