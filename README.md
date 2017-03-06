# Excel

[![Latest Stable Version](https://poser.pugx.org/dprmc/excel/version)](https://packagist.org/packages/dprmc/excel) [![Build Status](https://travis-ci.org/DPRMC/Excel.svg?branch=master)](https://travis-ci.org/DPRMC/Excel) [![Total Downloads](https://poser.pugx.org/dprmc/excel/downloads)](https://packagist.org/packages/dprmc/excel) [![License](https://poser.pugx.org/dprmc/excel/license)](https://packagist.org/packages/dprmc/excel)  

A php library that is a wrapper around the PhpSpreadsheet library. 

<code>composer require dprmc/excel</code>

## Usage

Below is an example showing usage of this class.

You can see we create a couple of associative arrays:
- $rows
- $totals

This library will take the keys from the $rows array, and make those the column headers.

<code>

        $rows = [];
        foreach($prices as $price):
            $rows[] = [
                'scheme_identifier' => $price->cusip,
                'scheme_name' => 'CUSIP',
                'market_data_authority_name' => 'ABC',
                'action' => 'ADDUPDATE',
                'as_of_date' => date('n/d/Y', strtotime($price->date)),
                'price' => $price->price
            ];
        endforeach;
        
        $totals = [
            'scheme_identifier' = '',
            'scheme_name' => '',
            'market_data_authority_name' => '',
            'action' => '',
            'as_of_date' => '',
            'price' => ''
        ];
        foreach($prices as $price):
            $rows[] = [
                'scheme_identifier' => $price->cusip,
                'scheme_name' => 'CUSIP',
                'market_data_authority_name' => 'ABC',
                'action' => 'ADDUPDATE',
                'as_of_date' => date('n/d/Y', strtotime($price->date)),
                'price' => $price->price
            ];
        endforeach;
        

        $outputPath = storage_path() . '/report_' . $date. '.xlsx';
        
        $options = [
            'title' => "Update Template - Month End Pricing - " . $date,
            'subject' => "Month End Pricing",
            'category' => "Month End Pricing",
        ];

        $path = \DPRMC\Excel::simple($rows,$totals,'Pricing_Update',$outputPath,$options);

</code>
