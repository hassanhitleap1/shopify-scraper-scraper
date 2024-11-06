<?php
require 'vendor/autoload.php'; // Make sure to install PhpSpreadsheet with Composer

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Function to fetch JSON data from a paginated URL
function fetchData($url)
{
    $options = [
        "http" => [
            "header" => "User-Agent: PHP-Script"
        ]
    ];
    $context = stream_context_create($options);
    $json = file_get_contents($url, false, $context);
    return json_decode($json, true);
}

// Function to format data as per requirements
function formatProductData($product)
{
    return [
        'ID' => '',
        'Name' => mb_substr($product['title'], 0, 250),
        'Description' => $product['body_html'] ?? '',
        'Slug' => '',
        'URL' => '',
        'SKU' => mb_substr($product['variants'][0]['sku'] ?? '', 0, 50),
        'Categories' => '', // Can be empty
        'Status' => 'published',
        'Is Featured' => false,
        'Brand' => $product['vendor'] ?? '',
        'Product Collections' => '',
        'Labels' => '',
        'Taxes' => '',
        'Image' => $product['image']['src'] ?? '',
        'Images' => isset($product['images']) ? implode(', ', array_column($product['images'], 'src')) : '', // Convert array to comma-separated string
        'Price' => (float) ($product['variants'][0]['price'] ?? 0),
        'Product Attributes' => '',
        'Import Type' => 'product',
        'Is Variation Default' => false,
        'Stock Status' => isset($product['variants'][0]['inventory_quantity']) && $product['variants'][0]['inventory_quantity'] > 0 ? 'in_stock' : 'out_of_stock',
        'With Storehouse Management' => true,
        'Quantity' => isset($product['variants'][0]['inventory_quantity']) ? min(max($product['variants'][0]['inventory_quantity'], 0), 100000000) : 0,
        'Sale Price' => '',
        'Start Date' => '',
        'End Date' => '',
        'Weight' => '',
        'Length' => '',
        'Wide' => '',
        'Height' => '',
        'Cost Per Item' => '',
        'Barcode' => mb_substr($product['variants'][0]['barcode'] ?? '', 0, 50),
        'Content' => $product['body_html'] ?? '',
        'Tags' => is_array($product['tags']) ? implode(', ', $product['tags']) : $product['tags'], // Convert array to comma-separated string
        'Product Type' => 'physical',
        'Auto Generate Sku' => false,
        'Generate License Code' => false,
        'Minimum Order Quantity' => 1,
        'Maximum Order Quantity' => 100000000,
        'Vendor' => $product['vendor'] ?? '',
        'Name (AR)' => '',
        'Description (AR)' => '',
        'Content (AR)' => ''
    ];
}



// Main function to scrape and export all pages to an Excel file
function scrapeAndExportToExcel()
{
    $baseUrl = 'https://rock-store.net/products.json';
    $productsData = [];
    $page = 1;

    // Fetch all pages
    while (true) {
        $url = $baseUrl . '?page=' . $page;
        $data = fetchData($url);

        if (empty($data['products']))
            break; // Stop if no more products

        // Format and append products data
        foreach ($data['products'] as $product) {
            $productsData[] = formatProductData($product);
        }
        $page++;
    }

    // Initialize Spreadsheet and add headers
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $headers = array_keys($productsData[0]);
    $sheet->fromArray($headers, NULL, 'A1');

    // Add each product as a row
    $row = 2;
    foreach ($productsData as $product) {
        $sheet->fromArray($product, NULL, 'A' . $row);
        $row++;
    }

    // Save as Excel file
    $writer = new Xlsx($spreadsheet);
    $fileName = 'shopify_products.xlsx';
    $writer->save($fileName);

    echo "Data has been exported to $fileName";
}

scrapeAndExportToExcel();
