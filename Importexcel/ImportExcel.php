<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ImportExcel implements ImportExcelInterface
{
    public function import_excel($uploadControlName)
    {
        $fileName = $_FILES[$uploadControlName]['name'];
        $file_ext = pathinfo($fileName, PATHINFO_EXTENSION);

        $allowed_ext = ['xls','csv','xlsx'];

        if(in_array($file_ext, $allowed_ext))
        {
            $inputFileNamePath = $_FILES[$uploadControlName]['tmp_name'];
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileNamePath);
            $data = $spreadsheet->getActiveSheet()->toArray();

            $count = "0";

            foreach($data as $row)
            {
                if($count > 0)
                {
                    $invoice_id = $row['0'];
                    $invoice_date = $row['1'];
                    $customer_name = $row['2'];
                    $customer_address = $row['3'];
                    $product_name = $row['4'];
                    $quantity = $row['5'];
                    $price = $row['6'];
                    $total = $row['7'];
                    $grand_total = $row['8'];

                    $pdo = new PDO('sqlite:../excel.db') or die("cannot open the database");
                    $query = "INSERT INTO invoices (invoice_id, invoice_date, customer_name, customer_address, product_name, quantity, price, total, grand_total)
                                     VALUES ('$invoice_id','$invoice_date','$customer_name','$customer_address','$product_name','$quantity','$price','$total','$grand_total')";

                    $pdo->exec($query);
                }
                else
                {
                    $count = "1";
                }
            }
        }
        else
        {
            exit(0);
        }
    }
}