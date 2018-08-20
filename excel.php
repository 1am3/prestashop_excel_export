<?php

require 'vendor/autoload.php';
require '../../config/settings.inc.php';
require 'tools.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$styleArray = [
    'borders' => [
        'outline' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR,
            'color' => ['argb' => '000'],
        ],
    ],
];

$month = array( '01' => 'января', '02' => 'февраля','03' => 'марта','04' => 'апреля','05' => 'мая', '06' => 'июня','07' => 'июля','08' => 'августа','09' => 'сентября','10' => 'октября','11' => 'ноября','12' => 'декабря');
//откуда начинать вывод товаров
$current_col = 16;

if(isset($_GET['id'])) {
	$link = mysqli_connect(_DB_SERVER_,_DB_USER_,_DB_PASSWD_,_DB_NAME_) or die('Error DB connect');
	mysqli_set_charset($link, "utf8");

	$query = mysqli_query($link, "SELECT o.id_order,o.date_add,o.total_paid, c.firstname, c.lastname, a.city,a.address1,a.address2  FROM "._DB_PREFIX_."orders o LEFT JOIN "._DB_PREFIX_."customer c ON c.id_customer = o.id_customer LEFT JOIN "._DB_PREFIX_."address a ON a.id_address = o.id_address_delivery WHERE o.id_order = ".(int)$_GET['id']);

	if(mysqli_num_rows($query) > 0) 
	{
		$order = mysqli_fetch_assoc($query);

		$order_items = mysqli_query($link, 'SELECT product_name,product_quantity,product_price,total_price_tax_incl FROM '._DB_PREFIX_.'order_detail WHERE id_order = '.(int)$_GET['id']);
		$tmp = explode(' ', $order['date_add']);
		$date = explode('-', current($tmp));

		$title = 'Счёт наличный от '.$date[2].' '.$month[$date[1]].' '.$date[0].' г';
		
		$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load('template.xlsx');
		$worksheet = $spreadsheet->getActiveSheet();

		$name = $order['firstname'].' '.$order['lastname'];
		$address = 'г.'.$order['city'].','.$order['address1'].' '.$order['address2'];
		$worksheet->getCell('I6')->setValue($title);
		$worksheet->getCell('A9')->setValue($name);
		$worksheet->getCell('A10')->setValue('');
		$worksheet->getCell('A11')->setValue($address);

		$i = 0;
		while($item = mysqli_fetch_assoc($order_items)) {
			$i++;
			$worksheet->getCell('A'.$current_col)->setValue($i);
			$worksheet->getStyle( 'A'.$current_col )->applyFromArray($styleArray);
			
			$worksheet->getCell('B'.$current_col)->setValue($item['product_name']);
			$worksheet->getStyle( 'B'.$current_col )->applyFromArray($styleArray);

			$worksheet->getCell('C'.$current_col)->setValue('шт');
			$worksheet->getStyle( 'C'.$current_col )->applyFromArray($styleArray);

			$worksheet->getCell('D'.$current_col)->setValue($item['product_quantity']);
			$worksheet->getStyle( 'D'.$current_col )->applyFromArray($styleArray);
			
			$worksheet->getCell('E'.$current_col)->setValue(round($item['product_price'],2));
			$worksheet->getStyle( 'E'.$current_col )->applyFromArray($styleArray);

			$worksheet->getCell('F'.$current_col)->setValue(round($item['product_price'],2));
			$worksheet->getStyle( 'F'.$current_col )->applyFromArray($styleArray);

			$worksheet->getCell('G'.$current_col)->setValue('без НДС');
			$worksheet->getStyle( 'G'.$current_col )->applyFromArray($styleArray);

			$worksheet->getCell('H'.$current_col)->setValue('0,00');
			$worksheet->getStyle( 'H'.$current_col )->applyFromArray($styleArray);

			$worksheet->getCell('I'.$current_col)->setValue($item['total_price_tax_incl']);
			$worksheet->getStyle( 'I'.$current_col )->applyFromArray($styleArray);

			$current_col++;
		}
		//итого
		$worksheet->getCell('E'.$current_col)->setValue('Итого:');
		$worksheet->getStyle( 'E'.$current_col )->getFont()->setBold(true);

		$worksheet->getCell('F'.$current_col)->setValue($order['total_paid']);
		$worksheet->getStyle( 'F'.$current_col)->getFont()->setBold(true);
		$worksheet->getStyle( 'F'.$current_col )->applyFromArray($styleArray);

		$worksheet->getCell('H'.$current_col)->setValue('0,00');
		$worksheet->getStyle( 'H'.$current_col)->getFont()->setBold(true);
		$worksheet->getStyle( 'H'.$current_col )->applyFromArray($styleArray);

		$worksheet->getCell('I'.$current_col)->setValue($order['total_paid']);
		$worksheet->getStyle( 'I'.$current_col)->getFont()->setBold(true);
		$worksheet->getStyle( 'I'.$current_col)->applyFromArray($styleArray);

		$current_col = $current_col+3;

		$total = 'Всего к оплате: '.rtrim($order['total_paid'],'0').'р. ('.sum2words($order['total_paid']).') Без НДС';
		$worksheet->getCell('B'.$current_col)->setValue($total);
		$worksheet->getStyle( 'B'.$current_col)->getFont()->setBold(true);
		$worksheet->getStyle( 'B'.$current_col)->applyFromArray($styleArray);


		$writer = new Xlsx($spreadsheet);
		$writer->save('download/'.$order['id_order'].'.xlsx');
	}
	
}