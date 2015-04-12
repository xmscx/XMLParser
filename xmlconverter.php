<?php



error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

echo "<form action='xmlconverter.php?action=do' method='post' enctype='multipart/form-data'><input type='file' size='50' accept=\"\" name='file'><BR>";
echo "<input type='submit'></form>";







$convertedcomplete = "";

if ($_GET['action'] == "do") {
	$callStartTime = microtime(true);
	/** Include PHPExcel */
	require_once dirname(__FILE__) . '/Classes/PHPExcel.php';


	// Create new PHPExcel object
	echo date('H:i:s') , " Create new PHPExcel object" , EOL;
	$objPHPExcel = new PHPExcel();
	// Set document properties
	echo date('H:i:s') , " Set document properties" , EOL;
	$objPHPExcel->getProperties()->setCreator("Michael Sczepanski")
				->setLastModifiedBy("himself")
				->setTitle("PHPExcel XML Converter")
				->setSubject("PHPExcel XML Converter")
				->setDescription("Blub.")
				->setKeywords("php xml Excel super jimmy")
				->setCategory("xml");



	$file = $_FILES['file']['tmp_name'];
	$origfile = $_FILES['file']['name'];
	$contents = file_get_contents($file);
	//* $contents = str_replace('tns:','',$contents);
	$xml = simplexml_load_string($contents);

	$tns = $xml->getNamespaces(true);
	//* var_dump($tns);
	$child = $xml->children($tns['tns']);
	
	foreach ($child->PartnerKey as $PartnerKey) {
		$Country = $PartnerKey->Country;
		$Brand = $PartnerKey->Brand;
		$PartnerNumber = $PartnerKey->PartnerNumber;
	}
	$IsCumulative = $child->IsCumulative;
	foreach ($child->AccountingDate as $AccountingDate) {
		$AccountingMonth = $AccountingDate->AccountingMonth;
		$AccountingYear = $AccountingDate->AccountingYear;
	}
	$counter = 2;
	//* $convertedcomplete = "Betrieb;Standort;Kumuliert;Jahr;Monat;Marke;Konto;KSt;Abs;Ktr;Saldo;\n";
	$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue('A1', 'Betrieb')
				->setCellValue('B1', 'Standort')
				->setCellValue('C1', 'Kumuliert')
				->setCellValue('D1', 'Jahr')
				->setCellValue('E1', 'Monat')
				->setCellValue('F1', 'Marke')
				->setCellValue('G1', 'Konto')
				->setCellValue('H1', 'KSt')
				->setCellValue('I1', 'Abs')
				->setCellValue('J1', 'Ktr')
				->setCellValue('K1', 'Saldo');
	
	foreach ($child->Accounts as $Accounts) {
		foreach ($Accounts as $Attributes) {
			$ProfitCenter = $Attributes->ProfitCenter;
			$AccountKey = $Attributes->AccountKey;
			$AccountValue = $Attributes->AccountValue;
			$Marke = substr($AccountKey, 2, 4);
			$Konto = substr($AccountKey, 6, 4);
			$Kst = substr($AccountKey, 10, 2);
			$Abs = substr($AccountKey, 12, 2);
			$Ktr = substr($AccountKey, 14, 2);
			$Standort = substr($AccountKey, 16, 2);
			$Betrieb = $Country.$PartnerNumber.$Brand;
			//* $AccountValue = str_replace('.',',',$AccountValue);
			$AccountValue = str_replace('+','',$AccountValue);
			
			//* $convertedcomplete .= $Country.$PartnerNumber.$Brand.";".$Standort.";".$IsCumulative.";".$AccountingYear.";".$AccountingMonth.";".$Marke.";".$Konto.";".$Kst.";".$Abs.";".$Ktr.";".$AccountValue.";\n";
			$A = "A".$counter;
			$B = "B".$counter;
			$C = "C".$counter;
			$D = "D".$counter;
			$E = "E".$counter;
			$F = "F".$counter;
			$G = "G".$counter;
			$H = "H".$counter;
			$I = "I".$counter;
			$J = "J".$counter;
			$K = "K".$counter;

			$objPHPExcel->getActiveSheet()->getStyle($B)->getNumberFormat()->setFormatCode('00');			
			$objPHPExcel->getActiveSheet()->getStyle($F)->getNumberFormat()->setFormatCode('0000');
			$objPHPExcel->getActiveSheet()->getStyle($G)->getNumberFormat()->setFormatCode('0000');
			$objPHPExcel->getActiveSheet()->getStyle($H)->getNumberFormat()->setFormatCode('00');
			$objPHPExcel->getActiveSheet()->getStyle($I)->getNumberFormat()->setFormatCode('00');
			$objPHPExcel->getActiveSheet()->getStyle($J)->getNumberFormat()->setFormatCode('00');
			$objPHPExcel->getActiveSheet()->getStyle($K)->getNumberFormat()->setFormatCode('[black]#,##0.00 €;[red]-#,##0.00 €');
			if ($counter%2 == 0) {
				$objPHPExcel->getActiveSheet()->getStyle('A'.$counter.':K'.$counter)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID)->getStartColor()->setARGB('848482');
			}
			//* echo "$Standort<BR>";
			$objPHPExcel->setActiveSheetIndex(0)
						->setCellValue($A, $Betrieb)
						->setCellValue($B, $Standort)
						->setCellValue($C, $IsCumulative)
						->setCellValue($D, $AccountingYear)
						->setCellValue($E, $AccountingMonth)
						->setCellValue($F, $Marke)
						->setCellValue($G, $Konto)
						->setCellValue($H, $Kst)
						->setCellValue($I, $Abs)
						->setCellValue($J, $Ktr)
						->setCellValue($K, $AccountValue);
			
			// Add some data
			//* echo date('H:i:s') , " Add some data" , EOL;

			$counter++;
		}
	}
	//* $objPHPExcel->setActiveSheetIndex(0)->getStyle('A2:K1000')->getNumberFormat()->setFormatCode( PHPExcel_Style_NumberFormat::FORMAT_TEXT );
	$counter = $counter-1;
	$objPHPExcel->getActiveSheet()->setAutoFilter('A1:K'.$counter);
	$objPHPExcel->getActiveSheet()->getColumnDimension('K')->setWidth(20);
	$objPHPExcel->getActiveSheet()->setTitle('XMLParser');
	
	
	$dir = "finishedxml/";
	$datestring = date('YmdHis');	
	$ext = pathinfo($file, PATHINFO_EXTENSION);
	$basename = pathinfo($origfile, PATHINFO_BASENAME);
	$newfilename = $dir.$basename.$datestring.".xlsx";
	

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	$objWriter->save($newfilename);
	$callEndTime = microtime(true);
	$callTime = $callEndTime - $callStartTime;

	echo date('H:i:s') , " File written to " , $newfilename;
	echo "<BR>";
	echo 'Call time to write Workbook was ' , sprintf('%.4f',$callTime) , " seconds" , EOL;
	// Echo memory usage
	echo date('H:i:s') , ' Current memory usage: ' , (memory_get_usage(true) / 1024 / 1024) , " MB" , EOL;
	
	//* Generate temp filename

	//* $handle = fopen($newfilename, 'w');
	//* fwrite($handle, $convertedcomplete);
	//* fclose($handle);
	//* file_put_contents($newfilename, $convertedcomplete);
	echo "<BR><a href='".$newfilename."'>Konvertierte Datei herunterladen</a>";
}

?>