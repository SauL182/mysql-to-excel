<?php

$host = '';
$port = '';
$dbName = '';
$dbUser = '';
$dbPass = '';
$tbl_name = ''; //Set db table name to export

$date = date('d-m-Y');
$fileName = $tbl_name."-export-".$date.".xls"; //Dynamic file name

try{
	$pdo = new PDO("mysql:host=".$host.":".$port.";dbname=".$dbName, $dbUser, $dbPass);
	$pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
	
	$pdo->exec('SET NAMES utf8;');
	$select = "SELECT * FROM `".$tbl_name."`";
	
	$export = $pdo->prepare($select);
	$export->execute();
	
	$fields = $export->columnCount();

	for ($i = 0; $i < $fields; $i++) {
		$column = $export->getColumnMeta($i);
		$col_title .= '<Cell ss:StyleID="2"><Data ss:Type="String">'.$column['name'].'</Data></Cell>';
	}
	
	$col_title = '<Row>'.$col_title.'</Row>';
	
	while($row = $export->fetch(\PDO::FETCH_ASSOC)) {
		$line = '';
		foreach($row as $value) {
			if ((!isset($value)) OR ($value == "")) {
				$value = '<Cell ss:StyleID="1"><Data ss:Type="String"></Data></Cell>\t';
			} else {
				$value = str_replace('"', '', $value);
				$value = '<Cell ss:StyleID="1"><Data ss:Type="String">' . $value . '</Data></Cell>\t';
			}
			$line .= $value;
		}
		$data .= trim("<Row>".$line."</Row>")."\n";
	}
	
	$data = str_replace("\r","",$data);
	
	header("Content-Type: application/vnd.ms-excel;");
	header("Content-Disposition: attachment; filename=".$fileName);
	header("Pragma: no-cache");
	header("Expires: 0");
	
	$xls_header = '<?xml version="1.0" encoding="utf-8"?>
	<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">
	<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
	<Author></Author>
	<LastAuthor></LastAuthor>
	<Company></Company>
	</DocumentProperties>
	<Styles>
	<Style ss:ID="1">
	<Alignment ss:Horizontal="Left"/>
	</Style>
	<Style ss:ID="2">
	<Alignment ss:Horizontal="Left"/>
	<Font ss:Bold="1"/>
	</Style>
	
	</Styles>
	<Worksheet ss:Name="Export">
	<Table>';
	
	$xls_footer = '</Table>
	<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
	<Selected/>
	<FreezePanes/>
	<FrozenNoSplit/>
	<SplitHorizontal>1</SplitHorizontal>
	<TopRowBottomPane>1</TopRowBottomPane>
	</WorksheetOptions>
	</Worksheet>
	</Workbook>';
	
	print $xls_header.$col_title.$data.$xls_footer;
	
}
catch(PDOException $e){
	die("ERROR: Could not connect. " . $e->getMessage());
}

unset($pdo);		
?>
