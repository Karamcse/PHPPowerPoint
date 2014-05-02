<?php

/*r Error reporting */

error_reporting(E_ALL ^ E_NOTICE);


/** Include path **/

//set_include_path(get_include_path() . PATH_SEPARATOR . '../Classes/');


/** PHPPowerPoint */

require_once 'Excel.php';

$obj=new ExcelPpt();

$objPHPPowerPoint = new PHPPowerPoint();

$objPHPPowerPoint->removeSlideByIndex(0);



$today = date('Y-m-d', mktime(0, 0, 0, date("m") , date("d"), date("Y")));

$today1 = date('M', mktime(0, 0, 0, date("m") , date("d"), date("Y")));

$yesterday = date('Y-m-d', mktime(0, 0, 0, date("m") , date("d") - 1, date("Y")));

$day=date("l");


//echo $yesterday;

$strMonthYear = date('Md-Y', strtotime($yesterday));

$strMonthYear1 = date('M d, Y', strtotime($today));

$filepath="/home/karam/";

$bmawstats=$filepath."karam".$strMonthYear.".xls";


$currentSlide = $obj->createTemplatedSlide($objPHPPowerPoint);

$body = $obj->logoTemplate($currentSlide,$strMonthYear1);

$currentSlide = $obj->createTemplatedSlide($objPHPPowerPoint);

$body = $obj->extractExcel($currentSlide,"$bmdailyindclick",2,0,0,0,0);


$currentSlide = $obj->createTemplatedSlide($objPHPPowerPoint);

$body = $obj->logoTemplate($currentSlide,"Thank You");


$filename="/home/karam/".$today1."_".date("d")."_ppt_blackwhite.pptx";

$file=$obj->saveFile($objPHPPowerPoint,$filename);

?>
