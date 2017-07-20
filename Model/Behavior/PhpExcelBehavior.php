<?php

// CakePHP friendly wrapper for PHPExcel
// found at: https://github.com/PHPOffice/PHPExcel

require_once ROOT. DS. 'Vendor'. DS. 'phpoffice'.DS.'phpexcel'.DS.'Classes'.DS.'PHPExcel'.DS.'IOFactory.php';

class PhpExcelBehavior extends ModelBehavior
{
	public $settings = array();
	
	protected $_defaults = array(
	);
	
	public $Model = false;
	
	public function setup(Model $Model, $settings = array())
	{
		$this->Model = $Model;
		
		if (!isset($this->settings[$Model->alias])) 
		{
			$this->settings[$Model->alias] = $this->_defaults;
		}
		$this->settings[$Model->alias] = array_merge($this->settings[$Model->alias], $settings);
		
	}
	
	public function Excel_csvToArray(Model $Model, $csvString = false)
	{
		$csvString = trim($csvString);
		if(!$csvString)
		{
			$Model->modelError = __('Invalid or empty CSV String');
			return false;
		}
		$csvArray = [];
		
		$rows = explode("\n", $csvString);
		foreach($rows as $row)
		{
			$row = trim($row);
			if(!$row)
				continue;
			
			if(!$row = str_getcsv($row))
				continue;
			$emptyCnt = 0;
			foreach($row as $k => $v)
			{
				$v = trim($v);
				if(!$v)
					$emptyCnt++;
				$row[$k] = $v;
			}
			if($emptyCnt >= count($row))
				continue;
			$csvArray[] = $row;
		}
		
		return $csvArray;
	}
	
	public function Excel_fileToHtml(Model $Model, $inputFileName = false)
	{
		if(!$inputFileName)
		{
			$Model->modelError = __('Unknown File Path');
			return false;
		}
		if(!is_readable($inputFileName))
		{
			$Model->modelError = __('Unable to read the File');
			return false;
		}
		
		PHPExcel_Calculation_Functions::setReturnDateType(PHPExcel_Calculation_Functions::RETURNDATE_EXCEL);
		
		// find out which type of 
		$inputFileType = PHPExcel_IOFactory::identify($inputFileName);
		$objReader = PHPExcel_IOFactory::createReader($inputFileType);
		$objReader->setReadDataOnly(true);
		$objPHPExcel = $objReader->load($inputFileName);
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'HTML');
		$objWriter->setSheetIndex(0);
		
		ob_start();
		$objWriter->save('php://output');
		$results = ob_get_contents();
		ob_end_clean();
		return $results;
	}
	
	public function Excel_fileToArray(Model $Model, $inputFileName = false, $includeHiddenRows = false)
	{
		if(!$inputFileName)
		{
			$Model->modelError = __('Unknown File Path');
			return false;
		}
		if(!is_readable($inputFileName))
		{
			$Model->modelError = __('Unable to read the File');
			return false;
		}
		
		PHPExcel_Calculation_Functions::setReturnDateType(PHPExcel_Calculation_Functions::RETURNDATE_PHP_NUMERIC);
		
		$baseDate = PHPExcel_Shared_Date::getExcelCalendar();
		
		// find out which type of 
		$inputFileType = PHPExcel_IOFactory::identify($inputFileName);
		$objReader = PHPExcel_IOFactory::createReader($inputFileType);
		$objReader->setReadDataOnly(true);
		
		$objPHPExcel = $objReader->load($inputFileName);
		$objWorksheet = $objPHPExcel->getActiveSheet();
		
		$objWorksheet->getAutoFilter()->showHideRows();
		
		// go through the active worksheet (which is the first sheet, 0 index by default)
		// retrieve the header names
		
		$headers = array();
		$alldata = array();
		
		$first = true;
		foreach ($objWorksheet->getRowIterator() as $row)
		{
			$cellIterator = $row->getCellIterator();
			$cellIterator->setIterateOnlyExistingCells(FALSE);

			// only work with visible rows
			if(!$includeHiddenRows)
				if (!$objWorksheet->getRowDimension($row->getRowIndex())->getVisible())
					continue;
					
			
			$rowdata = array();
			// find the column names
			$cell_i = 0;
			foreach ($cellIterator as $cell)
			{
				$cell_value = $cell->getFormattedValue();
				
				if($first)
				{
					$headers[] = strtolower(Inflector::slug($cell_value));
				}
				else
				{
					$cell_key = (isset($headers[$cell_i])?$headers[$cell_i]:$cell_i);
					
					if($cell_key == 'date')
					{
						$cell_value = $this->Excel_fixDate($Model, $cell->getValue());
					}
					
					if($cell_key) $rowdata[$cell_key] = $cell_value;
				}
				$cell_i++;
			}
			if($rowdata) $alldata[] = $rowdata;
			
			$first = false;
		}
		
		return $alldata;
	}
	
	public function Excel_fixDate(Model $Model, $date = false)
	{
		$date = trim($date);
		if($date)
		{
			$date = PHPExcel_Shared_Date::ExcelToPHP($date);
			$date = strtotime('+1 day', $date);
			$date = date('Y-m-d H:i:s', $date);
		}
		return $date;
	}
}