<?php

//App::import('Vendor', 'PHPExcel', array('file' => 'PHPExcel/Classes/PHPExcel/IOFactory.php'));

require_once ROOT. DS. 'Vendor'. DS. 'phpoffice'.DS.'phpexcel'.DS.'Classes'.DS.'PHPExcel'.DS.'IOFactory.php';
App::uses('PhpExcelAppHelper', 'PhpExcel.View/Helper');
/*
 * References:
 *
 * Styles: http://www.bainweb.com/2012/01/phpexcel-style-reference-complete-list.html
 */

class PhpExcelHelper extends PhpExcelAppHelper 
{
	/// data format is multitiered 
	// expects the first level to be sheets in a book
	/* example: 
	array(
		'properties' => array(
			'Creator' => 'Name of creator',
			'LastModifiedBy' => 'Modifier name',
			'Title' => 'Title of book,
			'Subject' => 'Subject of the book',
			'Description' => '',
			'Keywords' => '',
			'Category' => '',
		),
		'sheets' => array(
			'title' => 'Title of the page at the top, if needed'
			'subtitle' => 'If needed',
			//// like csv
			'headers' => array('title', 'another title'),
			'data' => array(
				/// row
				array('value', 'another value'),
			),
			//// OR, custom layout
			'matrix' => array(
				'A1' => 'content here',
				'B2' => 'another cell',
			),
		),
	);
	*/
	
	public $objPHPExcel = false;
	
	public $defaultProperties = array();
	
	public $defaultStyles = array(
		'font.name' => 'Arial',
		'font.size' => 14,
		'font.color.rgb' => PHPExcel_Style_Color::COLOR_BLACK,
	);
	
	public function loadPHPExcel($reload = false)
	{
		if(!$this->objPHPExcel)
		{
			if($this->objPHPExcel = new PHPExcel())
			{
				$this->setDefaults();
				return true;
			}
			throw new NotFoundException(__('Unable to load PHPExcel.'));
		}
		elseif($reload)
		{
			if($this->objPHPExcel = new PHPExcel())
			{
				$this->setDefaults();
				return true;
			}
			throw new NotFoundException(__('Unable to load PHPExcel.'));
		}
		return true;
	}
	
	public function setDefaults()
	{
		// set the default styles
		$this->objPHPExcel->getDefaultStyle()->applyFromArray(Hash::expand($this->defaultStyles));
		// path to fonts for auto sizing
		PHPExcel_Shared_Font::setTrueTypeFontPath('/usr/share/fonts/truetype/msttcorefonts/');
		// allow automatix cell resizing
		PHPExcel_Shared_Font::setAutoSizeMethod(PHPExcel_Shared_Font::AUTOSIZE_METHOD_EXACT);
	}
	
	public function setProperties($properties = array())
	{
		$this->defaultProperties['company'] = Configure::read('Site.title');
		
		$properties = array_merge($this->defaultProperties, $properties);
		
		if(isset($properties['creator']))
			$this->objPHPExcel->getProperties()->setCreator($properties['creator']);
		if(isset($properties['modifier']))
			$this->objPHPExcel->getProperties()->setLastModifiedBy($properties['modifier']);
		if(isset($properties['title']))
			$this->objPHPExcel->getProperties()->setTitle($properties['title']);
		if(isset($properties['subject']))
			$this->objPHPExcel->getProperties()->setSubject($properties['subject']);
		if(isset($properties['description']))
			$this->objPHPExcel->getProperties()->setDescription($properties['description']);
		if(isset($properties['keywords']))
			$this->objPHPExcel->getProperties()->setKeywords($properties['keywords']);
		if(isset($properties['category']))
			$this->objPHPExcel->getProperties()->setCategory($properties['category']);
		if(isset($properties['manager']))
			$this->objPHPExcel->getProperties()->setManager($properties['manager']);
		if(isset($properties['company']))
			$this->objPHPExcel->getProperties()->setCompany($properties['company']);
		if(isset($properties['created']))
			$this->objPHPExcel->getProperties()->setCreated($properties['created']);
		if(isset($properties['modified']))
			$this->objPHPExcel->getProperties()->setModified($properties['modified']);
	}
	
	public function buildSheet($sheet_index = 0, $sheet = array())
	{
		// go to the sheet by number
		$sheetNames = $this->objPHPExcel->getSheetNames();
		if(!isset($sheetNames[$sheet_index]))
		{
			$this->objPHPExcel->createSheet(NULL, $sheet_index);
		}
		
		$this->objPHPExcel->setActiveSheetIndex($sheet_index);
		
		
		$row_id = 1;
		$column_id = 'A';
		
		$jump = false;
		if(isset($sheet['sheet_title']))
		{
			$this->objPHPExcel->getActiveSheet()->setTitle($sheet['sheet_title']);
		}
		
		$title_location = false;
		if(isset($sheet['title']))
		{
			$this->objPHPExcel->getActiveSheet()->setCellValue($column_id.$row_id, $sheet['title']);
			$size = $this->objPHPExcel->getActiveSheet()->getStyle($column_id.$row_id)->getFont()->getSize();
			$size = $size + 3;
			$this->objPHPExcel->getActiveSheet()->getStyle($column_id.$row_id)->getFont()->setSize($size);
			$this->objPHPExcel->getActiveSheet()->getStyle($column_id.$row_id)->getFont()->setBold(true);
			$jump = true;
			$title_location = $column_id.':'.$row_id;
			$row_id++;
		}
		
		$subtitle_location = false;
		if(isset($sheet['subtitle']))
		{
			$this->objPHPExcel->getActiveSheet()->setCellValue($column_id.$row_id, $sheet['subtitle']);
			$size = $this->objPHPExcel->getActiveSheet()->getStyle($column_id.$row_id)->getFont()->getSize();
			$size = $size + 1;
			$this->objPHPExcel->getActiveSheet()->getStyle($column_id.$row_id)->getFont()->setSize($size);
			$this->objPHPExcel->getActiveSheet()->getStyle($column_id.$row_id)->getFont()->setBold(true);
			$jump = true;
			$subtitle_location = $column_id.':'.$row_id;
			$row_id++;
		}
		
		if($jump) // jump 2 rows
		{
			$row_id++;
			$row_id++;
		}
		
		// if we're dealing with a matrix
		if(isset($sheet['matrix']))
		{
			list($column_id, $row_id) = $this->buildFromMatrix($sheet['matrix'], $column_id, $row_id);
		}
		
		// if we're dealing with basically a csv input
		if(isset($sheet['csv']))
		{
			if(!is_array($sheet['csv']))
			{
				$sheet['csv'] = array($sheet['csv']);
			}
			
			foreach($sheet['csv'] as $csv_content)
			{
				list($column_id, $row_id) = $this->buildFromCsv($csv_content, $column_id, $row_id);
				
				// fix the title and subtitle
				$column_index = range('A', 'Z');
				if($title_location and isset($this->csvColCount))
				{
					$title_location = explode(':', $title_location);
					$col_start = $title_location[0];
					$row = $title_location[1];
					$col_end = $column_index[($this->csvColCount -1)];
					$merge_range = __('%s%s:%s%s', $col_start, $row, $col_end, $row);
					$this->objPHPExcel->getActiveSheet()->mergeCells($merge_range);
				}
				if($subtitle_location and isset($this->csvColCount))
				{
					$subtitle_location = explode(':', $subtitle_location);
					$col_start = $subtitle_location[0];
					$row = $subtitle_location[1];
					$col_end = $column_index[($this->csvColCount -1)];
					$merge_range = __('%s%s:%s%s', $col_start, $row, $col_end, $row);
					
					$this->objPHPExcel->getActiveSheet()->mergeCells($merge_range);
				}
			}
		}
		
		// adjust the widths of the columns
		foreach(range('A','Z') as $columnID)
		{
			$this->objPHPExcel->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
		}
		
		$this->objPHPExcel->getActiveSheet()->getProtection()->setSheet(true);
		$this->objPHPExcel->getActiveSheet()->getProtection()->setInsertRows(true);
		$this->objPHPExcel->getActiveSheet()->getProtection()->setFormatCells(true);
//		$this->objPHPExcel->getActiveSheet()->getProtection()->setPassword( (isset($sheet['title'])?md5($sheet['title']):'This is a Sheet Secret!') );
	}
	
	public function buildFromCsv($csv_data = array(), $column_id = 'A', $row_id = 1)
	{
		if(!$csv_data)
		{
			return array($column_id, $row_id);
		}
		if(is_string($csv_data))
		{
			$csv_data = array_map('str_getcsv', explode("\n", trim($csv_data)));
		}
		
		if(!$csv_data)
		{
			return array($column_id, $row_id);
		}
		$orig_column_id = $column_id;
		$column_index = range('A', 'Z');
		
		$this->csvColCount = 0;
		foreach($csv_data as $row_idx => $row)
		{
			$this->csvColCount = ($this->csvColCount > count($row)?$this->csvColCount:count($row));
			foreach($row as $col_idx => $column)
			{
				$column_id = $column_index[$col_idx];
				$cell_name = $column_id.$row_id;
				$this->objPHPExcel->getActiveSheet()->setCellValue($cell_name, $column);
				$this->objPHPExcel->getActiveSheet()->getStyle("$cell_name:$cell_name")->getAlignment()->setWrapText(true);
			}
			$row_id++;
		}
		
		return array($orig_column_id, $row_id);
	}
	
	public function buildFromMatrix($matrix = array(), $column_id = 'A', $row_id = 1)
	{
		if(!is_array($matrix))
		{
			return array($column_id, $row_id);
		}
		$orig_column_id = $column_id;
		$column_index = range('A', 'Z');
		
		foreach($matrix as $row)
		{
			foreach($row as $col_idx => $column)
			{
				if($col_idx === 'csv')
				{
					list($column_id, $row_id) = $this->buildFromCsv($column, $column_id, $row_id);
				}
				else
				{
					if(!is_array($column))
					{
						$column = array('content' => $column);
					}
					$column_id = $column_index[$col_idx];
					
					$cell_name = $column_id.$row_id;
					$this->objPHPExcel->getActiveSheet()->setCellValue($cell_name, $column['content']);
					if(isset($column['style']))
					{
						$this->objPHPExcel->getActiveSheet()->getStyle("$cell_name:$cell_name")->applyFromArray(Hash::expand($column['style']));
					}
					$this->objPHPExcel->getActiveSheet()->getStyle("$cell_name:$cell_name")->getAlignment()->setWrapText(true);
				}
			}
			$row_id++;
		}
		return array($orig_column_id, $row_id);
	}
	
	public function export($sheets = array(), $properties = array())
	{
		if(!$sheets) return false;
		
		$this->loadPHPExcel(true);
		$this->setProperties($properties);
		
		$sheet_index = 0;
		foreach($sheets as $sheet)
		{
			$this->buildSheet($sheet_index, $sheet);
			$sheet_index++;
		}
		
		$this->objPHPExcel->setActiveSheetIndex(0);
		$this->objPHPExcel->getSecurity()->setLockWindows(true);
		$this->objPHPExcel->getSecurity()->setLockStructure(true);
//		$this->objPHPExcel->getSecurity()->setWorkbookPassword('Um, This Is A Secret!');
		
		$objWriter = new PHPExcel_Writer_Excel2007($this->objPHPExcel);
		ob_start();
		$objWriter->save('php://output');
		$excelOutput = ob_get_clean();
		return $excelOutput;
	}
}