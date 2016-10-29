<?php
/****************
 ExcelTemplateWriter: Write data to .xlsx files.
 
 Dependencies:
 	ZipArchive <http://php.net/manual/en/class.ziparchive.php>
	DOMDocument <http://php.net/manual/en/class.domdocument.php>
	Exception <http://php.net/manual/en/language.exceptions.php>
	RecursiveDirectoryIterator <http://php.net/manual/en/class.recursivedirectoryiterator.php>
	RecursiveIteratorIterator <http://php.net/manual/en/class.recursiveiteratoriterator.php>
 
 Public Methods:
	__construct:
		$source_path: path to xlsx file
		$options: unused, reserved for future features
	
	selectSheet:
		$name: string name of the sheet
		
	writeToCell:
		$column: letter or position (1+)
		$row: position (1+)
		$data: string, or number
		$type: default is auto detect (string or number), avaliable types are:
			"string"
			"number"
		
	fillColumn:
		$column: see writeToCell
		$starting_row: position (1+) of the row to start with. The row will be incremented to write all the data
		$data: see writeToCell
		$type: see writeToCell
		
	save:
		$destination_path: path for the new xlsx file

 Example:
	require('includes/ExcelTemplateWriter.php');

	$writer = new ExcelTemplateWriter('sheets/monthly_report_template.xlsx');

	$writer->selectSheet("Demographics");
	$writer->writeToCell('B', 2, 500);
	$writer->writeToCell('B', 3, 200);

	$dates = array('1/10', '2/10', '3/10', '4/10', '5/10');
	$reach = array(1, 2, 3, 4, 5);
	$comments = array(6, 7, 8, 9, 10);
	$likes = array(11, 12, 13, 14, 15);
	$shares = array(16, 17, 18, 19, 20);
	$writer->selectSheet("Engagement-Reach per post");
	$writer->fillColumn('E', 2, $dates);
	$writer->fillColumn('F', 2, $reach);
	$writer->fillColumn('G', 2, $comments);
	$writer->fillColumn('H', 2, $likes);
	$writer->fillColumn('I', 2, $shares);

	$writer->save('sheets/monthly_filled.xlsx');
	
****************/
//namespace if needed
use ZipArchive;
use DOMDocument;
use ArrayObject;
use RecursiveIteratorIterator;
use RecursiveDirectoryIterator;
use Exception;

class ExcelTemplateWriter {
	private $source_path;
	private $options = array(
	);
	private $extracted_directory;
	private $sheet_name_map = array();
	private $selected_sheet = 1;
	private $changes = array();
	
	public function __construct($source_path, $options = array()) {
		$this->source_path = realpath($source_path);
		$this->options = array_replace($this->options, $options);
		
		$this->extractXLSX();
		$this->fetchSheetNames();
		$this->deleteExtractedFolder();
	}
	
	private function extractXLSX() {
		if (!$this->extracted_directory) {
			$path_parts = pathinfo($this->source_path);
			$base_directory = $path_parts['dirname'];
			do {
				$temp_directory = md5($this->source_path . mt_rand() . microtime());
			} while(file_exists($base_directory . "/" . $temp_directory));
			$this->extracted_directory = $base_directory . "/" . $temp_directory;
			$zip = new ZipArchive;
			if ($zip->open($this->source_path) === true) {
				$zip->extractTo($this->extracted_directory);
				$zip->close();
			} else {
				throw new Exception('Cannot extract ' . $this->source_path . ' to ' . $this->extracted_directory);
			}
		}
	}
	
	private function fetchSheetNames() {
		$dom = new DOMDocument();
		$dom->load($this->extracted_directory . '/xl/workbook.xml');
		$root = $dom->documentElement;
		$sheets = $root->getElementsByTagName("sheet");
		for ($i=0;$i<$sheets->length;$i++) {
			$sheet = $sheets->item($i);
			$name = $sheet->getAttribute('name');
			$this->sheet_name_map[trim(strtolower($name))] = $i+1;
		}
	}
	
	private function deleteExtractedFolder() {
		if ($this->extracted_directory) {
			$path = $this->extracted_directory;
			$this->extracted_directory = null;
			if (is_dir($path) === true) {
				$files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($path), RecursiveIteratorIterator::CHILD_FIRST);
				foreach ($files as $file) {
					if (in_array($file->getBasename(), array('.', '..')) !== true) {
						if ($file->isDir() === true) {
							rmdir($file->getPathName());
						} elseif (($file->isFile() === true) || ($file->isLink() === true)) {
							unlink($file->getPathname());
						}
					}
				}
				return rmdir($path);
			} elseif ((is_file($path) === true) || (is_link($path) === true)) {
				return unlink($path);
			}
		}
		return false;
	}
	
	public function selectSheet($name) {
		$this->selected_sheet = $this->sheetPositionFromName($name);
	}
	
	private function sheetPositionFromName($name){
		$name_lower = strtolower(trim($name));
		if (!array_key_exists($name_lower, $this->sheet_name_map)){
			throw new Exception('Cannot find a sheet with the name ' . $name);
		}
		return $this->sheet_name_map[$name_lower];
	}
	
	private function columnNumberToLetters($num) {
		$numeric = ($num - 1) % 26;
		$letter = chr(65 + $numeric);
		$num2 = intval(($num - 1) / 26);
		if ($num2 > 0) {
			return $this->columnNumberToLetters($num2) . $letter;
		} else {
			return $letter;
			}
	}
	
	public function writeToCell($column, $row, $data, $type = null) {
		if (!is_string($column)) {
			if ($column < 1) {
				throw new Exception('Column cannot be less than 1');
			}
			$column = $this->columnNumberToLetters($column);
		}
		if ($row < 1) {
			throw new Exception('Row cannot be less than 1');
		}
                                
		if ($type === null) {
			if (is_string($data)) {
				$type = "string";
			} else {
				$type = "number";
			}
		}
		
		if (!array_key_exists($this->selected_sheet, $this->changes)) {
			$this->changes[$this->selected_sheet] = array();
		}
		if (!array_key_exists($row, $this->changes[$this->selected_sheet])){
			$this->changes[$this->selected_sheet][$row] = array();
		}
		
		$this->changes[$this->selected_sheet][$row][$column] = array('data' => $data, 'type' => $type);
		
		//sorting changes array on insert, optimize for a faster method?
		$that = $this;
		uksort ($this->changes[$this->selected_sheet][$row], function($a, $b) use (&$that){
			return $that->compareColumnLetters($a, $b);
		});
		uksort ($this->changes[$this->selected_sheet], function ($a, $b){
			return $a>$b;
		});
	}
	
	public function compareColumnLetters($a, $b) {
		$slen_a = strlen($a);
		$slen_b = strlen($b);
		if ($slen_a == $slen_b){
			if ($a == $b){
				return 0;
			}
			if ($a > $b){
				return 1;
			} else {
				return -1;
			}
		} else {
			if ($slen_a > $slen_b){				
				return 1;
			} else {
				return -1;
			}
		}
	}
	
	public function fillColumn($column, $starting_row, $data, $type = null) {
		$current_row = $starting_row;
		foreach ($data as $key => $value) {
			$this->writeToCell($column, $current_row, $value, $type);
			$current_row++;
		}
	}
	
	private function __echoNode($node){
		$newdoc = new DOMDocument;
		$node = $newdoc->importNode($node, true);
		$newdoc->appendChild($node);
		echo $newdoc->saveHTML();
	}
	
	private function getRowInfo($rows, $current_index) {
		if (($rows->length > 0) && ($rows->length > $current_index)) {
			$row = $rows->item($current_index);
			$row_number = intval($row->getAttribute('r'));
			if ($row_number < 1){
				$row = null;
				$row_number = null;
			}
		} else {
			$row = null;
			$row_number = null;
		}
		return array('row' => $row, 'row_number' => $row_number);
	}
	
	private function createRow($dom, $row_number) {
		$row = $dom->createElement('row');
		$row->setAttribute('r', $row_number);
		return $row;
	}
	
	private function processCells($dom, $cell_changes, $row, $row_number) {
		$cells = $row->getElementsByTagName('c');
																
		$current_cell_index = 0;
		$cell_info = $this->getCellInfo($row_number, $cells, $current_cell_index);
		$this->fixCellFormula($cell_info['cell']);
		if ($cell_changes) {
			$cell_changes_iterator = new ArrayObject($cell_changes);
			$cell_iterator = $cell_changes_iterator->getIterator();
			while ($cell_iterator->valid()) {
				$column_number = $cell_iterator->key();
				$cell_change = $cell_iterator->current();
				if ($cell_info['column_number'] == null) {
					$new_cell = $this->createCell($dom, $row_number, $column_number, $cell_change);
					$row->appendChild($new_cell);
					$cell_iterator->next();
				} else {
					$comparison = $this->compareColumnLetters($column_number, $cell_info['column_number']);
					if ($comparison == -1){
						$new_cell = $this->createCell($dom, $row_number, $column_number, $cell_change);
						$row->insertBefore($new_cell, $cell_info['cell']);
						$cell_iterator->next();
					} elseif ($comparison == 0){
						$this->applyCellValue($dom, $cell_info['cell'], $cell_change);
						$cell_iterator->next();
					} else {
						$this->fixCellFormula($cell_info['cell']);
						$current_cell_index++;
						$cell_info = $this->getCellInfo($row_number, $cells, $current_cell_index);
					}
				}
			}
		}
		
		while ($cell_info['cell'] !== null) {
			$this->fixCellFormula($cell_info['cell']);
			$current_cell_index++;
			$cell_info = $this->getCellInfo($row_number, $cells, $current_cell_index);
		}
	}
	
	private function getCellInfo($row_number, $cells, $current_index) {
		if (($cells->length > 0) && ($cells->length > $current_index)) {
			$cell = $cells->item($current_index);
			$cell_name = $cell->getAttribute('r');
			$column_number = explode($row_number, $cell_name);
			$column_number = $column_number[0];
		} else {
			$cell = null;
			$column_number = null;
		}
		return array('cell' => $cell, 'column_number' => $column_number);
	}
	
	private function createCell($dom, $row_number, $column, $change) {
		$cell = $dom->createElement('c');
		$cell->setAttribute('r', $column . $row_number);
		$value = $dom->createElement('v');
		$cell->appendChild($value);
		$this->applyCellValue($dom, $cell, $change);
		return $cell;
	}
	
	private function applyCellValue($dom, $cell, $change) {
		if ($change){
			$parent_of_value = $cell;
			
			$current_type = $cell->getAttribute('t');
			$inline_string = null;
			if ($current_type == 's'){
				$cell->removeAttribute('s');
				//delete shared string?
			} elseif ($current_type == 'inlineStr') {
				$inline_string = $cell->getElementsByTagName('is');
				if ($inline_string->length > 0){
					$cell->removeChild($inline_string);
				}
			}
			
			$value_wrapper_name = 'v';
			if ($change['type'] == 'number'){
				$cell->setAttribute('t', 'n');
			} else {
				$cell->setAttribute('t', 'inlineStr');
				if (!$inline_string){
					$inline_string = $dom->createElement('is');
				}
				$cell->appendChild($inline_string);
				$parent_of_value = $inline_string;
				$value_wrapper_name = 't';
			}
			
			$value = $parent_of_value->getElementsByTagName($value_wrapper_name);
			if ($value->length == 0){
				$value = $dom->createElement($value_wrapper_name);
				$parent_of_value->appendChild($value);
			} else {
				$value = $value->item(0);
			}
			
			$value->nodeValue = htmlspecialchars($change['data']);
			//unset value from the changes object to save memory?
		}
	}
	
	private function fixCellFormula($cell) {
		if ($cell){
			//clear formula default value
			$formula = $cell->getElementsByTagName('f');
			if ($formula->length > 0){
				$value = $cell->getElementsByTagName('v');
				if ($value->length > 0){
					$cell->removeChild($value->item(0));
				}
			}
		}
	}
	
	public function save($destination_path) {
		$this->extractXLSX();

		foreach ($this->sheet_name_map as $name => $id) {
			$dom = new DOMDocument();
			$dom->substituteEntities = true;
			$dom->load($this->extracted_directory . "/xl/worksheets/sheet$id.xml");
			$root = $dom->documentElement;
			$sheetData = $root->getElementsByTagName("sheetData");
			if ($sheetData->length == 1) {
				$sheetData = $sheetData->item(0);
				$rows = $sheetData->getElementsByTagName("row");
				
				$current_row_index = 0;
				$row_info = $this->getRowInfo($rows, $current_row_index);
				$row_changes = @$this->changes[$id];
				if (count($row_changes) > 0){
					$row_changes_iterator = new ArrayObject($row_changes);
					$row_iterator = $row_changes_iterator->getIterator();
					while ($row_iterator->valid()) {
						$row_number = $row_iterator->key();
						$row_change = $row_iterator->current();
						if ($row_info['row_number'] == null) {
							$new_row = $this->createRow($dom, $row_number);
							$this->processCells($dom, $row_change, $new_row, $row_number);
							$sheetData->appendChild($new_row);
							$row_iterator->next();
						} elseif ($row_number < $row_info['row_number']){
							$new_row = $this->createRow($dom, $row_number);
							$this->processCells($dom, $row_change, $new_row, $row_number);
							$sheetData->insertBefore($new_row, $row_info['row']);
							$row_iterator->next();
						} elseif ($row_number == $row_info['row_number']){
							$this->processCells($dom, $row_change, $row_info['row'], $row_info['row_number']);
							$row_iterator->next();
						} else {
							$this->processCells($dom, null, $row_info['row'], $row_info['row_number']);
							$current_row_index++;
							$row_info = $this->getRowInfo($rows, $current_row_index);
						}
					}
				}
				while ($row_info['row'] !== null) {
					$this->processCells($dom, null, $row_info['row'], $row_info['row_number']);
					$current_row_index++;
					$row_info = $this->getRowInfo($rows, $current_row_index);
				}
			}
			$dom->save($this->extracted_directory . "/xl/worksheets/sheet$id.xml", LIBXML_NOEMPTYTAG);
		}
		$root_path = $this->extracted_directory;
		//create zip file
		$zip = new ZipArchive();
		
		$zip->open($destination_path, ZipArchive::CREATE | ZipArchive::OVERWRITE);
		$files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($root_path), RecursiveIteratorIterator::LEAVES_ONLY);

		foreach ($files as $name => $file) {
			//skip directories (they would be added automatically)
			if (!$file->isDir()) {
				$file_path = $file->getRealPath();
				$relative_path = substr($file_path, strlen($root_path) + 1);
				$zip->addFile($file_path, $relative_path);
			}
		}
		$zip->close();
		$this->deleteExtractedFolder();
	}
}
?>