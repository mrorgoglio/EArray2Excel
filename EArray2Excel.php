<?php

	/**
	* @author Salvo Di Mare
	* @license GNU v2
	* @version 0.1
	*
	* Inspired by Nikola Kostadinov's EExcelView extension
	*/
	
	class EArray2Excel
	{
		//Document properties
		public $creator = 'Salvo Di Mare';
		public $title = '';
		public $subject = '';
		public $description = '';
		public $category = '';

		//the PHPExcel object
		public $objPHPExcel = null;
		public $libPath = 'ext.phpexcel.Classes.PHPExcel'; //the path to the PHP excel lib

		//config
		public $exportType = 'Excel5';
		public $filename = null; //export FileName
		public $stream = true; //stream to browser
		
		//sheets
		public $sheets = null;
		
		//mime types used for streaming
		public $mimeTypes = array(
			'Excel5'	=> array(
				'Content-type'=>'application/vnd.ms-excel',
				'extension'=>'xls',
				'caption'=>'Excel(*.xls)',
			),
			'Excel2007'	=> array(
				'Content-type'=>'application/vnd.ms-excel',
				'extension'=>'xlsx',
				'caption'=>'Excel(*.xlsx)',				
			),
		);

		/**
		 * Initialization 
		 * 
		 * Loads PHPExcel library and creates a workbook
		 * 
		 * @param string $title 
		 * 					the document title
		 * @param string $creator 
		 * 					the document creator
		 * @param string $subject 
		 * 					the document subject
		 * @param string $description 
		 * 					the document description
		 * @param string $category 
		 * 					the document category
		 * 
		 */
		public function init($title=null, $creator=null, $subject=null, $description=null, $category=null)
		{
			//Autoload fix
			spl_autoload_unregister(array('YiiBase','autoload'));             
			Yii::import($this->libPath, true);
			$this->objPHPExcel = new PHPExcel();
			spl_autoload_register(array('YiiBase','autoload'));  
			//Creating a workbook
			if ($creator)
			$this->objPHPExcel->getProperties()->setCreator($creator?$creator:$this->creator);
			$this->objPHPExcel->getProperties()->setTitle($title?$title:$this->title);
			$this->objPHPExcel->getProperties()->setSubject($subject?$subject:$this->subject);
			$this->objPHPExcel->getProperties()->setDescription($description?$description:$this->description);
			$this->objPHPExcel->getProperties()->setCategory($category?$category:$this->category);
		}
		
		/**
		 * Constructor
		 * 
		 * Creates new obj
		 * 
		 * @param string $title 
		 * 					the document title
		 * @param string $creator 
		 * 					the document creator
		 * @param string $subject 
		 * 					the document subject
		 * @param string $description 
		 * 					the document description
		 * @param string $category 
		 * 					the document category
		 * 
		 */
		public function EArray2Excel($title=null, $creator=null, $subject=null, $description=null, $category=null)
		{
			//Init
			$this->init($title, $creator, $subject, $description, $category);
		}
		
		/**
		 * Sets some export settings and calls the render function
		 * 
		 * @param array $sheets
		 * 				 associative array that contains data to be exported; it is in the form:
		 * 				{'Sheet name' => {'ColumnName'=>'ColumnValue',..},...}
		 * @param string $filename
		 * 				filename in which to store the generated excel or to be used to set http header content-disposition
		 * @param bool $stream
		 * 				indicates if stream content to the browser(true) or save content to file(false) indicated by filename
		 * @param string $export_type
		 * 				indicates the Export Type; options availables:
		 * 					Excel5
		 * 					Excel2007						
		 */
		public function export($sheets,$filename=null,$stream=true,$export_type=null)
		{
			//Setting export information
			$this->exportType = $export_type?$export_type:$this->exportType;
			$this->stream = $stream;
			$this->filename=$filename;
			$this->sheets = $sheets;
			//render excel document
			$this->render();
		}
		
		
		/**
		 * Generates Excel document
		 * 
		 */
		public function render()
		{
			if (is_array($this->sheets))
			{
				$sheet_index = 0;
				//iterate over sheets
				foreach($this->sheets as $sheet_name=>$sheet_data)
				{
					//render a single sheet
					if ($this->renderSheet($sheet_index,$sheet_name,$sheet_data))
						$sheet_index++;
				}
				//create writer for saving
				$objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, $this->exportType);
				if(!$this->stream)
					$objWriter->save($this->filename);
				else //output to browser
				{
					if(!$this->filename)
						$this->filename = $this->title;
					//clean output
					$this->cleanOutput();
					//set http headers
					header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
					header('Pragma: public');
					header('Content-type: '.$this->mimeTypes[$this->exportType]['Content-type']);
					header('Content-Disposition: attachment; filename="'.$this->filename.'.'.$this->mimeTypes[$this->exportType]['extension'].'"');
					header('Cache-Control: max-age=0');
					//send content				
					$objWriter->save('php://output');			
					Yii::app()->end();
				}
			}
		}
		
		
		/**
		 * Generates a single sheet
		 * 
		 */
		public function renderSheet($index, $name, $data)
		{
			if (is_array($data) && !empty($data))
			{
				//Create new sheet
				$this->objPHPExcel->createSheet(NULL, $index);
				$this->objPHPExcel->setActiveSheetIndex($index);
				$this->objPHPExcel->getActiveSheet()->setTitle($name);
				
				//write headers
				$headers = array_keys($data[0]);
				$a=0;
				foreach ($headers as $h) {
					$a=$a+1;
					$cell = $this->objPHPExcel->getActiveSheet()->setCellValue($this->columnName($a)."1" ,ucfirst($h), true);
					if ($a==1)
						$this->objPHPExcel->getActiveSheet()->getStyle($this->columnName($a)."1")->getFont()->setBold(true);
				}
					
				//write body
				$row_index=3;
				foreach ($data as $row) {
					$a=0;
					foreach ($headers as $h) {
						$a=$a+1;
						$cell = $this->objPHPExcel->getActiveSheet()->setCellValue($this->columnName($a).$row_index ,$row[$h], true);
					}
					$row_index++;
				}
				return true;
			}
			return false;
		}

		/**
		* Returns the coresponding excel column.(Abdul Rehman from yii forum)
		* 
		* @param int $index
		* @return string
		*/
		public function columnName($index)
		{
			--$index;
			if($index >= 0 && $index < 26)
				return chr(ord('A') + $index);
			else if ($index > 25)
				return ($this->columnName($index / 26)).($this->columnName($index%26 + 1));
				else
					throw new Exception("Invalid Column # ".($index + 1));
		}		
		
		/**
		* Performs cleaning on mutliple levels.
		* 
		* From le_top @ yiiframework.com
		* 
		*/
		private static function cleanOutput() 
		{
            for($level=ob_get_level();$level>0;--$level)
            {
                @ob_end_clean();
            }
        }		


	}
