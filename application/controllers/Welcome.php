<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Welcome extends CI_Controller {

	public function index()
	{
		$this->load->view('welcome_message');
	}
	
	public function excel()
	{
		$tmpfname 		= "template.xls";
        $excelReader 	= PHPExcel_IOFactory::createReaderForFile($tmpfname);
        $objPHPExcel 	= $excelReader->load($tmpfname);
        
        // Set document properties
        $objPHPExcel->getProperties()->setCreator("Rakhesh Thayyur")
							 ->setLastModifiedBy("Rakhesh Thayyur")
							 ->setTitle("SampleExcelFile")
							 ->setSubject("SampleExcelFile")
							 ->setDescription("SampleExcelFile")
							 ->setKeywords("SampleExcelFile")
							 ->setCategory("SampleExcelFile");

        // Create a first sheet
        $objPHPExcel->setActiveSheetIndex(0);
        
        // Set Font Color, Font Style and Font Alignment
        $head=array(
            		'borders' 		=> array(
              							'allborders' => array(
                												'style' => PHPExcel_Style_Border::BORDER_THIN,
                												'color' => array('rgb' => '000000')
              												  )
            								),
            		'alignment' 	=> array(
            							'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            							'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        									),
           	 		'fill' 	=> array(
											'type'       => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
	                            			'rotation'   => 90,
	                            			'startcolor' => array(
	                                		'argb' => 'FF007DC3'
	                           		 ),
	                'endcolor'   => array(
	                    			'argb' => 'FFFFFFFF'
	                            )
							)
        );

        
        // Merge Cells
        $objPHPExcel->getActiveSheet()->mergeCells('A1:U1');
        $objPHPExcel->getActiveSheet()->setCellValue('A1', "Salary Statement for the month July- 2020");
        $objPHPExcel->getActiveSheet()->getStyle('A1:U1')->applyFromArray($head);
        $objPHPExcel->getActiveSheet()->getStyle("A1:U1")->getFont()->setSize(20);
        
        //Set Auto
        //$objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(-1); 
        $objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(40);
		//$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(100);
        
        
        // Set Font Color, Font Style and Font Alignment
        $Gradi=array(
            'borders' => array(
              'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => array('rgb' => '000000')
              )
            ),
            'fill' 	=> array(
								'type'       => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
	                            'rotation'   => 90,
	                            'startcolor' => array(
	                                'argb' => 'ff9900'
	                            ),
	                            'endcolor'   => array(
	                                'argb' => 'ffff66'
	                            )
							),
            'alignment' => array(
              'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            )
        );
        
        // Merge Cells
        $objPHPExcel->getActiveSheet()->mergeCells('P2:P3');
        $objPHPExcel->getActiveSheet()->setCellValue('P2', "Salary to Bank");
        $objPHPExcel->getActiveSheet()->getStyle('P2:P3')->applyFromArray($Gradi);
        $objPHPExcel->getActiveSheet()->getStyle('P2:P3')->getFont()->setBold(true);
        
        
        $objPHPExcel->getActiveSheet()->setCellValue('B3', "NORMAL");
        $objPHPExcel->getActiveSheet()->setCellValue('B4', "BOLD");
        $objPHPExcel->getActiveSheet()->getStyle('B4')->getFont()->setBold(true);
        
        $objPHPExcel->getActiveSheet()->setCellValue('B5', "Underline");
        $objPHPExcel->getActiveSheet()->getStyle('B5')->getFont()->setUnderline(true);
        
        $objPHPExcel->getActiveSheet()->setCellValue('B6', "Italic");
        $objPHPExcel->getActiveSheet()->getStyle('B6')->getFont()->setItalic(true);
        
        $objPHPExcel->getActiveSheet()->setCellValue('A10', "Rak");
        $objPHPExcel->getActiveSheet()->setCellValue('B10', "Raj");
        $objPHPExcel->getActiveSheet()->setCellValue('C10', "Sou");
        $objPHPExcel->getActiveSheet()->setCellValue('D10', "Sij");
        $objPHPExcel->getActiveSheet()->setCellValue('E10', "Lak");

        // Hide V and W column
        $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setVisible(false);
        $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setVisible(false);

        // Set auto size
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);

        // Add data
        for ($i = 11; $i <= 16; $i++) 
        {
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $i + 5)
                                        ->setCellValue('B' . $i, $i + 5)
                                        ->setCellValue('C' . $i, $i + 5)
                                        ->setCellValue('D' . $i, $i + 5)
                                        ->setCellValue('E' . $i, $i + 5)
                                        ->setCellValue('F' . $i, '=SUM(A'.$i.':E'.$i.')');
        }
        

        //$objPHPExcel->getActiveSheet()->getStyle('A3:E3')->applyFromArray($stil);

        // Merge Cells
        //$objPHPExcel->getActiveSheet()->mergeCells('A5:E5');
        //$objPHPExcel->getActiveSheet()->setCellValue('A5', "MERGED CELL");
        //$objPHPExcel->getActiveSheet()->getStyle('A5:E5')->applyFromArray($stil);
        
        // Save Excel xls File >> Excel5
        /*$filename="SampleExcelFile.xls";
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        ob_end_clean();
        header('Content-type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename='.$filename);
        $objWriter->save('php://output');*/
        
        // Save Excel xls File >> Excel2007
        $filename="SampleExcelFile.xlsx";
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        ob_end_clean();
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Disposition: attachment;filename="userList.xlsx"');
        $objWriter->save('php://output');
	}
}
