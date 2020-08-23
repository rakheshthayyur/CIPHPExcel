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

	public function mynakam()
	{
		$tmpfname 		= "template.xls";
        $excelReader 	= PHPExcel_IOFactory::createReaderForFile($tmpfname);
        $objPHPExcel 	= $excelReader->load($tmpfname);
        
        //Set document properties
	    $objPHPExcel->getProperties()->setCreator("Rakhesh Thayyur")
							 ->setLastModifiedBy("Raji Das")
							 ->setTitle("Salary Statement")
							 ->setSubject("Mynakam Salary Statement")
							 ->setDescription("Mynakam Salary Statement For All Branches")
							 ->setKeywords("Mynakam Salary Statement All Branches")
							 ->setCategory("Salary Statement");

        //Create a first sheet
        $objPHPExcel->setActiveSheetIndex(0);
        
        //Set Font Color, Font Style and Font Alignment
        $Title=array(
            		'borders' 	=> array(
            							'allborders' => array(
            													'style' => PHPExcel_Style_Border::BORDER_THIN,
                												'color' => array(
                																	'rgb' => '000000'
                																)
              												  )
            							),
            		'alignment' => array(
            							'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            							'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        								),
           	 		'fill' 		=> array(
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
        $objPHPExcel->getActiveSheet()->getStyle('A1:U1')->applyFromArray($Title);
        $objPHPExcel->getActiveSheet()->getStyle("A1:U1")->getFont()->setSize(20);
        $objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(50);
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setUnderline(true);
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setItalic(true);
        $objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
        

        $Head=array(
            		'borders' 	=> array(
            							'allborders' => array(
            													'style' => PHPExcel_Style_Border::BORDER_THIN,
                												'color' => array(
                																	'rgb' => '000000'
                																)
              												  )
            							),
            		'alignment' => array(
            							'horizontal'	=> PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            							'vertical' 		=> PHPExcel_Style_Alignment::VERTICAL_CENTER,
            							'wrap'      	=> TRUE
        								)
        			);
        
        $objPHPExcel->getActiveSheet()->mergeCells('A2:A3');
        $objPHPExcel->getActiveSheet()->setCellValue('A2', "SL No.");
        $objPHPExcel->getActiveSheet()->getStyle('A2:A3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('A2:A3')->applyFromArray($Head);
        
        $objPHPExcel->getActiveSheet()->mergeCells('B2:B3');
        $objPHPExcel->getActiveSheet()->setCellValue('B2', "Name");
        $objPHPExcel->getActiveSheet()->getStyle('B2:B3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('B2:B3')->applyFromArray($Head);        

        $objPHPExcel->getActiveSheet()->mergeCells('C2:C3');
        $objPHPExcel->getActiveSheet()->setCellValue('C2', "Actual");
        $objPHPExcel->getActiveSheet()->getStyle('C2:C3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('C2:C3')->applyFromArray($Head); 

        $objPHPExcel->getActiveSheet()->mergeCells('D2:D3');
        $objPHPExcel->getActiveSheet()->setCellValue('D2', "Basic");
        $objPHPExcel->getActiveSheet()->getStyle('D2:D3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('D2:D3')->applyFromArray($Head); 

        $objPHPExcel->getActiveSheet()->mergeCells('E2:E3');
        $objPHPExcel->getActiveSheet()->setCellValue('E2', "DA");
        $objPHPExcel->getActiveSheet()->getStyle('E2:E3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('E2:E3')->applyFromArray($Head); 
        
        $objPHPExcel->getActiveSheet()->mergeCells('F2:F3');
        $objPHPExcel->getActiveSheet()->setCellValue('F2', "Other");
        $objPHPExcel->getActiveSheet()->getStyle('F2:F3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('F2:F3')->applyFromArray($Head);        

        $objPHPExcel->getActiveSheet()->mergeCells('G2:G3');
        $objPHPExcel->getActiveSheet()->setCellValue('G2', "Interim Relief");
        $objPHPExcel->getActiveSheet()->getStyle('G2:G3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('G2:G3')->applyFromArray($Head); 

        $objPHPExcel->getActiveSheet()->mergeCells('H2:H3');
        $objPHPExcel->getActiveSheet()->setCellValue('H2', "Gross Salary");
        $objPHPExcel->getActiveSheet()->getStyle('H2:H3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('H2:H3')->applyFromArray($Head); 

        $objPHPExcel->getActiveSheet()->mergeCells('I2:I3');
        $objPHPExcel->getActiveSheet()->setCellValue('I2', "Total PF Salary");
        $objPHPExcel->getActiveSheet()->getStyle('I2:I3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('I2:I3')->applyFromArray($Head); 

        $objPHPExcel->getActiveSheet()->mergeCells('J2:O2');
        $objPHPExcel->getActiveSheet()->setCellValue('J2', "Employee Deductions");
        $objPHPExcel->getActiveSheet()->getStyle('J2:O2')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('J2:O2')->applyFromArray($Head); 

        $objPHPExcel->getActiveSheet()->setCellValue('J3', "PF");
        $objPHPExcel->getActiveSheet()->getStyle('J3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('J3')->applyFromArray($Head); 
    
        $objPHPExcel->getActiveSheet()->setCellValue('K3', "ESI JUL");
        $objPHPExcel->getActiveSheet()->getStyle('K3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('K3')->applyFromArray($Head);    
    
        $objPHPExcel->getActiveSheet()->setCellValue('L3', "LWF");
        $objPHPExcel->getActiveSheet()->getStyle('L3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('L3')->applyFromArray($Head);      
    
        $objPHPExcel->getActiveSheet()->setCellValue('M3', "ESI MAY");
        $objPHPExcel->getActiveSheet()->getStyle('M3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('M3')->applyFromArray($Head);     

        $objPHPExcel->getActiveSheet()->setCellValue('N3', "Others");
        $objPHPExcel->getActiveSheet()->getStyle('N3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('N3')->applyFromArray($Head); 

        $objPHPExcel->getActiveSheet()->setCellValue('O3', "Loans");
        $objPHPExcel->getActiveSheet()->getStyle('O3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('O3')->applyFromArray($Head); 
        
        // Set Font Color, Font Style and Font Alignment
        $Gradient=array(
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
        
        $objPHPExcel->getActiveSheet()->mergeCells('P2:P3');
        $objPHPExcel->getActiveSheet()->setCellValue('P2', "Salary to Bank");
        $objPHPExcel->getActiveSheet()->getStyle('P2:P3')->applyFromArray($Gradient);
        $objPHPExcel->getActiveSheet()->getStyle('P2:P3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('P2:P3')->applyFromArray($Head);

        $objPHPExcel->getActiveSheet()->mergeCells('Q2:T2');
        $objPHPExcel->getActiveSheet()->setCellValue('Q2', "Employer Contributions");
        $objPHPExcel->getActiveSheet()->getStyle('Q2:T2')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('Q2:T2')->applyFromArray($Head); 

        $objPHPExcel->getActiveSheet()->setCellValue('Q3', "PF");
        $objPHPExcel->getActiveSheet()->getStyle('Q3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('Q3')->applyFromArray($Head); 

        $objPHPExcel->getActiveSheet()->setCellValue('R3', "Admin");
        $objPHPExcel->getActiveSheet()->getStyle('R3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('R3')->applyFromArray($Head); 
        
        $objPHPExcel->getActiveSheet()->setCellValue('S3', "ESI JUL");
        $objPHPExcel->getActiveSheet()->getStyle('S3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('S3')->applyFromArray($Head);         
        
        $objPHPExcel->getActiveSheet()->setCellValue('T3', "LWF");
        $objPHPExcel->getActiveSheet()->getStyle('T3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('T3')->applyFromArray($Head);         

        $objPHPExcel->getActiveSheet()->mergeCells('U2:U3');
        $objPHPExcel->getActiveSheet()->setCellValue('U2', "CTC");
        $objPHPExcel->getActiveSheet()->getStyle('U2:U3')->applyFromArray($Gradient);
        $objPHPExcel->getActiveSheet()->getStyle('U2:U3')->getFont()->setBold(true);
        $objPHPExcel->getActiveSheet()->getStyle('U2:U3')->applyFromArray($Head);


        $Detail=array(
            		'borders' 	=> array(
            							'allborders' => array(
            													'style' => PHPExcel_Style_Border::BORDER_DOTTED,
                												'color' => array(
                																	'rgb' => '000000'
                																)
              												  )
            							)
        			);

        for ($i = 4; $i <= 13; $i++) 
        {
            $objPHPExcel->getActiveSheet()->setCellValue('A' . $i, $i - 3)
                                        ->setCellValue('B' . $i, 'Rakhesh Thayyur')
                                        ->setCellValue('C' . $i, '12800.00')
                                        ->setCellValue('D' . $i, '8910.00')
                                        ->setCellValue('E' . $i, '910.00')
                                        ->setCellValue('F' . $i, '2680.00')
                                        ->setCellValue('G' . $i, '300.00')
                                        ->setCellValue('H' . $i, '=SUM(D'.$i.'+E'.$i.'+F'.$i.'+G'.$i.')')
                                        ->setCellValue('I' . $i, '=IF((D'.$i.'+E'.$i.'+F'.$i.')>15000,0,D'.$i.'+E'.$i.')')
                                        ->setCellValue('J' . $i, '=ROUND(I'.$i.'*12%,0)')
                                        ->setCellValue('K' . $i, '=ROUNDUP(H'.$i.'*0.75%,0)')
                                        ->setCellValue('L' . $i, '20.00')
                                        ->setCellValue('M' . $i, '')
                                        ->setCellValue('N' . $i, '')
                                        ->setCellValue('O' . $i, '1379.00')
                                        ->setCellValue('P' . $i, '=H'.$i.'-J'.$i.'-K'.$i.'-O'.$i.'-L'.$i.'+M'.$i.'+N'.$i.'')
                                        ->setCellValue('Q' . $i, '=ROUND(I'.$i.'*12%,0)')
                                        ->setCellValue('R' . $i, '=ROUND(I'.$i.'*1%,0)')
                                        ->setCellValue('S' . $i, '=ROUNDUP(H'.$i.'*3.25%,0)')
                                        ->setCellValue('T' . $i, '20.00')
                                        ->setCellValue('U' . $i, '=H'.$i.'+Q'.$i.'+R'.$i.'+S'.$i.'+T'.$i.'');
        }

		$objPHPExcel->getActiveSheet()->getStyle('A4:U13')->applyFromArray($Detail);

		$objPHPExcel->getActiveSheet()->getStyle("U4:U13")->getNumberFormat()->setFormatCode('0.00');


        //Set Auto
        //$objPHPExcel->getActiveSheet()->getRowDimension('1')->setRowHeight(-1); 
        $objPHPExcel->getActiveSheet()->getRowDimension('15')->setRowHeight(30);
		//$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(100);
        

        // Hide V and W column
        $objPHPExcel->getActiveSheet()->getColumnDimension('V')->setVisible(false);
        $objPHPExcel->getActiveSheet()->getColumnDimension('W')->setVisible(false);

        // Set auto size
        $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
        $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        

        
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
		header('Content-Disposition: attachment;filename='.$filename);
        $objWriter->save('php://output');
        
	}
}
