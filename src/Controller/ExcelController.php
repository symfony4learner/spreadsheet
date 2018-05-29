<?php

namespace App\Controller;

use Symfony\Component\Routing\Annotation\Route;
use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\Config\FileLocator;

class ExcelController extends Controller
{
    /**
     * @Route("/excel", name="excel")
     */
    public function index()
    {
        return $this->render('excel/index.html.twig', [
            'controller_name' => 'ExcelController',
        ]);
    }

    /**
     * @Route("/excel/xlsx/write", name="write")
     */
    public function write() 
    {
    	$data = [];
		$spreadsheet = $this->get('phpoffice.spreadsheet')->createSpreadsheet();
		$spreadsheet->getActiveSheet()->setCellValue('A1', 'Hello world');
		$spreadsheet->getActiveSheet()->setCellValue('B1', 'Hello again');
        $new_file = $this->get('kernel')->getProjectDir()."/public/new.xlsx";
		$writerXlsx = $this->get('phpoffice.spreadsheet')->createWriter($spreadsheet, 'Xlsx');
		// from array
		$arrayData = [
		    [NULL, 2010, 2011, 2012],
		    ['Q1',   12,   15,   21],
		    ['Q2',   56,   73,   86],
		    ['Q3',   52,   61,   69],
		    ['Q4',   30,   32,    0],
		];
		$spreadsheet->getActiveSheet()
		    ->fromArray(
		        $arrayData,  // The data to set
		        NULL,        // Array values with this value will not be set
		        'C3'         // Top left coordinate of the worksheet range where
		                     //    we want to set these values (default is A1)
		    );
		// $spreadsheet->getActiveSheet()->getCell('E11')->getCalculatedValue();

		$writerXlsx->save($new_file);
        return $this->render('excel/write.html.twig', $data);
    }

    /**
     * @Route("/excel/xlsx/read", name="read")
     */
    public function read() 
    {
    	$data = [];
    	// path to file
        $sample_file = $this->get('kernel')->getProjectDir()."/public/new.xlsx";
        // create a reader
		$readerXlsx  = $this->get('phpoffice.spreadsheet')->createReader('Xlsx');
		// load the spreadsheet to a variable
		$spreadsheet = $readerXlsx->load($sample_file);
		// take data from spreadsheet and make array with it
		$dataArray = $spreadsheet->getActiveSheet()
		    ->rangeToArray(
		        'C3:F5',     // The worksheet range that we want to retrieve
		        NULL,        // Value that should be returned for empty cells
		        TRUE,        // Should formulas be calculated (the equivalent of getCalculatedValue() for each cell)
		        TRUE,        // Should values be formatted (the equivalent of getFormattedValue() for each cell)
		        TRUE         // Should the array be indexed by cell row and cell column
		    );

		// iterate through, get data and make table
		$readerXlsx->setReadDataOnly(TRUE);
		$worksheet = $spreadsheet->getActiveSheet();

		$table = '<table>' . PHP_EOL;
		foreach ($worksheet->getRowIterator() as $row) {
		    $table .= '<tr>' . PHP_EOL;
		    $cellIterator = $row->getCellIterator();
		    $cellIterator->setIterateOnlyExistingCells(FALSE); // This loops through all cells,
		                                                       //    even if a cell value is not set.
		                                                       // By default, only cells that have a value
		                                                       //    set will be iterated.
		    foreach ($cellIterator as $cell) {
		        $table .= '<td>' .
		             $cell->getValue() .
		             '</td>' . PHP_EOL;
		    }
		    $table .= '</tr>' . PHP_EOL;
		}
		$table .= '</table>' . PHP_EOL;		  
        $data['sample_file'] = $sample_file;
        $data['dataArray'] = $dataArray;
        $data['table'] = $table;
        return $this->render('excel/read.html.twig', $data);
    }
    
	// $spreadsheet->getProperties()
	//     ->setCreator("Maarten Balliauw")
	//     ->setLastModifiedBy("Maarten Balliauw")
	//     ->setTitle("Office 2007 XLSX Test Document")
	//     ->setSubject("Office 2007 XLSX Test Document")
	//     ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
	//     ->setKeywords("office 2007 openxml php")
	//     ->setCategory("Test result file");

}
