<?php

namespace App\Controller;

use Symfony\Component\Routing\Annotation\Route;
use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\BinaryFileResponse;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

class ExportController extends Controller
{
    /**
     * @Route("/export/{file_name}", name="export")
     */
    public function export($file_name)
    {

		$spreadsheet = $this->get('phpoffice.spreadsheet')->createSpreadsheet();
		$writerXlsx = $this->get('phpoffice.spreadsheet')->createWriter($spreadsheet, 'Xlsx');
        $new_file = $this->get('kernel')->getProjectDir()."/public/exports/$file_name.xlsx";
		// from array
		$arrayData = $this->makeArray('Contact');
		if( 
			$spreadsheet->getActiveSheet()
		    	->fromArray(
			        $arrayData,  // The data to set
			        NULL,        // Array values with this value will not be set
			        'A1'         // Top left coordinate of the worksheet range where we want to set these values (default is A1)
		    	)
		    ) {
				$writerXlsx->save($new_file);
		        $message = "Successfully created $file_name.xlsx in /public/exports/$file_name.xlsx";
		    } else {
				$message = "File not created";
			}
		
        return $this->render('export/index.html.twig', [
            'controller_name' => 'ExportController',
            'message' => $message,
            'file_name' => $file_name,
        ]);
    }

    /**
     * @Route("/download/{file_name}", name="download")
     */
    public function downloadFile($file_name){
    	$file_path = $this->get('kernel')->getProjectDir()."/public/exports/$file_name.xlsx";
		$response = new BinaryFileResponse($file_path);
		$response->setContentDisposition(ResponseHeaderBag::DISPOSITION_ATTACHMENT);

		return $response;    	
    }

    private function makeArray($db){
    	$contacts = $this->em()->getRepository("App:$db")
    		->findAll();
    	$contacts_array = [];
    	foreach($contacts as $contact){
    		$contacts_array[] = [$contact->getName(), $contact->getPhoneNo()];
    	}
    	return $contacts_array;
    }

    private function em(){
        $em = $this->getDoctrine()->getManager();
        return $em;
    }

}
