<?php

namespace App\Controller;

use App\Form\UploadType;
use App\Entity\Upload;
use App\Entity\Contact;
use Sensio\Bundle\FrameworkExtraBundle\Configuration\Route;
use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\Request;

class UploadController extends Controller
{
    /**
     * @Route("/upload", name="contacts_upload")
     */
    public function uploadAction(Request $request){
        $data = [];
        $upload = new Upload();
        $form = $this->createForm(UploadType::class, $upload);
        $form->handleRequest($request);

        if($form->isSubmitted() && $form->isValid()){
            $excelFile = $form->get('file')->getData();
            $originalName = $excelFile->getClientOriginalName();
        	$filepath = $this->get('kernel')->getProjectDir()."/public/excelFiles/";
            $excelFile->move($filepath, $originalName);
            $upload->setFile($originalName);
            $rows = $this->readThrough($originalName);
            $data['rows'] = $rows;
            if($this->addToDatabase($rows)){
                $this->save($upload);
                return $this->redirectToRoute('contact_index');
            } else {
                $this->addFlash('error', "There is nothing in that excel file. Please add contacts with column A: name and column B: phone number");
                return $this->redirectToRoute('contacts_upload');
            }
            
        }

        return $this->render('upload/upload.html.twig', [
            'form' => $form->createView(),
            'data' => $data,
        ]);
    }

    private function readThrough($filename){
    	// path to file
        $file = $this->get('kernel')->getProjectDir()."/public/excelFiles/$filename";
        // create a reader
		$reader  = $this->get('phpoffice.spreadsheet')->createReader('Xlsx');
		// load the spreadsheet to a variable
		$spreadsheet = $reader->load($file);
		// take data from spreadsheet and make array with it

		$reader->setReadDataOnly(TRUE);
		$worksheet = $spreadsheet->getActiveSheet();

		$rows = [];
		foreach ($worksheet->getRowIterator() as $row) {
		    $cellIterator = $row->getCellIterator();
		    $cellIterator->setIterateOnlyExistingCells(FALSE);

		    $col = [];
		    foreach ($cellIterator as $cell) {
		    	if($cell->getValue() != NULL && $cell->getValue() != ""){
		    		$col[] = $cell->getValue();
		    	}
		    	
		    }

		    $rows[] = $col;
		}

		return $rows;

    }

    private function addToDatabase($rows){
    	if(isset($rows[0][0]) ){
            foreach($rows as $row){
        		$contact = new Contact;	
        		$contact->setName($row[0]);
        		$contact->setPhoneNo("0".$row[1]);
        		$this->save($contact);
    	    }
            return true;
        } else {
            return false;
        }
    }

    private function em(){
        $em = $this->getDoctrine()->getManager();
        return $em;
    }

    private function save($entity){
        $this->em()->persist($entity);
        $this->em()->flush();        
    } 

}