<?php

	$ticketID = $_POST['ticketid'];
	$seno = $_POST['seno'];
	
	//define('UPLOAD_DIR', '../clientfiles/' . $seno . '/signaturesave/');
	
	//define('UPLOAD_DIR', '../clientfiles/1071/signaturesave/');
	
	define('UPLOAD_DIR', '../clientfiles/1071/signaturesave/');
	
 	$date = date('-F-d-Y-H-i-s');
	$img = $_POST['imgBase64'];
	$img = str_replace('data:image/png;base64,', '', $img);
	$img = str_replace(' ', '+', $img);
	$data = base64_decode($img);
	$file = UPLOAD_DIR  . 'TicketID-' . $ticketID . '.png';
	//$file = UPLOAD_DIR . uniqid() . $date. '-TicketID-' . $ticketID . '.png';
	$success = file_put_contents($file, $data);
	//send request to ocr 
	print $success ? $file : 'Unable to save the file.';
?>