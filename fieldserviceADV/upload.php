<?php


	$ticketID = $_POST['ticketid'];
	$seno = $_POST['seno'];
	
	define('UPLOAD_DIR', '../clientfiles/' . $seno . '/signaturesave/');
	
 	$date = date('-F-d-Y-H-i-s');
 	

 	//Get full size signature from post and set upload directory
 	
	$imgFull = $_POST['imgBase64'];   	
		
	$imgFull = str_replace('data:image/png;base64,', '', $imgFull);
	$imgFull = str_replace(' ', '+', $imgFull);
	
	$dataFull = base64_decode($imgFull);
	
	$fileFull = UPLOAD_DIR  . 'TicketID-' . $ticketID . '.png';
	
	
	
 	//Get full size signature from post and save as thumbnail  	
	// Create image from file
	$imgThumbnail = imagecreatefrompng($_POST['imgBase64']); 
	
	
	
	//Code To Resize Full Signature Into Thumbnail
	//**********************************************************
	// Target dimensions
	$max_width = 300;
	$max_height = 180;
	
	// Get current dimensions
	$old_width  = imagesx($imgThumbnail);
	$old_height = imagesy($imgThumbnail);
	
	// Calculate the scaling we need to do to fit the image inside our frame
	$scale = min($max_width/$old_width, $max_height/$old_height);
	
	// Get the new dimensions
	$newPNG_width  = ceil($scale*$old_width);
	$newPNG_height = ceil($scale*$old_height);
	
	// Create new empty image
	$newPNG = imagecreatetruecolor($newPNG_width, $newPNG_height);
	
    imagesavealpha($newPNG, true);

    $trans_colour = imagecolorallocatealpha($newPNG, 0, 0, 0, 127);
    imagefill($newPNG, 0, 0, $trans_colour);
    	
	
	// Resize old image into new
	imagecopyresampled($newPNG, $imgThumbnail, 0, 0, 0, 0, $newPNG_width, $newPNG_height, $old_width, $old_height);
	
 	// Catch the imagedata
	ob_start();
	imagepng($newPNG, NULL, 9);
	$dataThumb = ob_get_clean();
	
	//**********************************************************
	
	$fileThumb = UPLOAD_DIR  . 'TicketID-' . $ticketID . '-thumb.png';
	
	$successThumb = file_put_contents($fileThumb, $dataThumb);
	$successFull = file_put_contents($fileFull, $dataFull);
	
	//send request to ocr 
	print $successFull ? $fileFull : 'Unable to save the file.';
	
?>