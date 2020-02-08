    <?php
     
    $country = $_POST['txtCountry'];
    
	if ($country == 'United States') {
	    $a = 3;
	} elseif ($country == 'Canada') {
	    $a = 1;
	} else {
	    $a = 3;
	}
     
    $data = array(
      'User'          => $_POST['txtu1'],
      'Password'      => $_POST['txtu2'],
      'PhoneNumbers'  => array($_POST['txtTEXT_TO']),
      'Subject'       => $_POST['txtSubject'],
      'Message'       => $_POST['txtMessage'],
      'StampToSend'   => '1305582245',
      'MessageTypeID' => $a
      );
     
    $curl = curl_init('https://app.eztexting.com/sending/messages?format=xml');
    curl_setopt($curl, CURLOPT_POST, 1);
    curl_setopt($curl, CURLOPT_POSTFIELDS, http_build_query($data));
    curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
      // If you experience SSL issues, perhaps due to an outdated SSL cert
      // on your own server, try uncommenting the line below
      curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false);
    $response = curl_exec($curl);
    curl_close($curl);
     
    $xml = new SimpleXMLElement($response);
    if ( 'Failure' == $xml->Status ) {
      $errors = array();
      foreach( $xml->Errors->children() as $error ) {
        $errors[] = (string) $error;
      }
     
      echo 'Status: ' . $xml->Status . "\n" .
      'Errors: ' . implode(', ' , $errors) . "\n";
    } else {
      $phoneNumbers = array();
      foreach( $xml->Entry->PhoneNumbers->children() as $phoneNumber ) {
        $phoneNumbers[] = (string) $phoneNumber;
      }
     
      $localOptOuts = array();
      foreach( $xml->Entry->LocalOptOuts->children() as $phoneNumber ) {
        $localOptOuts[] = (string) $phoneNumber;
      }
     
      $globalOptOuts = array();
      foreach( $xml->Entry->GlobalOptOuts->children() as $phoneNumber ) {
        $globalOptOuts[] = (string) $phoneNumber;
      }
     
      $groups = array();
      foreach( $xml->Entry->Groups->children() as $group ) {
        $groups[] = (string) $group;
      }
     
//      echo 'Status: ' . $xml->Status . "\n" .
//     'Message ID : ' . $xml->Entry->ID . "\n" .
//      'Subject: ' . $xml->Entry->Subject . "\n" .
//      'Message: ' . $xml->Entry->Message . "\n" .
//      'Message Type ID: ' . $xml->Entry->MessageTypeID . "\n" .
//      'Total Recipients: ' . $xml->Entry->RecipientsCount . "\n" .
//      'Credits Charged: ' . $xml->Entry->Credits . "\n" .
//      'Time To Send: ' . $xml->Entry->StampToSend . "\n" .
//      'Phone Numbers: ' . implode(', ' , $phoneNumbers) . "\n" .
//      'Groups: ' . implode(', ' , $groups) . "\n" .
//      'Locally Opted Out Numbers: ' . implode(', ' , $localOptOuts) . "\n" .
//      'Globally Opted Out Numbers: ' . implode(', ' , $globalOptOuts) . "\n";
    }
     
    ?>