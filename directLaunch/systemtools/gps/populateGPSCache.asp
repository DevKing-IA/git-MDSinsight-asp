<script language="javascript" runat="server" src="json2.min.asp"></script>

<%

	Server.ScriptTimeout = 250000
	

	Set cnnClientdb = Server.CreateObject("ADODB.Connection")
	cnnClientdb.open Session("ClientCnnString")
	Set rsClientTables = Server.CreateObject("ADODB.Recordset")
	rsClientTables.CursorLocation = 3
	Set rsClientTablesForUpdating = Server.CreateObject("ADODB.Recordset")
	rsClientTablesForUpdating.CursorLocation = 3
   
	useAddr2 = False
	If ClientKey = "1071" or ClientKey = "1071d" Then useAddr2 = True

	ReDim tables(7)
	tables(0) = "AR_Customer"
	tables(1) = "AR_CustomerBillTo"
	tables(2) = "AR_CustomerShipTo"
	tables(3) = "IN_InvoiceHistHeader_1"
	tables(4) = "IN_InvoiceHistHeader_2"
	tables(5) = "PR_ProspectContacts"
	tables(6) = "PR_Prospects"
	tables(7) = "IC_Partners"
   
	For Each table In tables
	
	  Select Case table
	     Case "AR_Customer"
	        idField = "CustNum"
	        If useAddr2 Then streetField = "Addr2" Else streetField = "Addr1"
	        cityField = "City"
	        stateField = "State"
	        zipField = "Zip"
	        latField = "Latitude"
	        lonField = "Longitude"
	     Case "AR_CustomerBillTo", "AR_CustomerShipTo"
	        idField = "InternalRecordIdentifier"
	        If useAddr2 Then streetField = "Addr2" Else streetField = "Addr1"
	        cityField = "City"
	        stateField = "State"
	        zipField = "Zip"
	        latField = "Latitude"
	        lonField = "Longitude"
	     Case "IN_InvoiceHistHeader_1"
	        table = "IN_InvoiceHistHeader"
	        idField = "InternalRecordIdentifier"
	        If useAddr2 Then streetField = "ShipToAddr2" Else streetField = "ShipToAddr1"
	        cityField = "ShipToCity"
	        stateField = "ShipToState"
	        zipField = "ShipToPostalCode"
	        latField = "ShipToLatitude"
	        lonField = "ShipToLongitude"
	     Case "IN_InvoiceHistHeader_2"
	        table = "IN_InvoiceHistHeader"
	        idField = "InternalRecordIdentifier"
	        If useAddr2 Then streetField = "BillToAddr2" Else streetField = "BillToAddr1"
	        cityField = "BillToCity"
	        stateField = "BillToState"
	        zipField = "BillToPostalCode"
	        latField = "BillToLatitude"
	        lonField = "BillToLongitude"
	     Case "PR_ProspectContacts"
	        If left(database, 3) = "USC" Then idField = "ProspectIntRecID" Else idField = "InternalRecordIdentifier"
	        streetField = "Address1"
	        cityField = "City"
	        stateField = "State"
	        zipField = "PostalCode"
	        latField = "Latitude"
	        lonField = "Longitude"
	     Case "PR_Prospects"
	        idField = "InternalRecordIdentifier"
	        streetField = "Street"
	        cityField = "City"
	        stateField = "State"
	        zipField = "PostalCode"
	        latField = "Latitude"
	        lonField = "Longitude"
	     Case "IC_Partners"
	        idField = "InternalRecordIdentifier"
	        streetField = "partnerAddress"
	        cityField = "partnerCity"
	        stateField = "partnerState"
	        zipField = "partnerZip"
	        latField = "Latitude"
	        lonField = "Longitude"
	  End Select
	  
	  SQLClientTable  = "SELECT " & idField & "," & streetField & "," & cityField & "," & stateField & "," & zipField & "," & latField & "," & lonField & " FROM " & table
	  
	  Response.Write ("<font color='green'>" & ClientKey  & "    " & Now() &  "</font>" & "111: " & SQLClientTable & "<BR>")

	  Set rsClientTables = cnnClientdb.Execute(SQLClientTable)
	  rsClientTables.CacheSize = 20000
	  
	  i = 0
	  
	  Do Until rsClientTables.EOF
	  
	     i = i + 1
	     
	     latlong = lookupGPS(rsClientTables.Fields.Item(streetField), rsClientTables.Fields.Item(cityField), rsClientTables.Fields.Item(stateField), rsClientTables.Fields.Item(zipField),ClientKey)
	     
	     If latlong  <> "" Then

	        	lat = left(latlong,Instr(latlong,":")-1)
	        	lon = right(latlong,len(latlong)-instr(latlong,":"))
	        	
				Response.Write("<font color='green'>" & ClientKey  & "    " & Now() &  "</font>" & "Lat:" & lat & "<br>")
				Response.Write("<font color='green'>" & ClientKey  & "    " & Now() &  "</font>" & "Lon:" & lon & "<br>")	   
				
				SQLUpdate = "UPDATE " & table & " SET " & latField & " = '" & lat & "'," & lonField & " = '" & lon & "' WHERE "
				SQLUpdate = SQLUpdate & idField & "= '" & rsClientTables.Fields.Item(idField) & "'"
				
				Response.Write("<font color='green'>" & ClientKey  & "    " & Now() &  "</font>"  & SQLUpdate & "<br>")
				
	        	Set rsClientTablesForUpdating = cnnClientdb.Execute(SQLUpdate)

	     End If
	     
	     rsClientTables.MoveNext
	     If i Mod 100 = 0 Then
	         Response.Write "i= " & i & "<BR>"
	         Response.Flush
	     End If
	  Loop
	  
	  rsClientTables.close
	  Response.Write "Done: " & i & "<BR>"
	Next


   cnnClientDB.close
   set rsClientTables = Nothing
   set cnnClientDB = Nothing

Response.Write ("<font color='green'>" & ClientKey  & "    " & Now() &  "</font>" & "<BR>Done.<BR>")



%>

