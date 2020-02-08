<%	

	'***********************************************************************************************************************************************************
	 Response.Write("Begin Entries for EQ_Equipment Zero Dollar Rentals<br><br>")
	'***********************************************************************************************************************************************************
	
	Set cnnEQEquipment = Server.CreateObject("ADODB.Connection")
	cnnEQEquipment.open (Session("ClientCnnString"))
	Set rsEQEquipment = Server.CreateObject("ADODB.Recordset")
	rsEQEquipment.CursorLocation = 3 	


	'**************************************************************************************************************
	'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR EQUIPMENT ZREO DOLLAR RENTALS
	'**************************************************************************************************************
	SQL_EQEquipment = "DELETE FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SubModule ='Zero Dollar Rentals'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	'**************************************************************************************************************

	'**************************************************************************************************************
	'CHECK EQUIPMENT FOR UNDEFINED BRANDS
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT EQ_CustomerEquipment.CustID, EQ_CustomerEquipment.EquipIntRecID, AR_Customer.Name, EQ_Models.Model, EQ_Equipment.SerialNumber, "
	SQL_EQEquipment = SQL_EQEquipment & "EQ_Equipment.AssetTag1, EQ_CustomerEquipment.RentAmt, EQ_CustomerEquipment.RentalFrequencyType, "
	SQL_EQEquipment = SQL_EQEquipment & "EQ_CustomerEquipment.RentalFrequencyNumber FROM EQ_Equipment INNER JOIN EQ_StatusCodes ON "
	SQL_EQEquipment = SQL_EQEquipment & "EQ_Equipment.StatusCodeIntRecID = EQ_StatusCodes.InternalRecordIdentifier INNER JOIN EQ_CustomerEquipment "
	SQL_EQEquipment = SQL_EQEquipment & "ON EQ_CustomerEquipment.EquipIntRecID = EQ_Equipment.InternalRecordIdentifier INNER JOIN EQ_Models "
	SQL_EQEquipment = SQL_EQEquipment & "ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier INNER JOIN AR_Customer "
	SQL_EQEquipment = SQL_EQEquipment & "ON EQ_CustomerEquipment.CustID = AR_Customer.CustNum "
	SQL_EQEquipment = SQL_EQEquipment & "WHERE (EQ_StatusCodes.statusDesc = 'RENTED') AND (EQ_CustomerEquipment.RentAmt <= 0) "
	SQL_EQEquipment = SQL_EQEquipment & "ORDER BY EQ_CustomerEquipment.CustID, EQ_Models.Model "
	
	'Response.Write(SQL_EQEquipment)
		
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		Do While NOT rsEQEquipment.EOF
		
			CustNum = rsEQEquipment("CustID")
			CustName = rsEQEquipment("Name")
			EquipIDIfApplicable = rsEQEquipment("EquipIntRecID")
			EquipModel = rsEQEquipment("Model")
			EquipSerialNumber = rsEQEquipment("SerialNumber")
			EquipAssetTag1 = rsEQEquipment("AssetTag1")
				
			SCNeedToKnow_Module = "Equipment"
			SCNeedToKnow_SubModule = "Zero Dollar Rentals"
			SCNeedToKnow_SummaryDescription = "Zero Dollar Rentals Exist for Equipment"
			SCNeedToKnow_DetailedDescription1 = "Customer " & CustNum & " (" & CustName & ") equipment model " & EquipModel & " with serial # " & EquipSerialNumber &  " and asset tag " & EquipAssetTag1 & "."
			If Len(CustNum) > 1 Then SCNeedToKnow_CustID = CustNum Else SC_NeedToKnowCustID = ""
			If Len(EquipIDIfApplicable) > 1 Then SCNeedToKnow_EquipIDIfApplicable = EquipIDIfApplicable Else SC_EquipIDIfApplicable = ""
			
			SCNeedToKnow_InsightStaffOnly = 0
			
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND CustIDIfApplicable = '" & SCNeedToKnow_CustID & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND EquipIDIfApplicable = '" & SCNeedToKnow_EquipIDIfApplicable & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & Replace(SCNeedToKnow_DetailedDescription1,"'","''") & "' "
			
			'Response.Write(SQL_SCNeedToKnowCheckIfExists & "<br>")
			
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, CustIDIfApplicable, EquipIDIfApplicable, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & Replace(SCNeedToKnow_DetailedDescription1,"'","''") & "', '" & SCNeedToKnow_CustID & "', '" & SCNeedToKnow_EquipIDIfApplicable & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			
			'*****************************************************************************************************************
			
		rsEQEquipment.MoveNext
		Loop

	End If
	
	
	
	Set rsEQEquipment = Nothing
	cnnEQEquipment.Close
	Set cnnEQEquipment = Nothing

	'***********************************************************************************************************************************************************
	 Response.Write("End Entries for EQ_Equipment Zero Dollar Rentals<br><br>")
	'***********************************************************************************************************************************************************
	
							
%>
