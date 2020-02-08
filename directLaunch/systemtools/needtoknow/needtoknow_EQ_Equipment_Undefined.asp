<%	

	'***********************************************************************************************************************************************************
	 Response.Write("Begin Entries for EQ_Equipment Undefined Groups, Brands, Models, Manf, Class Codes, Status Codes, Acquistion Codes, Movement Codes<br><br>")
	'***********************************************************************************************************************************************************
	
	Set cnnEQEquipment = Server.CreateObject("ADODB.Connection")
	cnnEQEquipment.open (Session("ClientCnnString"))
	Set rsEQEquipment = Server.CreateObject("ADODB.Recordset")
	rsEQEquipment.CursorLocation = 3 	


	'**************************************************************************************************************
	'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR EQUIPMENT INSIGHT ASSET TAGS
	'**************************************************************************************************************
	SQL_EQEquipment = "DELETE FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SubModule ='Undefined Fields'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	'**************************************************************************************************************
	


	'**************************************************************************************************************
	'CHECK EQUIPMENT FOR UNDEFINED BRANDS
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_Brands WHERE UPPER(Brand) = 'UNDEFINED'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		BrandIntRecID = rsEQEquipment("InternalRecordIdentifier")
		UndefinedBrandEquipCount = NumberEquipmentRecsDefinedForBrand(BrandIntRecID)
		
		If UndefinedBrandEquipCount > 0 Then
	
			SCNeedToKnow_Module = "Equipment"
			SCNeedToKnow_SubModule = "Undefined Fields"
			SCNeedToKnow_SummaryDescription = "Undefined Brand Exists for Equipment"
			SCNeedToKnow_DetailedDescription1 = "You have " & UndefinedBrandEquipCount & " equipment records with undefined brands. Each piece of equipment must have a defined brand."
			SCNeedToKnow_InsightStaffOnly = 0
			
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			'*****************************************************************************************************************
			
		End If 
	
	End If
	
	




	'**************************************************************************************************************
	'CHECK EQUIPMENT FOR UNDEFINED MANUFACTURERS
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_Manufacturers WHERE UPPER(ManufacturerName) = 'UNDEFINED'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		ManufacturerIntRecID = rsEQEquipment("InternalRecordIdentifier")
		UndefinedManufacturerEquipCount = NumberEquipmentRecsDefinedForManufacturer(ManufacturerIntRecID)
		
		If UndefinedManufacturerEquipCount > 0 Then
	
			SCNeedToKnow_Module = "Equipment"
			SCNeedToKnow_SubModule = "Undefined Fields"
			SCNeedToKnow_SummaryDescription = "Undefined Manufacturer Exists for Equipment"
			SCNeedToKnow_DetailedDescription1 = "You have " & UndefinedManufacturerEquipCount & " equipment records with undefined manufacturers. Each piece of equipment must have a defined manufacturer."
			SCNeedToKnow_InsightStaffOnly = 0
			
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			'*****************************************************************************************************************
			
		End If 
	
	End If
	



	'**************************************************************************************************************
	'CHECK EQUIPMENT FOR UNDEFINED GROUPS
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_Groups WHERE UPPER(GroupName) = 'UNDEFINED'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		GroupIntRecID = rsEQEquipment("InternalRecordIdentifier")
		UndefinedGroupEquipCount = NumberEquipmentRecsDefinedForGroup(GroupIntRecID)
		
		If UndefinedGroupEquipCount > 0 Then
	
			SCNeedToKnow_Module = "Equipment"
			SCNeedToKnow_SubModule = "Undefined Fields"
			SCNeedToKnow_SummaryDescription = "Undefined Group Exists for Equipment"
			SCNeedToKnow_DetailedDescription1 = "You have " & UndefinedGroupEquipCount & " equipment records with undefined groups. Each piece of equipment must have a defined group."
			SCNeedToKnow_InsightStaffOnly = 0
			
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			'*****************************************************************************************************************
			
		End If 
	
	End If
	
	


	'**************************************************************************************************************
	'CHECK EQUIPMENT FOR UNDEFINED CLASSES
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_Classes WHERE UPPER(Class) = 'UNDEFINED'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		ClassIntRecID = rsEQEquipment("InternalRecordIdentifier")
		UndefinedClassEquipCount = NumberEquipmentRecsDefinedForClass(ClassIntRecID)
		
		If UndefinedClassEquipCount > 0 Then
	
			SCNeedToKnow_Module = "Equipment"
			SCNeedToKnow_SubModule = "Undefined Fields"
			SCNeedToKnow_SummaryDescription = "Undefined Class Exists for Equipment"
			SCNeedToKnow_DetailedDescription1 = "You have " & UndefinedClassEquipCount & " equipment records with undefined classes. Each piece of equipment must have a defined class."
			SCNeedToKnow_InsightStaffOnly = 0
			
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			'*****************************************************************************************************************
			
		End If 
	
	End If






	'**************************************************************************************************************
	'CHECK EQUIPMENT FOR UNDEFINED MODELS
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_Models WHERE UPPER(Model) = 'UNDEFINED'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		ModelIntRecID = rsEQEquipment("InternalRecordIdentifier")
		UndefinedModelEquipCount = NumberEquipmentRecsDefinedForModel(ModelIntRecID)
		
		If UndefinedModelEquipCount > 0 Then
	
			SCNeedToKnow_Module = "Equipment"
			SCNeedToKnow_SubModule = "Undefined Fields"
			SCNeedToKnow_SummaryDescription = "Undefined Model Exists for Equipment"
			SCNeedToKnow_DetailedDescription1 = "You have " & UndefinedModelEquipCount & " equipment records with undefined models. Each piece of equipment must have a defined model."
			SCNeedToKnow_InsightStaffOnly = 0
			
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			'*****************************************************************************************************************
			
		End If 
	
	End If
	




	'**************************************************************************************************************
	'CHECK EQUIPMENT FOR UNDEFINED STATUS CODES
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_StatusCodes WHERE UPPER(statusDesc) = 'UNDEFINED'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		StatusCodeIntRecID = rsEQEquipment("InternalRecordIdentifier")
		UndefinedStatusCodeEquipCount = NumberEquipmentRecsDefinedForStatusCode(StatusCodeIntRecID)
		
		If UndefinedStatusCodeEquipCount > 0 Then
	
			SCNeedToKnow_Module = "Equipment"
			SCNeedToKnow_SubModule = "Undefined Fields"
			SCNeedToKnow_SummaryDescription = "Undefined Status Code Exists for Equipment"
			SCNeedToKnow_DetailedDescription1 = "You have " & UndefinedStatusCodeEquipCount & " equipment records with undefined status codes. Each piece of equipment must have a defined status code."
			SCNeedToKnow_InsightStaffOnly = 0
			
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			'*****************************************************************************************************************
			
		End If 
	
	End If
	




	'**************************************************************************************************************
	'CHECK EQUIPMENT FOR UNDEFINED ACQUISITION CODES
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_AcquisitionCodes WHERE UPPER(acquisitionDesc) = 'UNDEFINED'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		AcquisitionCodeIntRecID = rsEQEquipment("InternalRecordIdentifier")
		UndefinedAcquisitionCodeEquipCount = NumberEquipmentRecsDefinedForAcquisitionCode(AcquisitionCodeIntRecID)
		
		If UndefinedAcquisitionCodeEquipCount > 0 Then
	
			SCNeedToKnow_Module = "Equipment"
			SCNeedToKnow_SubModule = "Undefined Fields"
			SCNeedToKnow_SummaryDescription = "Undefined Acquisition Code Exists for Equipment"
			SCNeedToKnow_DetailedDescription1 = "You have " & UndefinedAcquisitionCodeEquipCount & " equipment records with undefined acquisition codes. Each piece of equipment must have a defined acquisition code."
			SCNeedToKnow_InsightStaffOnly = 0
			
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			'*****************************************************************************************************************
			
		End If 
	
	End If
	




	'**************************************************************************************************************
	'CHECK EQUIPMENT FOR UNDEFINED CONDITION CODES
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_Condition WHERE UPPER(Condition) = 'UNDEFINED'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		ConditionCodeIntRecID = rsEQEquipment("InternalRecordIdentifier")
		UndefinedConditionCodeEquipCount = NumberEquipmentRecsDefinedForCondition(ConditionCodeIntRecID)
		
		If UndefinedConditionCodeEquipCount > 0 Then
	
			SCNeedToKnow_Module = "Equipment"
			SCNeedToKnow_SubModule = "Undefined Fields"
			SCNeedToKnow_SummaryDescription = "Undefined Condition Code Exists for Equipment"
			SCNeedToKnow_DetailedDescription1 = "You have " & UndefinedConditionCodeEquipCount & " equipment records with undefined condition codes. Each piece of equipment must have a defined condition code."
			SCNeedToKnow_InsightStaffOnly = 0
			
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			'*****************************************************************************************************************
			
		End If 
	
	End If
	



	'**************************************************************************************************************
	'CHECK EQUIPMENT FOR UNDEFINED Movement CODES
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_MovementCodes WHERE UPPER(movementCode) = 'UNDEFINED'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		MovementCodeIntRecID = rsEQEquipment("InternalRecordIdentifier")
		UndefinedMovementCodeEquipCount = NumberEquipmentRecsDefinedForMovementCode(MovementCodeIntRecID)
		
		If UndefinedMovementCodeEquipCount > 0 Then
	
			SCNeedToKnow_Module = "Equipment"
			SCNeedToKnow_SubModule = "Undefined Fields"
			SCNeedToKnow_SummaryDescription = "Undefined Movement Code Exists for Equipment"
			SCNeedToKnow_DetailedDescription1 = "You have " & UndefinedMovementCodeEquipCount & " equipment records with undefined movement codes. Each piece of equipment must have a defined movement code."
			SCNeedToKnow_InsightStaffOnly = 0
			
	
			'*****************************************************************************************************************
			'Check to see if record already exists in SC_NeedToKnow
			'*****************************************************************************************************************
			
			SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
			SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
			
			Set rsSCNeedToKnowCheckIfExists = cnnSCNeedToKnowCheckIfExists.Execute(SQL_SCNeedToKnowCheckIfExists)
			
			If rsSCNeedToKnowCheckIfExists.EOF Then
						
				SQL_SCNeedToKnow = "INSERT INTO SC_NeedToKnow (Module, SubModule, SummaryDescription, DetailedDescription1, InsightStaffOnly) VALUES "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " ('" & SCNeedToKnow_Module & "', '" & SCNeedToKnow_SubModule & "', '" & SCNeedToKnow_SummaryDescription & "', "
				SQL_SCNeedToKnow = SQL_SCNeedToKnow & " '" & SCNeedToKnow_DetailedDescription1 & "', " & SCNeedToKnow_InsightStaffOnly & ")"
				
				Set rsSCNeedToKnow = cnnSCNeedToKnow.Execute(SQL_SCNeedToKnow)
				
				If QuietMode = False Then
					Response.Write(SCNeedToKnow_Module & " - " & SCNeedToKnow_SubModule & " - " & SCNeedToKnow_SummaryDescription & " - " & SCNeedToKnow_DetailedDescription1 & "<br>")
					Response.Write("<hr>")
				End If
			
			End If
			'*****************************************************************************************************************
			
		End If 
	
	End If
	


	
	
	Set rsEQEquipment = Nothing
	cnnEQEquipment.Close
	Set cnnEQEquipment = Nothing

	'***********************************************************************************************************************************************************
	 Response.Write("End Entries for EQ_Equipment Undefined Groups, Brands, Models, Manf, Class Codes, Status Codes, Acquistion Codes, Movement Codes<br><br>")
	'***********************************************************************************************************************************************************
	
							
%>
