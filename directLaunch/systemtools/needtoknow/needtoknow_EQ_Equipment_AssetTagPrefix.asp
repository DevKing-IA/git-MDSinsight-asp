<%	

	'****************************************************************************************
	 Response.Write("Begin Entries for EQ_Equipment Insight Asset Tag<br><br>")
	'****************************************************************************************
	
	Set cnnEQEquipment = Server.CreateObject("ADODB.Connection")
	cnnEQEquipment.open (Session("ClientCnnString"))
	Set rsEQEquipment = Server.CreateObject("ADODB.Recordset")
	rsEQEquipment.CursorLocation = 3 	


	'**************************************************************************************************************
	'FIRST CLEAR OUT THE ENTRIES IN THE SC_NEEDTOKNOW TABLE FOR EQUIPMENT INSIGHT ASSET TAGS
	'**************************************************************************************************************
	SQL_EQEquipment = "DELETE FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SubModule ='Insight Asset Tag'"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	'**************************************************************************************************************
	


	'**************************************************************************************************************
	'CHECK CLASSES FOR BLANK INSIGHT ASSET TAG PREFIXES
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_Classes WHERE Class IN (SELECT DISTINCT Class FROM EQ_Classes)"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		Do While NOT rsEQEquipment.EOF
	
			EquipClassName = rsEQEquipment("Class")
			InsightAssetTagPrefix = rsEQEquipment("InsightAssetTagPrefix")
			
			
			If InsightAssetTagPrefix = "" OR IsNull(InsightAssetTagPrefix) OR IsEmpty(InsightAssetTagPrefix) OR Len(InsightAssetTagPrefix) < 1 Then
			
				SCNeedToKnow_Module = "Equipment"
				SCNeedToKnow_SubModule = "Insight Asset Tag"
				SCNeedToKnow_SummaryDescription = "Blank Insight Asset Tag Class Prefix"
				SCNeedToKnow_DetailedDescription1 = "The insight asset tag for the class " & EquipClassName & " is blank. Every class must have an insight asset tag prefix."
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
		
		rsEQEquipment.MoveNext
		Loop	
	
	End If
	
	'**************************************************************************************************************
	'CHECK MANUFACTURERS FOR BLANK INSIGHT ASSET TAG PREFIXES
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_Manufacturers WHERE ManufacturerName IN (SELECT DISTINCT ManufacturerName FROM EQ_Manufacturers)"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		Do While NOT rsEQEquipment.EOF
	
			ManufacturerName = rsEQEquipment("ManufacturerName")
			If ManufacturerName <> "" OR NOT IsNull(ManufacturerName) OR NOT IsEmpty(ManufacturerName) OR Len(ManufacturerName) >= 1 Then
				ManufacturerName = Replace(ManufacturerName,"'","''")
			End If

			InsightAssetTagPrefix = rsEQEquipment("InsightAssetTagPrefix")
			
			
			If InsightAssetTagPrefix = "" OR IsNull(InsightAssetTagPrefix) OR IsEmpty(InsightAssetTagPrefix) OR Len(InsightAssetTagPrefix) < 1 Then
			
				SCNeedToKnow_Module = "Equipment"
				SCNeedToKnow_SubModule = "Insight Asset Tag"
				SCNeedToKnow_SummaryDescription = "Blank Insight Asset Tag Manufacturer Prefix"
				SCNeedToKnow_DetailedDescription1 = "The insight asset tag for the manufacturer " & ManufacturerName & " is blank. Every manufacturer must have an insight asset tag prefix."
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
		
		rsEQEquipment.MoveNext
		Loop	
	
	End If
	
	
	'**************************************************************************************************************
	'CHECK BRANDS FOR BLANK INSIGHT ASSET TAG PREFIXES
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_Brands WHERE Brand IN (SELECT DISTINCT Brand FROM EQ_Brands)"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		Do While NOT rsEQEquipment.EOF
	
			BrandName = rsEQEquipment("Brand")
			If BrandName <> "" OR NOT IsNull(BrandName) OR NOT IsEmpty(BrandName) OR Len(BrandName) >= 1 Then
				BrandName = Replace(BrandName,"'","''")
			End If

			InsightAssetTagPrefix = rsEQEquipment("InsightAssetTagPrefix")
			
			
			If InsightAssetTagPrefix = "" OR IsNull(InsightAssetTagPrefix) OR IsEmpty(InsightAssetTagPrefix) OR Len(InsightAssetTagPrefix) < 1 Then
			
				SCNeedToKnow_Module = "Equipment"
				SCNeedToKnow_SubModule = "Insight Asset Tag"
				SCNeedToKnow_SummaryDescription = "Blank Insight Asset Tag Brand Prefix"
				SCNeedToKnow_DetailedDescription1 = "The insight asset tag for the brand " & BrandName & " is blank. Every brand must have an insight asset tag prefix."
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
		
		rsEQEquipment.MoveNext
		Loop	
	
	End If
	

	
	'**************************************************************************************************************
	'CHECK MODELS FOR BLANK INSIGHT ASSET TAG PREFIXES
	'**************************************************************************************************************
	
	SQL_EQEquipment = "SELECT * FROM EQ_Models WHERE Model IN (SELECT DISTINCT Model FROM EQ_Models)"
	Set rsEQEquipment = cnnEQEquipment.Execute(SQL_EQEquipment)
	
	If NOT rsEQEquipment.EOF Then
	
		Do While NOT rsEQEquipment.EOF
	
			ModelName = rsEQEquipment("Model")
			
			If ModelName <> "" OR NOT IsNull(ModelName) OR NOT IsEmpty(ModelName) OR Len(ModelName) >= 1 Then
				ModelName = Replace(ModelName,"'","''")
			End If
			
			InsightAssetTagPrefix = rsEQEquipment("InsightAssetTagPrefix")
			
			
			If InsightAssetTagPrefix = "" OR IsNull(InsightAssetTagPrefix) OR IsEmpty(InsightAssetTagPrefix) OR Len(InsightAssetTagPrefix) < 1 Then
			
				SCNeedToKnow_Module = "Equipment"
				SCNeedToKnow_SubModule = "Insight Asset Tag"
				SCNeedToKnow_SummaryDescription = "Blank Insight Asset Tag Model Prefix"
				SCNeedToKnow_DetailedDescription1 = "The insight asset tag for the model " & ModelName & " is blank. Every model must have an insight asset tag prefix."
				SCNeedToKnow_InsightStaffOnly = 0
				
		
				'*****************************************************************************************************************
				'Check to see if record already exists in SC_NeedToKnow
				'*****************************************************************************************************************
				
				SQL_SCNeedToKnowCheckIfExists = " SELECT * FROM SC_NeedToKnow WHERE Module = '" & SCNeedToKnow_Module & "' AND SubModule = '" & SCNeedToKnow_SubModule & "' "
				SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND SummaryDescription = '" & SCNeedToKnow_SummaryDescription & "' "
				SQL_SCNeedToKnowCheckIfExists = SQL_SCNeedToKnowCheckIfExists & " AND DetailedDescription1 = '" & SCNeedToKnow_DetailedDescription1 & "' "
				
				'Response.write(SQL_SCNeedToKnowCheckIfExists & "<br>")
				
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
		
		rsEQEquipment.MoveNext
		Loop	
	
	End If


	
	Set rsEQEquipment = Nothing
	cnnEQEquipment.Close
	Set cnnEQEquipment = Nothing

	'****************************************************************************************
	 Response.Write("End Entries for EQ_Equipment Insight Asset Tag<br><br>")
	'****************************************************************************************
	
							
%>