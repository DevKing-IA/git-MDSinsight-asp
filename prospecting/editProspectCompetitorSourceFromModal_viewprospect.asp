<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->
<!--#include file="../inc/InSightFuncs_Users.asp"-->
<%

txtInternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")
ProspectName = GetProspectNameByNumber(txtInternalRecordIdentifier)

txtPrimaryCompetitor = Request.Form("txtPrimaryCompetitor")
txtTelemarketerUserNo = Request.Form("txtTelemarketerUserNo")
txtLeadSourceNumber = Request.Form("txtLeadSource")

chkBottledWater = Request.Form("chkBottledWater")
chkFilteredWater = Request.Form("chkFilteredWater")
chkOCS = Request.Form("chkOCS")
chkOCS_Supply = Request.Form("chkOCS_Supply")
chkOfficeSupplies = Request.Form("chkOfficeSupplies")
chkVending = Request.Form("chkVending")
chkMicroMarket = Request.Form("chkMicroMarket")
chkPantry = Request.Form("chkPantry")

txtFormerCustomerNumber = Request.Form("txtFormerCustomerNumber")
txtFormerCustomerCancelDate = Request.Form("txtFormerCustomerCancelDate")

'*******************************************************************************************************************
'GET ORIGINAL VALUES FOR OPPORTUNITY FIELDS FOR AUDIT TRAIL CHANGES
'*******************************************************************************************************************

	SQLProspect = "SELECT * FROM PR_Prospects WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier 

	Set cnnProspect = Server.CreateObject("ADODB.Connection")
	cnnProspect.open (Session("ClientCnnString"))
	Set rsProspect = Server.CreateObject("ADODB.Recordset")
	rsProspect.CursorLocation = 3 
	Set rsProspect = cnnProspect.Execute(SQLProspect)

	If not rsProspect.EOF Then
		ORIG_TelemarketerUserNo = rsProspect("TelemarketerUserNo")
		ORIG_LeadSourceNumber = rsProspect("LeadSourceNumber")	
		ORIG_FormerCustNum = rsProspect("FormerCustNum")
		ORIG_FormerCustomerCancelDate = rsProspect("CancelDate")										
	End If
	set rsProspect = Nothing
	cnnProspect.close
	set cnnProspect = Nothing
	
	If txtPrimaryCompetitor <> "" Then
	
		SQLCompetitors = "SELECT * FROM PR_ProspectCompetitors WHERE ProspectRecID = " & txtInternalRecordIdentifier & " AND CompetitorRecID = " & txtPrimaryCompetitor
		
		Response.Write(SQLCompetitors & "<br>") 	
			
		Set cnnCompetitors = Server.CreateObject("ADODB.Connection")
		cnnCompetitors.open (Session("ClientCnnString"))
		Set rsCompetitors = Server.CreateObject("ADODB.Recordset")
		rsCompetitors.CursorLocation = 3 
		Set rsCompetitors = cnnCompetitors.Execute(SQLCompetitors)
			
		If not rsCompetitors.EOF Then
			ORIG_CompetitorRecID = rsCompetitors("CompetitorRecID")
			ORIG_PrimaryCompetitor = rsCompetitors("PrimaryCompetitor")
			Orig_BottledWater = rsCompetitors("BottledWater")
			Orig_FilteredWater = rsCompetitors("FilteredWater")
			Orig_OCS = rsCompetitors("OCS")
			Orig_OCS_Supply = rsCompetitors("OCS_Supply")
			Orig_OfficeSupplies = rsCompetitors("OfficeSupplies")
			Orig_Vending = rsCompetitors("Vending")
			Orig_Micromarket = rsCompetitors("Micromarket")
			Orig_Pantry = rsCompetitors("Pantry")	
		End If
	
		If (Orig_OCS <> "" AND Orig_OCS <> 0) Then Orig_OCS = 1 Else Orig_OCS = 0
		If (Orig_OCS_Supply <> "" AND Orig_OCS_Supply <> 0) Then Orig_OCS_Supply = 1 Else Orig_OCS_Supply = 0
		If (Orig_BottledWater <> "" AND Orig_BottledWater <> 0) Then Orig_BottledWater = 1 Else Orig_BottledWater = 0
		If (Orig_FilteredWater <> "" AND Orig_FilteredWater <> 0) Then Orig_FilteredWater = 1 Else Orig_FilteredWater = 0	
		If (Orig_Vending <> "" AND Orig_Vending <> 0) Then Orig_Vending = 1 Else Orig_Vending = 0
		If (Orig_Micromarket <> "" AND Orig_Micromarket <> 0) Then Orig_Micromarket = 1 Else Orig_Micromarket = 0
		If (Orig_Pantry <> "" AND Orig_Pantry <> 0) Then Orig_Pantry = 1 Else Orig_Pantry = 0
		If (Orig_OfficeSupplies <> "" AND Orig_OfficeSupplies <> 0) Then Orig_OfficeSupplies = 1 Else Orig_OfficeSupplies = 0
		
	End If

'*******************************************************************************************************************
'SET DEFAULT VALUES FOR ANY NON REQUIRED FIELDS LEFT BLANK DURING THE EDIT PROCESS
'*******************************************************************************************************************

	If txtTelemarketerUserNo = "" Then txtTelemarketerUserNo = 0
	If ORIG_TelemarketerUserNo= "" Then txtORIG_TelemarketerUserNoLeadSource = 0
	
	If ORIG_CompetitorRecID = "" Then ORIG_CompetitorRecID = 0
	If txtPrimaryCompetitor = "" Then txtPrimaryCompetitor = 0
	
	If txtLeadSourceNumber = "" Then txtLeadSourceNumber = 0
	If ORIG_LeadSourceNumber = "" Then ORIG_LeadSourceNumber = 0
	
	
	If (chkBottledWater <> "" AND chkBottledWater = "on") Then chkBottledWater = 1 Else chkBottledWater = 0
	If (chkFilteredWater <> "" AND chkFilteredWater = "on") Then chkFilteredWater = 1 Else chkFilteredWater = 0
	If (chkOCS <> "" AND chkOCS = "on") Then chkOCS = 1 Else chkOCS = 0
	If (chkOCS_Supply <> "" AND chkOCS_Supply = "on") Then chkOCS_Supply = 1 Else chkOCS_Supply = 0
	If (chkOfficeSupplies <> "" AND chkOfficeSupplies = "on") Then chkOfficeSupplies = 1 Else chkOfficeSupplies = 0
	If (chkVending <> "" AND chkVending = "on") Then chkVending = 1 Else chkVending = 0
	If (chkMicroMarket <> "" AND chkMicroMarket = "on") Then chkMicroMarket = 1 Else chkMicroMarket = 0
	If (chkPantry <> "" AND chkPantry = "on") Then chkPantry = 1 Else chkPantry = 0

'*******************************************************************************************************************



'*******************************************************************************************************************
'PERFORM SQL UPDATE INTO PR_Prospects AND PR_ProspectCompetitors
'*******************************************************************************************************************

	'******************************************************
	'Update PR_Prospects
	'******************************************************
	
	SQLProspectUpdate = "UPDATE PR_Prospects SET LeadSourceNumber = " & cInt(txtLeadSourceNumber) & ", TelemarketerUserNo = " & cInt(txtTelemarketerUserNo) & ", "
	SQLProspectUpdate = SQLProspectUpdate & "FormerCustNum = '" & txtFormerCustomerNumber & "', CancelDate = '" & txtFormerCustomerCancelDate & "' "
	SQLProspectUpdate = SQLProspectUpdate & "WHERE InternalRecordIdentifier = " & txtInternalRecordIdentifier 
	
	'Response.write(SQLProspectUpdate & "<br><br>")
	
	Set cnnProspectUpdate = Server.CreateObject("ADODB.Connection")
	cnnProspectUpdate.open (Session("ClientCnnString"))
	Set rsProspectUpdate = Server.CreateObject("ADODB.Recordset")
	rsProspectUpdate.CursorLocation = 3 
	Set rsProspectUpdate = cnnProspectUpdate.Execute(SQLProspectUpdate)
	
	Set rsProspectUpdate = Nothing
	cnnProspectUpdate.Close
	Set cnnProspectUpdate = Nothing


	'******************************************************
	'Update/Insert into PR_ProspectCompetitors
	'******************************************************
	Set cnnProspectCompetitorCheck = Server.CreateObject("ADODB.Connection")
	cnnProspectCompetitorCheck.open (Session("ClientCnnString"))
	Set rsProspectCompetitorCheck = Server.CreateObject("ADODB.Recordset")
	rsProspectCompetitorCheck.CursorLocation = 3 
	
	SQLProspectCompetitorCheck = "SELECT * FROM PR_ProspectCompetitors WHERE ProspectRecID = " & txtInternalRecordIdentifier & " AND PrimaryCompetitor = 1"
	Set rsProspectCompetitorCheck = cnnProspectCompetitorCheck.Execute(SQLProspectCompetitorCheck)
	
	Response.Write("<br>" & SQLProspectCompetitorCheck)

	'******************************************************************************************************************
	'CHECK TO SEE IF THIS PROSPECT HAS A PRIMARY COMPETITOR RECORD ALREADY
	'THIS EDIT MODAL IS FOR ADDING/UPDATING A PRIMARY COMPETITOR ONLY, SO ALL CHECKS ARE BASED ON PRIMARY
	'COMPETITOR FIELD
	'******************************************************************************************************************

	If NOT rsProspectCompetitorCheck.EOF Then
	
		PrimaryCompetitorID = GetPrimaryCompetitorIDByProspectNumber(txtInternalRecordIdentifier)
		
		If PrimaryCompetitorID <> "" Then
			CurrentCompetitorName = GetCompetitorByNum(PrimaryCompetitorID)
		End If
	
		'***************************************************************************************************************
		'IF THIS PROSPECT HAS A PRIMARY COMPETITOR, WE WILL BE IN THIS BRANCH
		'BECAUSE THEY ALREADY HAVE A PRIMARY COMPETITOR, WE NEED TO SEE IF THE COMPETITOR WAS CHANGED FROM THE ORIGINAL
		'PRIMARY COMPETITOR THAT LOADED IN THE MODAL.
		'***************************************************************************************************************

		If cInt(ORIG_CompetitorRecID) <> cInt(txtPrimaryCompetitor) Then
		
			'***************************************************************************************************************
			'IF THE PRIMARY COMPETITOR WAS CHANGED, WE HAVE TO CHECK AND SEE IF THIS COMPETITOR ALREADY EXISTS AS A 
			'COMPETITOR FOR THIS PROSPECT. IF IT DOES NOT EXIST AS A COMPETITOR FOR THIS PROSPECT, WE MUST ADD THIS
			'COMPETITOR TO THIS PROSPECT VIA A SQL INSERT. IF THE NEW COMPETITOR SELECTED IS ALREADY A COMPETITOR FOR
			'THIS PROSPECT, WE JUST UPDATE THE OFFERINGS.
			'LASTLY, WE HAVE TO SET THIS NEW COMPETITOR TO BE THE PRIMARY COMPETITOR FOR THIS PROSPECT AND CHANGE ALL
			'OTHER PRIMARY COMPETITORS FOR THIS PROSPECT TO 0
			'***************************************************************************************************************
				
			SQLProspectCompOfferingCheck = "SELECT * FROM PR_ProspectCompetitors WHERE CompetitorRecID = " & txtPrimaryCompetitor & " AND ProspectRecID = " &  txtInternalRecordIdentifier
			Set cnnProspectCompOfferingCheck = Server.CreateObject("ADODB.Connection")
			cnnProspectCompOfferingCheck.open (Session("ClientCnnString"))
			Set rsProspectCompOfferingCheck = Server.CreateObject("ADODB.Recordset")
			rsProspectCompOfferingCheck.CursorLocation = 3 
			Set rsProspectCompOfferingCheck = cnnProspectCompOfferingCheck.Execute(SQLProspectCompOfferingCheck)
			
			If NOT rsProspectCompOfferingCheck.EOF Then
				SQLProspectCompOfferingUpdate = "UPDATE PR_ProspectCompetitors SET BottledWater = " & chkBottledWater & ", FilteredWater = " & chkFilteredWater & ", OCS = " & chkOCS & ", "
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & "OCS_Supply = " & chkOCS_Supply & ", Vending = " & chkVending & ", MicroMarket = " & chkMicroMarket & ", "
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & "Pantry= " & chkPantry & ", OfficeSupplies= " & chkOfficeSupplies & " "
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & " WHERE CompetitorRecID = " & txtPrimaryCompetitor & " AND ProspectRecID = " &  txtInternalRecordIdentifier
				
				CompetitorName = GetCompetitorByNum(txtPrimaryCompetitor)
				
				'*****************************************
				'AUDIT TRAIL ENTRIES
				'*****************************************

				If Orig_BottledWater <> chkBottledWater Then
					If (chkBottledWater = 1 OR chkBottledWater = vbTrue) Then
						Description = "Bottled water was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Bottled water was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Bottled water was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Bottled water was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_FilteredWater <> chkFilteredWater Then
					If (chkFilteredWater = 1 OR chkFilteredWater = vbTrue) Then
						Description = "Filtered water was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Filtered water was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Filtered water was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Filtered water was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OCS <> chkOCS Then
					If (chkOCS = 1 OR chkOCS = vbTrue) Then
						Description = "OCS was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "OCS was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OCS_Supply <> chkOCS_Supply Then
				
					If (chkOCS_Supply = 1 OR chkOCS_Supply = vbTrue) Then
						Description = "OCS Supply was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS Supply was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "OCS Supply was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS Supply was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OfficeSupplies <> chkOfficeSupplies Then
				
					If (chkOfficeSupplies = 1 OR chkOfficeSupplies = vbTrue) Then
						Description = "Office Supplies was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Office Supplies was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Office Supplies was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Office Supplies was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
			
				End If
				If Orig_Vending <> chkVending Then
					If (chkVending = 1 OR chkVending = vbTrue) Then
						Description = "Vending was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Vending was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Vending was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Vending was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_Micromarket <> chkMicromarket Then

					If (chkMicromarket = 1 OR chkMicromarket = vbTrue) Then
						Description = "Micromarket was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Micromarket was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Micromarket was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Micromarket was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_Pantry <> chkPantry Then
					If (chkPantry = 1 OR chkPantry = vbTrue) Then
						Description = "Pantry was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Pantry was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Pantry was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Pantry was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				'*****************************************		
				'END AUDIT TRAIL ENTRIES
				'*****************************************				
			Else
				SQLProspectCompOfferingUpdate = "INSERT INTO PR_ProspectCompetitors (ProspectRecID, CompetitorRecID, PrimaryCompetitor, BottledWater, FilteredWater, "
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & " OCS, OCS_Supply, Vending, MicroMarket, Pantry, OfficeSupplies) VALUES ("
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & txtInternalRecordIdentifier & "," & txtPrimaryCompetitor & ",1, " & chkBottledWater & "," & chkFilteredWater & ", "
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & chkOCS & "," & chkOCS_Supply & "," & chkVending & ","
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & chkMicroMarket & "," & chkPantry & "," & chkOfficeSupplies & ")"

				'*****************************************
				'AUDIT TRAIL ENTRIES
				'*****************************************
				
				NewCompetitorName = GetCompetitorByNum(txtPrimaryCompetitor)
				
				Description = NewCompetitorName & " was set to be the primary competitor for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
				CreateAuditLogEntry GetTerm("Prospecting")& " primary competitor change ",GetTerm("Prospecting"),"Minor",0,Description
				Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
				
				If PrimaryCompetitorID <> "" Then
					Description = CurrentCompetitorName & " was un-set as the primary competitor for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
					CreateAuditLogEntry GetTerm("Prospecting")& " primary competitor change ",GetTerm("Prospecting"),"Minor",0,Description
					Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
				End If
				'*****************************************
				
			
				'*****************************************
				'AUDIT TRAIL ENTRIES
				'*****************************************

				If Orig_BottledWater <> chkBottledWater Then
					If (chkBottledWater = 1 OR chkBottledWater = vbTrue) Then
						Description = "Bottled water was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Bottled water was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Bottled water was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Bottled water was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_FilteredWater <> chkFilteredWater Then
					If (chkFilteredWater = 1 OR chkFilteredWater = vbTrue) Then
						Description = "Filtered water was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Filtered water was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Filtered water was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Filtered water was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OCS <> chkOCS Then
					If (chkOCS = 1 OR chkOCS = vbTrue) Then
						Description = "OCS was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "OCS was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OCS_Supply <> chkOCS_Supply Then
				
					If (chkOCS_Supply = 1 OR chkOCS_Supply = vbTrue) Then
						Description = "OCS Supply was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS Supply was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "OCS Supply was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS Supply was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OfficeSupplies <> chkOfficeSupplies Then
				
					If (chkOfficeSupplies = 1 OR chkOfficeSupplies = vbTrue) Then
						Description = "Office Supplies was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Office Supplies was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Office Supplies was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Office Supplies was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
			
				End If
				If Orig_Vending <> chkVending Then
					If (chkVending = 1 OR chkVending = vbTrue) Then
						Description = "Vending was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Vending was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Vending was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Vending was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_Micromarket <> chkMicromarket Then

					If (chkMicromarket = 1 OR chkMicromarket = vbTrue) Then
						Description = "Micromarket was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Micromarket was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Micromarket was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Micromarket was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_Pantry <> chkPantry Then
					If (chkPantry = 1 OR chkPantry = vbTrue) Then
						Description = "Pantry was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Pantry was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Pantry was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Pantry was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				'*****************************************				
				
			End If
			
			Set rsProspectCompOfferingCheck = Nothing
			cnnProspectCompOfferingCheck.Close
			Set cnnProspectCompOfferingCheck = Nothing
			
			Response.write("<br>" & SQLProspectCompOfferingUpdate)
			
			Set cnnProspectCompOfferingUpdate = Server.CreateObject("ADODB.Connection")
			cnnProspectCompOfferingUpdate.open (Session("ClientCnnString"))
			Set rsProspectCompOfferingUpdate = Server.CreateObject("ADODB.Recordset")
			rsProspectCompOfferingUpdate.CursorLocation = 3 
			Set rsProspectCompOfferingUpdate = cnnProspectCompOfferingUpdate.Execute(SQLProspectCompOfferingUpdate)
			
			SQLProspectCompOfferingUpdate = "UPDATE PR_ProspectCompetitors SET PrimaryCompetitor = 0 WHERE ProspectRecID = " & txtInternalRecordIdentifier & " AND CompetitorRecID <> " & txtPrimaryCompetitor 
			Set rsProspectCompOfferingUpdate = cnnProspectCompOfferingUpdate.Execute(SQLProspectCompOfferingUpdate)
				
			Set rsProspectCompOfferingUpdate = Nothing
			cnnProspectCompOfferingUpdate.Close
			Set cnnProspectCompOfferingUpdate = Nothing
			


			
		Else

		
			'***************************************************************************************************************
			'THE CURRENT PRIMARY COMPETITOR WAS NOT CHANGED IN THE MODAL, SO WE JUST NEED TO UPDATE THE OFFERINGS
			'FOR THE CURRENT PRIMARY COMPETITOR
			'***************************************************************************************************************
		
			Set cnnProspectCompOfferingUpdate = Server.CreateObject("ADODB.Connection")
			cnnProspectCompOfferingUpdate.open (Session("ClientCnnString"))
			Set rsProspectCompOfferingUpdate = Server.CreateObject("ADODB.Recordset")
			rsProspectCompOfferingUpdate.CursorLocation = 3 
	
				
			SQLProspectCompOfferingUpdate = "UPDATE PR_ProspectCompetitors SET BottledWater = " & chkBottledWater & ", FilteredWater = " & chkFilteredWater & ", OCS = " & chkOCS & ", "
			SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & "OCS_Supply = " & chkOCS_Supply & ", Vending = " & chkVending & ", MicroMarket = " & chkMicroMarket & ", "
			SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & "Pantry= " & chkPantry & ", OfficeSupplies= " & chkOfficeSupplies & " "
			SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & " WHERE CompetitorRecID = " & txtPrimaryCompetitor & " AND ProspectRecID = " &  txtInternalRecordIdentifier
			
			Set rsProspectCompOfferingUpdate = cnnProspectCompOfferingUpdate.Execute(SQLProspectCompOfferingUpdate)
			
			Set rsProspectCompOfferingUpdate = Nothing
			cnnProspectCompOfferingUpdate.Close
			Set cnnProspectCompOfferingUpdate = Nothing
	
			
			Response.write("<br>" & SQLProspectCompOfferingUpdate)
			
			PrimaryCompetitorID = GetPrimaryCompetitorIDByProspectNumber(txtInternalRecordIdentifier)
			
			If PrimaryCompetitorID <> "" Then
				CompetitorName = GetCompetitorByNum(PrimaryCompetitorID)
			End If

			
				'*****************************************
				'AUDIT TRAIL ENTRIES
				'*****************************************

				If Orig_BottledWater <> chkBottledWater Then
					If (chkBottledWater = 1 OR chkBottledWater = vbTrue) Then
						Description = "Bottled water was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Bottled water was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Bottled water was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Bottled water was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_FilteredWater <> chkFilteredWater Then
					If (chkFilteredWater = 1 OR chkFilteredWater = vbTrue) Then
						Description = "Filtered water was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Filtered water was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Filtered water was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Filtered water was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OCS <> chkOCS Then
					If (chkOCS = 1 OR chkOCS = vbTrue) Then
						Description = "OCS was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "OCS was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OCS_Supply <> chkOCS_Supply Then
				
					If (chkOCS_Supply = 1 OR chkOCS_Supply = vbTrue) Then
						Description = "OCS Supply was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS Supply was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "OCS Supply was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS Supply was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OfficeSupplies <> chkOfficeSupplies Then
				
					If (chkOfficeSupplies = 1 OR chkOfficeSupplies = vbTrue) Then
						Description = "Office Supplies was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Office Supplies was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Office Supplies was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Office Supplies was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
			
				End If
				If Orig_Vending <> chkVending Then
					If (chkVending = 1 OR chkVending = vbTrue) Then
						Description = "Vending was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Vending was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Vending was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Vending was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_Micromarket <> chkMicromarket Then

					If (chkMicromarket = 1 OR chkMicromarket = vbTrue) Then
						Description = "Micromarket was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Micromarket was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Micromarket was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Micromarket was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_Pantry <> chkPantry Then
					If (chkPantry = 1 OR chkPantry = vbTrue) Then
						Description = "Pantry was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Pantry was added as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Pantry was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Pantry was removed as a competitor category being offered by the competitor " & CompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				'*****************************************
						
			
		End If
		
	Else
			'***************************************************************************************************************
			'THIS PROSPECT HAD NO PRIMARY COMPETITOR PREVIOUSLY, SO IF THEY CHOSE A PRIMARY COMPETITOR IN THE EDIT MODAL,
			'WE MUST PERFORM AN INSERT INTO PR_PROSPECTCOMPETITORS TABLE - NO UPDATE WILL EVER NEED TO HAPPEN
			'***************************************************************************************************************
	
			If cInt(txtPrimaryCompetitor) <> 0 Then
			
				Set cnnProspectCompOfferingUpdate = Server.CreateObject("ADODB.Connection")
				cnnProspectCompOfferingUpdate.open (Session("ClientCnnString"))
				Set rsProspectCompOfferingUpdate = Server.CreateObject("ADODB.Recordset")
				rsProspectCompOfferingUpdate.CursorLocation = 3 
	
				SQLProspectCompOfferingUpdate = "INSERT INTO PR_ProspectCompetitors (ProspectRecID, CompetitorRecID, PrimaryCompetitor, BottledWater, FilteredWater, "
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & " OCS, OCS_Supply, Vending, MicroMarket, Pantry, OfficeSupplies) VALUES ("
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & txtInternalRecordIdentifier & "," & txtPrimaryCompetitor & ",1," & chkBottledWater & "," & chkFilteredWater & ", "
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & chkOCS & "," & chkOCS_Supply & "," & chkVending & ","
				SQLProspectCompOfferingUpdate = SQLProspectCompOfferingUpdate & chkMicroMarket & "," & chkPantry & "," & chkOfficeSupplies & ")"
				
				Response.write("<br>" & SQLProspectCompOfferingUpdate)
				
				Set rsProspectCompOfferingUpdate = cnnProspectCompOfferingUpdate.Execute(SQLProspectCompOfferingUpdate)
			
				Set rsProspectCompOfferingUpdate = Nothing
				cnnProspectCompOfferingUpdate.Close
				Set cnnProspectCompOfferingUpdate = Nothing
				
				'*****************************************
				'AUDIT TRAIL ENTRIES
				'*****************************************
				NewCompetitorName = GetCompetitorByNum(txtPrimaryCompetitor)
				
				Description = NewCompetitorName & " was set to be the primary competitor for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
				CreateAuditLogEntry GetTerm("Prospecting")& " primary competitor change ",GetTerm("Prospecting"),"Minor",0,Description
				Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
				
				
				If Orig_BottledWater <> chkBottledWater Then
					If (chkBottledWater = 1 OR chkBottledWater = vbTrue) Then
						Description = "Bottled water was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Bottled water was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Bottled water was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Bottled water was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_FilteredWater <> chkFilteredWater Then
					If (chkFilteredWater = 1 OR chkFilteredWater = vbTrue) Then
						Description = "Filtered water was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Filtered water was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Filtered water was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Filtered water was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OCS <> chkOCS Then
					If (chkOCS = 1 OR chkOCS = vbTrue) Then
						Description = "OCS was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "OCS was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OCS_Supply <> chkOCS_Supply Then
				
					If (chkOCS_Supply = 1 OR chkOCS_Supply = vbTrue) Then
						Description = "OCS Supply was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS Supply was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "OCS Supply was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "OCS Supply was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_OfficeSupplies <> chkOfficeSupplies Then
				
					If (chkOfficeSupplies = 1 OR chkOfficeSupplies = vbTrue) Then
						Description = "Office Supplies was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Office Supplies was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Office Supplies was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Office Supplies was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
			
				End If
				If Orig_Vending <> chkVending Then
					If (chkVending = 1 OR chkVending = vbTrue) Then
						Description = "Vending was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Vending was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Vending was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Vending was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_Micromarket <> chkMicromarket Then

					If (chkMicromarket = 1 OR chkMicromarket = vbTrue) Then
						Description = "Micromarket was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Micromarket was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Micromarket was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Micromarket was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				If Orig_Pantry <> chkPantry Then
					If (chkPantry = 1 OR chkPantry = vbTrue) Then
						Description = "Pantry was added as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Pantry was added as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					Else
						Description = "Pantry was removed as a competitor category being offered by the competitor " & NewCompetitorName & " for the prospect " & GetProspectNameByNumber(txtInternalRecordIdentifier) & " by " & GetUserDisplayNameByUserNo(Session("UserNo"))
						CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
						
						Description = "Pantry was removed as a competitor category being offered by the competitor " & NewCompetitorName 
						Record_PR_Activity txtInternalRecordIdentifier, Description, Session("UserNo")
					End If
				End If
				'*****************************************
				'End Audit Trail Entries
				'*****************************************
				
				
			End If
		
	
	End If
	
	Set rsProspectCompetitorCheck = Nothing
	cnnProspectCompetitorCheck.Close
	Set cnnProspectCompetitorCheck = Nothing


'*******************************************************************************************************************


'*******************************************************************************************************************
'PERFORM AUDIT LOG UPDATE ENTRIES
'*******************************************************************************************************************

	If cInt(ORIG_TelemarketerUserNo) = 0 AND cInt(txtTelemarketerUserNo) <> 0 AND (cInt(ORIG_TelemarketerUserNo) <> cInt(txtTelemarketerUserNo)) Then
	
		Description = "The telemarketer for prospect " & ProspectName  & " was changed to <strong><em>" & GetUserDisplayNameByUserNo(txtTelemarketerUserNo) & "</em></strong> from <strong><em>No Telemarketer Selected</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect telemarketer changed",GetTerm("Prospecting") & " prospect telemarketer changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	ElseIf cInt(ORIG_TelemarketerUserNo) <> 0 AND cInt(txtTelemarketerUserNo) = 0 AND (cInt(ORIG_TelemarketerUserNo) <> cInt(txtTelemarketerUserNo)) Then
		Description = "The telemarketer for prospect " & ProspectName  & " was changed to <strong><em>No Telemarketer Selected</em></strong> from <strong><em>" & GetUserDisplayNameByUserNo(ORIG_TelemarketerUserNo) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect telemarketer changed",GetTerm("Prospecting") & " prospect telemarketer changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	ElseIf cInt(ORIG_TelemarketerUserNo) <> cInt(txtTelemarketerUserNo) Then
		Description = "The telemarketer for prospect " & ProspectName  & " was changed to <strong><em>" & GetUserDisplayNameByUserNo(txtTelemarketerUserNo) & "</em></strong> from <strong><em>" & GetUserDisplayNameByUserNo(ORIG_TelemarketerUserNo) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect telemarketer changed",GetTerm("Prospecting") & " prospect telemarketer changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")		
	End If


	If cInt(ORIG_LeadSourceNumber) = 0 AND cInt(txtLeadSourceNumber) <> 0 AND (cInt(ORIG_LeadSourceNumber) <> cInt(txtLeadSourceNumber)) Then
	
		Description = "The lead source for prospect " & ProspectName  & " was changed to <strong><em>" & GetLeadSourceByNum(txtLeadSourceNumber) & "</em></strong> from <strong><em>No Telemarketer Selected</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect lead source changed",GetTerm("Prospecting") & " prospect lead source changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	ElseIf cInt(ORIG_LeadSourceNumber) <> 0 AND cInt(txtLeadSourceNumber) = 0 AND (cInt(ORIG_LeadSourceNumber) <> cInt(txtLeadSourceNumber)) Then
		Description = "The lead source for prospect " & ProspectName  & " was changed to <strong><em>No Telemarketer Selected</em></strong> from <strong><em>" & GetLeadSourceByNum(ORIG_LeadSourceNumber) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect lead source changed",GetTerm("Prospecting") & " prospect lead source changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")
				
	ElseIf cInt(ORIG_LeadSourceNumber) <> cInt(txtLeadSourceNumber) Then
		Description = "The lead source for prospect " & ProspectName  & " was changed to <strong><em>" & GetLeadSourceByNum(txtLeadSourceNumber) & "</em></strong> from <strong><em>" & GetLeadSourceByNum(ORIG_LeadSourceNumber) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect lead source changed",GetTerm("Prospecting") & " prospect lead source changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")		
	End If
	
	
	
	If ORIG_FormerCustNum = "" OR IsNull(ORIG_FormerCustNum) Then ORIG_FormerCustNum = "NONE ENTERED"
	If txtFormerCustomerNumber = "" Then txtFormerCustomerNumber = "NONE ENTERED"
		
	If ORIG_FormerCustNum <> txtFormerCustomerNumber Then

		Description = "The former customer number for prospect " & ProspectName  & " was changed to <strong><em>" & txtFormerCustomerNumber & "</em></strong> from <strong><em>" & ORIG_FormerCustNum & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
		CreateAuditLogEntry GetTerm("Prospecting") & " prospect former customer number changed",GetTerm("Prospecting") & " prospect former customer number changed","Major",0,Description
		Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")

	End If
	

	'Response.write("ORIG_FormerCustomerCancelDate: " & ORIG_FormerCustomerCancelDate & "<br>")
	'Response.write("txtFormerCustomerCancelDate : " & txtFormerCustomerCancelDate & "<br>")


	If ORIG_FormerCustomerCancelDate = "" OR IsNull(ORIG_FormerCustomerCancelDate) Then ORIG_FormerCustomerCancelDate = "NONE ENTERED"
	If txtFormerCustomerCancelDate = "" Then txtFormerCustomerCancelDate = "NONE ENTERED"
	
	If ORIG_FormerCustomerCancelDate <> "" AND ORIG_FormerCustomerCancelDate <> "NONE ENTERED" Then
		If cDate(ORIG_FormerCustomerCancelDate) = "1/1/1900" Then
			ORIG_FormerCustomerCancelDate = "NONE ENTERED"
		End If
	End If


	If ORIG_FormerCustomerCancelDate <> txtFormerCustomerCancelDate Then
		
		If ORIG_FormerCustomerCancelDate <> "NONE ENTERED" AND txtFormerCustomerCancelDate <> "NONE ENTERED" Then
		
			If DateDiff("d",cDate(ORIG_FormerCustomerCancelDate),cDate(txtFormerCustomerCancelDate)) <> 0 Then
				Description = "The former customer cancel date for prospect " & ProspectName  & " was changed to <strong><em>" & formatDateTime(txtFormerCustomerCancelDate,2) & "</em></strong> from <strong><em>" & formatDateTime(ORIG_FormerCustomerCancelDate,2) & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
				CreateAuditLogEntry GetTerm("Prospecting") & " prospect former customer cancel date changed",GetTerm("Prospecting") & " prospect former customer cancel date changed","Major",0,Description
				Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")	
			End If
			
		Else
			Description = "The former customer cancel date for prospect " & ProspectName  & " was changed to <strong><em>" & txtFormerCustomerCancelDate & "</em></strong> from <strong><em>" & ORIG_FormerCustomerCancelDate & "</em></strong> by " & GetUserDisplayNameByUserNo(Session("UserNo")) 
			CreateAuditLogEntry GetTerm("Prospecting") & " prospect former customer cancel date changed",GetTerm("Prospecting") & " prospect former customer cancel date changed","Major",0,Description
			Record_PR_Activity txtInternalRecordIdentifier,Description,Session("UserNo")

		End If
		
	End If


'*******************************************************************************************************************


Response.Redirect ("viewProspectDetail.asp?i=" & txtInternalRecordIdentifier)

%>