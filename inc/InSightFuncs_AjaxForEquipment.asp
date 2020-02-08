<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<!--#include file="InsightFuncs_Equipment.asp"-->
<!--#include file="mail.asp"-->

<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
'Func RandomAssetTagString()
'Sub AssetTag1ExistsInDB(passedEquipIntRecID)
'Sub GenerateAssetTagForEquipment()
'Sub StatusCodeExistsInDB()
'Sub InsertStatusCodeIntoDB()
'Sub ConditionCodeExistsInDB()
'Sub InsertConditionCodeIntoDB()
'Sub MovementCodeExistsInDB()
'Sub InsertMovementCodeIntoDB()
'Sub AcquisitionCodeExistsInDB()
'Sub InsertAcquisitionCodeIntoDB()
'Sub GetInsightAssetTagByEquipAndModelIntRecID()
'***************************************************
'End List of all the AJAX functions & subs
'***************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'ALL AJAX MODAL SUBROUTINES AND FUNCTIONS BELOW THIS AREA

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

action = Request("action")

Select Case action

	Case "GenerateAssetTagForEquipment" 
		GenerateAssetTagForEquipment()	
		
	Case "AssetTag1ExistsInDB" 
		AssetTag1ExistsInDB()	
	Case "StatusCodeExistsInDB" 
		StatusCodeExistsInDB()	
	Case "ConditionCodeExistsInDB" 
		ConditionCodeExistsInDB()	
	Case "AcquisitionCodeExistsInDB" 
		AcquisitionCodeExistsInDB()	
	Case "MovementCodeExistsInDB" 
		MovementCodeExistsInDB()	

	Case "InsertStatusCodeIntoDB"
		InsertStatusCodeIntoDB()		
	Case "InsertConditionCodeIntoDB"
		InsertConditionCodeIntoDB()	
	Case "InsertMovementCodeIntoDB"
		InsertMovementCodeIntoDB()	
	Case "InsertAcquisitionCodeIntoDB"
		InsertAcquisitionCodeIntoDB()
	Case "GetInsightAssetTagByEquipAndModelIntRecID"
		GetInsightAssetTagByEquipAndModelIntRecID	
							
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub AcquisitionCodeExistsInDB()

	AcquisitionCode = Request.Form("ac")
	AcquisitionCodeDesc = Request.Form("acd")
	
	AcquisitionCodeDesc = Replace(AcquisitionCodeDesc, "'", "''")

	Set cnnAcquisitionCodeExistsInDB = Server.CreateObject("ADODB.Connection")
	cnnAcquisitionCodeExistsInDB.open Session("ClientCnnString")

	resultAcquisitionCodeExistsInDB = "False"
	
	
	'***********************************
	'CHECK IF JUST Acquisition CODE EXISTS
	'***********************************
		
	SQLAcquisitionCodeExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_AcquisitionCodes WHERE AcquisitionCode = '" & AcquisitionCode & "'"
	Set rsAcquisitionCodeExistsInDB = Server.CreateObject("ADODB.Recordset")
	rsAcquisitionCodeExistsInDB.CursorLocation = 3 
	rsAcquisitionCodeExistsInDB.Open SQLAcquisitionCodeExistsInDB,cnnAcquisitionCodeExistsInDB 

	If NOT rsAcquisitionCodeExistsInDB.EOF Then
		If Not ISNULL(rsAcquisitionCodeExistsInDB("Expr1")) Then
			If rsAcquisitionCodeExistsInDB("Expr1") < 1 Then resultAcquisitionCodeExistsInDB = "False" Else resultAcquisitionCodeExistsInDB = "ACQUISITIONCODE"
		Else
			resultAcquisitionCodeExistsInDB = "False"
		End If
	Else
		resultAcquisitionCodeExistsInDB = "False"
	End If
	
	
	'*************************************************
	'CHECK IF JUST Acquisition CODE DESCRIPTION EXISTS
	'*************************************************
	
	If resultAcquisitionCodeExistsInDB = "False" Then
			
		SQLAcquisitionCodeExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_AcquisitionCodes WHERE AcquisitionDesc = '" & AcquisitionCodeDesc & "'"
		Set rsAcquisitionCodeExistsInDB = Server.CreateObject("ADODB.Recordset")
		rsAcquisitionCodeExistsInDB.CursorLocation = 3 
		rsAcquisitionCodeExistsInDB.Open SQLAcquisitionCodeExistsInDB,cnnAcquisitionCodeExistsInDB 
	
		If NOT rsAcquisitionCodeExistsInDB.EOF Then
			If Not ISNULL(rsAcquisitionCodeExistsInDB("Expr1")) Then
				If rsAcquisitionCodeExistsInDB("Expr1") < 1 Then resultAcquisitionCodeExistsInDB = "False" Else resultAcquisitionCodeExistsInDB = "ACQUISITIONCODEDESC"
			Else
				resultAcquisitionCodeExistsInDB = "False"
			End If
		Else
			resultAcquisitionCodeExistsInDB = "False"
		End If
	
	End If
	
	'****************************************************
	'CHECK IF BOTH Acquisition CODE AND DESCRIPTION EXIST
	'****************************************************
	
	If resultAcquisitionCodeExistsInDB = "False" Then
	
		SQLAcquisitionCodeExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_AcquisitionCodes WHERE AcquisitionCode = '" & AcquisitionCodeDesc & "' AND AcquisitionDesc = '" & AcquisitionCodeDesc & "'"
		Set rsAcquisitionCodeExistsInDB = Server.CreateObject("ADODB.Recordset")
		rsAcquisitionCodeExistsInDB.CursorLocation = 3 
		rsAcquisitionCodeExistsInDB.Open SQLAcquisitionCodeExistsInDB,cnnAcquisitionCodeExistsInDB 
	
		If NOT rsAcquisitionCodeExistsInDB.EOF Then
			If Not ISNULL(rsAcquisitionCodeExistsInDB("Expr1")) Then
				If rsAcquisitionCodeExistsInDB("Expr1") < 1 Then resultAcquisitionCodeExistsInDB = "False" Else resultAcquisitionCodeExistsInDB = "BOTH"
			Else
				resultAcquisitionCodeExistsInDB = "False"
			End If
		Else
			resultAcquisitionCodeExistsInDB = "False"
		End If

	End If
	
	
	rsAcquisitionCodeExistsInDB.Close
	set rsAcquisitionCodeExistsInDB = Nothing
	cnnAcquisitionCodeExistsInDB.Close	
	set cnnAcquisitionCodeExistsInDB = Nothing
	
	Response.Write(resultAcquisitionCodeExistsInDB)
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub InsertAcquisitionCodeIntoDB()

	AcquisitionCode = Request.Form("ac")
	AcquisitionCodeDesc = Request.Form("acd")
	AcquisitionCodeDesc = Replace(AcquisitionCodeDesc, "'", "''")
		
	Set cnnInsertAcquisitionCodeIntoDB = Server.CreateObject("ADODB.Connection")
	cnnInsertAcquisitionCodeIntoDB.open Session("ClientCnnString")
	
	
	'******************************************************
	'FIRST PERFORM THE INSERT INTO THE Acquisition CODES TABLE
	'******************************************************

	SQLInsertAcquisitionCodeIntoDB = "INSERT INTO EQ_AcquisitionCodes (AcquisitionCode,AcquisitionDesc)"
	SQLInsertAcquisitionCodeIntoDB = SQLInsertAcquisitionCodeIntoDB &  " VALUES (" 
	SQLInsertAcquisitionCodeIntoDB = SQLInsertAcquisitionCodeIntoDB & "'" & AcquisitionCode & "','" & AcquisitionCodeDesc & "')"

	Set rsInsertAcquisitionCodeIntoDB = Server.CreateObject("ADODB.Recordset")
	rsInsertAcquisitionCodeIntoDB.CursorLocation = 3 
	
	rsInsertAcquisitionCodeIntoDB.Open SQLInsertAcquisitionCodeIntoDB,cnnInsertAcquisitionCodeIntoDB 
	
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " Acquisition Code: " & AcquisitionCode & " (" & AcquisitionCodeDesc & ")"
	CreateAuditLogEntry GetTerm("Equipment") & " Acquisition Code Added",GetTerm("Equipment") & " Acquisition Code Added","Minor",0,Description
		
	
	'******************************************************
	'THEN GET THE INTRECID OF THE NEW Acquisition CODE
	'******************************************************
	
	SQLInsertAcquisitionCodeIntoDB = "SELECT MAX(InternalRecordIdentifier) AS EXPR1 FROM EQ_AcquisitionCodes"
		 
	Set rsInsertAcquisitionCodeIntoDB = Server.CreateObject("ADODB.Recordset")
	rsInsertAcquisitionCodeIntoDB.CursorLocation = 3 
	rsInsertAcquisitionCodeIntoDB.Open SQLInsertAcquisitionCodeIntoDB,cnnInsertAcquisitionCodeIntoDB 

	If Not rsInsertAcquisitionCodeIntoDB.EOF Then
		newIntRecID = rsInsertAcquisitionCodeIntoDB("EXPR1")
	Else
		newIntRecID = "Error"
	End If
	

	rsInsertAcquisitionCodeIntoDB.Close
	set rsInsertAcquisitionCodeIntoDB = Nothing
	cnnInsertAcquisitionCodeIntoDB.Close	
	set cnnInsertAcquisitionCodeIntoDB = Nothing
	
	
	Response.Write(newIntRecID)
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub MovementCodeExistsInDB()

	MovementCode = Request.Form("mc")
	MovementCodeDesc = Request.Form("mcd")
	
	MovementCodeDesc = Replace(MovementCodeDesc, "'", "''")

	Set cnnMovementCodeExistsInDB = Server.CreateObject("ADODB.Connection")
	cnnMovementCodeExistsInDB.open Session("ClientCnnString")

	resultMovementCodeExistsInDB = "False"
	
	
	'***********************************
	'CHECK IF JUST MOVEMENT CODE EXISTS
	'***********************************
		
	SQLMovementCodeExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_MovementCodes WHERE movementCode = '" & MovementCode & "'"
	Set rsMovementCodeExistsInDB = Server.CreateObject("ADODB.Recordset")
	rsMovementCodeExistsInDB.CursorLocation = 3 
	rsMovementCodeExistsInDB.Open SQLMovementCodeExistsInDB,cnnMovementCodeExistsInDB 

	If NOT rsMovementCodeExistsInDB.EOF Then
		If Not ISNULL(rsMovementCodeExistsInDB("Expr1")) Then
			If rsMovementCodeExistsInDB("Expr1") < 1 Then resultMovementCodeExistsInDB = "False" Else resultMovementCodeExistsInDB = "MOVEMENTCODE"
		Else
			resultMovementCodeExistsInDB = "False"
		End If
	Else
		resultMovementCodeExistsInDB = "False"
	End If
	
	
	'*************************************************
	'CHECK IF JUST MOVEMENT CODE DESCRIPTION EXISTS
	'*************************************************
		
	SQLMovementCodeExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_MovementCodes WHERE movementDesc = '" & MovementCodeDesc & "'"
	Set rsMovementCodeExistsInDB = Server.CreateObject("ADODB.Recordset")
	rsMovementCodeExistsInDB.CursorLocation = 3 
	rsMovementCodeExistsInDB.Open SQLMovementCodeExistsInDB,cnnMovementCodeExistsInDB 

	If NOT rsMovementCodeExistsInDB.EOF Then
		If Not ISNULL(rsMovementCodeExistsInDB("Expr1")) Then
			If rsMovementCodeExistsInDB("Expr1") < 1 Then resultMovementCodeExistsInDB = "False" Else resultMovementCodeExistsInDB = "MOVEMENTCODEDESC"
		Else
			resultMovementCodeExistsInDB = "False"
		End If
	Else
		resultMovementCodeExistsInDB = "False"
	End If
	
	
	'****************************************************
	'CHECK IF BOTH MOVEMENT CODE AND DESCRIPTION EXIST
	'****************************************************
	
	If resultMovementCodeExistsInDB = "False" Then
	
		SQLMovementCodeExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_MovementCodes WHERE movementCode = '" & MovementCodeDesc & "' AND movementDesc = '" & MovementCodeDesc & "'"
		Set rsMovementCodeExistsInDB = Server.CreateObject("ADODB.Recordset")
		rsMovementCodeExistsInDB.CursorLocation = 3 
		rsMovementCodeExistsInDB.Open SQLMovementCodeExistsInDB,cnnMovementCodeExistsInDB 
	
		If NOT rsMovementCodeExistsInDB.EOF Then
			If Not ISNULL(rsMovementCodeExistsInDB("Expr1")) Then
				If rsMovementCodeExistsInDB("Expr1") < 1 Then resultMovementCodeExistsInDB = "False" Else resultMovementCodeExistsInDB = "BOTH"
			Else
				resultMovementCodeExistsInDB = "False"
			End If
		Else
			resultMovementCodeExistsInDB = "False"
		End If

	End If
	
	
	rsMovementCodeExistsInDB.Close
	set rsMovementCodeExistsInDB = Nothing
	cnnMovementCodeExistsInDB.Close	
	set cnnMovementCodeExistsInDB = Nothing
	
	Response.Write(resultMovementCodeExistsInDB)
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************

'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub InsertMovementCodeIntoDB()

	MovementCode = Request.Form("mc")
	MovementCodeDesc = Request.Form("mcd")
	MovementCodeDesc = Replace(MovementCodeDesc, "'", "''")
		
	Set cnnInsertMovementCodeIntoDB = Server.CreateObject("ADODB.Connection")
	cnnInsertMovementCodeIntoDB.open Session("ClientCnnString")
	
	
	'******************************************************
	'FIRST PERFORM THE INSERT INTO THE MOVEMENT CODES TABLE
	'******************************************************

	SQLInsertMovementCodeIntoDB = "INSERT INTO EQ_MovementCodes (movementCode,movementDesc)"
	SQLInsertMovementCodeIntoDB = SQLInsertMovementCodeIntoDB &  " VALUES (" 
	SQLInsertMovementCodeIntoDB = SQLInsertMovementCodeIntoDB & "'" & MovementCode & "','" & MovementCodeDesc & "')"

	Set rsInsertMovementCodeIntoDB = Server.CreateObject("ADODB.Recordset")
	rsInsertMovementCodeIntoDB.CursorLocation = 3 
	
	rsInsertMovementCodeIntoDB.Open SQLInsertMovementCodeIntoDB,cnnInsertMovementCodeIntoDB 
	
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " Movement Code: " & MovementCode & " (" & MovementCodeDesc & ")"
	CreateAuditLogEntry GetTerm("Equipment") & " Movement Code Added",GetTerm("Equipment") & " Movement Code Added","Minor",0,Description
		
	
	'******************************************************
	'THEN GET THE INTRECID OF THE NEW MOVEMENT CODE
	'******************************************************
	
	SQLInsertMovementCodeIntoDB = "SELECT MAX(InternalRecordIdentifier) AS EXPR1 FROM EQ_MovementCodes"
		 
	Set rsInsertMovementCodeIntoDB = Server.CreateObject("ADODB.Recordset")
	rsInsertMovementCodeIntoDB.CursorLocation = 3 
	rsInsertMovementCodeIntoDB.Open SQLInsertMovementCodeIntoDB,cnnInsertMovementCodeIntoDB 

	If Not rsInsertMovementCodeIntoDB.EOF Then
		newIntRecID = rsInsertMovementCodeIntoDB("EXPR1")
	Else
		newIntRecID = "Error"
	End If
	

	rsInsertMovementCodeIntoDB.Close
	set rsInsertMovementCodeIntoDB = Nothing
	cnnInsertMovementCodeIntoDB.Close	
	set cnnInsertMovementCodeIntoDB = Nothing
	
	
	Response.Write(newIntRecID)
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub ConditionCodeExistsInDB()

	ConditionCode = Request.Form("cc")
	ConditionCodeDesc = Request.Form("ccd")

	Set cnnConditionCodeExistsInDB = Server.CreateObject("ADODB.Connection")
	cnnConditionCodeExistsInDB.open Session("ClientCnnString")

	resultConditionCodeExistsInDB = "False"
	
	
	'***********************************
	'CHECK IF CONDITION CODE EXISTS
	'***********************************
		
	SQLConditionCodeExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_Condition WHERE Condition = '" & ConditionCode & "'"
	Set rsConditionCodeExistsInDB = Server.CreateObject("ADODB.Recordset")
	rsConditionCodeExistsInDB.CursorLocation = 3 
	rsConditionCodeExistsInDB.Open SQLConditionCodeExistsInDB,cnnConditionCodeExistsInDB 

	If NOT rsConditionCodeExistsInDB.EOF Then
		If Not ISNULL(rsConditionCodeExistsInDB("Expr1")) Then
			If rsConditionCodeExistsInDB("Expr1") < 1 Then resultConditionCodeExistsInDB = "False" Else resultConditionCodeExistsInDB = "CONDITIONCODE"
		Else
			resultConditionCodeExistsInDB = "False"
		End If
	Else
		resultConditionCodeExistsInDB = "False"
	End If
		
	
	rsConditionCodeExistsInDB.Close
	set rsConditionCodeExistsInDB = Nothing
	cnnConditionCodeExistsInDB.Close	
	set cnnConditionCodeExistsInDB = Nothing
	
	Response.Write(resultConditionCodeExistsInDB)
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub InsertConditionCodeIntoDB()

	ConditionCode = Request.Form("cc")
	ConditionCodeDesc = Request.Form("ccd")

	ConditionCodeDesc = Replace(ConditionCodeDesc, "'", "''")
			
	Set cnnInsertConditionCodeIntoDB = Server.CreateObject("ADODB.Connection")
	cnnInsertConditionCodeIntoDB.open Session("ClientCnnString")
	
	
	'**********************************************************
	'FIRST PERFORM THE INSERT INTO THE CONDITION CODES TABLE
	'**********************************************************

	SQLInsertConditionCodeIntoDB = "INSERT INTO EQ_Condition (Condition,Description,RecordSource)"
	SQLInsertConditionCodeIntoDB = SQLInsertConditionCodeIntoDB &  " VALUES (" 
	SQLInsertConditionCodeIntoDB = SQLInsertConditionCodeIntoDB & "'" & ConditionCode & "','" & ConditionCodeDesc & "','Insight')"

	Set rsInsertConditionCodeIntoDB = Server.CreateObject("ADODB.Recordset")
	rsInsertConditionCodeIntoDB.CursorLocation = 3 
	
	rsInsertConditionCodeIntoDB.Open SQLInsertConditionCodeIntoDB,cnnInsertConditionCodeIntoDB 
	
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " condition: " & ConditionCode 
	CreateAuditLogEntry GetTerm("Equipment") & " Condition Code Added",GetTerm("Equipment") & " Condition Code Added","Minor",0,Description
	
	
	'******************************************************
	'THEN GET THE INTRECID OF THE NEW CONDITION CODE
	'******************************************************
	
	SQLInsertConditionCodeIntoDB = "SELECT MAX(InternalRecordIdentifier) AS EXPR1 FROM EQ_Condition"
		 
	Set rsInsertConditionCodeIntoDB = Server.CreateObject("ADODB.Recordset")
	rsInsertConditionCodeIntoDB.CursorLocation = 3 
	rsInsertConditionCodeIntoDB.Open SQLInsertConditionCodeIntoDB,cnnInsertConditionCodeIntoDB 

	If Not rsInsertConditionCodeIntoDB.EOF Then
		newIntRecID = rsInsertConditionCodeIntoDB("EXPR1")
	Else
		newIntRecID = "Error"
	End If
	

	rsInsertConditionCodeIntoDB.Close
	set rsInsertConditionCodeIntoDB = Nothing
	cnnInsertConditionCodeIntoDB.Close	
	set cnnInsertConditionCodeIntoDB = Nothing
	
	
	Response.Write(newIntRecID)
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************






'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub StatusCodeExistsInDB()

	statusCodeDesc = Request.Form("sc")
	statusBackendCode = Request.Form("bsc")

	Set cnnStatusCodeExistsInDB = Server.CreateObject("ADODB.Connection")
	cnnStatusCodeExistsInDB.open Session("ClientCnnString")

	resultStatusCodeExistsInDB = "False"
	
	
	'***********************************
	'CHECK IF JUST STATUS CODE EXISTS
	'***********************************
		
	SQLStatusCodeExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_StatusCodes WHERE statusDesc = '" & statusCodeDesc & "'"
	Set rsStatusCodeExistsInDB = Server.CreateObject("ADODB.Recordset")
	rsStatusCodeExistsInDB.CursorLocation = 3 
	rsStatusCodeExistsInDB.Open SQLStatusCodeExistsInDB,cnnStatusCodeExistsInDB 

	If NOT rsStatusCodeExistsInDB.EOF Then
		If Not ISNULL(rsStatusCodeExistsInDB("Expr1")) Then
			If rsStatusCodeExistsInDB("Expr1") < 1 Then resultStatusCodeExistsInDB = "False" Else resultStatusCodeExistsInDB = "STATUSCODEDESC"
		Else
			resultStatusCodeExistsInDB = "False"
		End If
	Else
		resultStatusCodeExistsInDB = "False"
	End If
	
	'****************************************************
	'CHECK IF JUST STATUS CODE NAME/DESCRIPTION EXISTS
	'****************************************************
	
	If resultStatusCodeExistsInDB = "False" Then
	
		SQLStatusCodeExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_StatusCodes WHERE statusBackendSystemCode = '" & statusBackendCode & "'"
		Set rsStatusCodeExistsInDB = Server.CreateObject("ADODB.Recordset")
		rsStatusCodeExistsInDB.CursorLocation = 3 
		rsStatusCodeExistsInDB.Open SQLStatusCodeExistsInDB,cnnStatusCodeExistsInDB 
	
		If NOT rsStatusCodeExistsInDB.EOF Then
			If Not ISNULL(rsStatusCodeExistsInDB("Expr1")) Then
				If rsStatusCodeExistsInDB("Expr1") < 1 Then resultStatusCodeExistsInDB = "False" Else resultStatusCodeExistsInDB = "BACKENDSYSTEMCODE"
			Else
				resultStatusCodeExistsInDB = "False"
			End If
		Else
			resultStatusCodeExistsInDB = "False"
		End If
	
	End If
	
	'****************************************************
	'CHECK IF BOTH STATUS CODE AND DESCRIPTION EXISTS
	'****************************************************
	
	If resultStatusCodeExistsInDB = "False" Then
	
		SQLStatusCodeExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_StatusCodes WHERE statusDesc = '" & statusCodeDesc & "' AND statusBackendSystemCode = '" & statusBackendCode & "'"
		Set rsStatusCodeExistsInDB = Server.CreateObject("ADODB.Recordset")
		rsStatusCodeExistsInDB.CursorLocation = 3 
		rsStatusCodeExistsInDB.Open SQLStatusCodeExistsInDB,cnnStatusCodeExistsInDB 
	
		If NOT rsStatusCodeExistsInDB.EOF Then
			If Not ISNULL(rsStatusCodeExistsInDB("Expr1")) Then
				If rsStatusCodeExistsInDB("Expr1") < 1 Then resultStatusCodeExistsInDB = "False" Else resultStatusCodeExistsInDB = "BOTH"
			Else
				resultStatusCodeExistsInDB = "False"
			End If
		Else
			resultStatusCodeExistsInDB = "False"
		End If

	End If
	
	
	rsStatusCodeExistsInDB.Close
	set rsStatusCodeExistsInDB = Nothing
	cnnStatusCodeExistsInDB.Close	
	set cnnStatusCodeExistsInDB = Nothing
	
	Response.Write(resultStatusCodeExistsInDB)
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub InsertStatusCodeIntoDB()

	statusCodeDesc = Request.Form("scd")
	statusBackendCode = Request.Form("bsc")
	availableForPlacement = Request.Form("afp")
	generatesRevenue = Request.Form("gr")
	
	If availableForPlacement = "on" then 
		availableForPlacement = 1 
	Else 
		availableForPlacement = 0
	End If
	
	If generatesRevenue = "on" then 
		generatesRevenue = 1 
	Else 
		generatesRevenue = 0
	End If
		
	Set cnnInsertStatusCodeIntoDB = Server.CreateObject("ADODB.Connection")
	cnnInsertStatusCodeIntoDB.open Session("ClientCnnString")
	
	
	'******************************************************
	'FIRST PERFORM THE INSERT INTO THE STATUS CODES TABLE
	'******************************************************

	SQLInsertStatusCodeIntoDB = "INSERT INTO EQ_StatusCodes (statusDesc,statusAvailableForPlacement,statusGeneratesRentalRevenue,statusBackendSystemCode,RecordSource)"
	SQLInsertStatusCodeIntoDB = SQLInsertStatusCodeIntoDB &  " VALUES (" 
	SQLInsertStatusCodeIntoDB = SQLInsertStatusCodeIntoDB & "'" & statusCodeDesc & "'," & availableForPlacement & "," & generatesRevenue & ",'" & statusBackendCode & "','Insight')"

	Set rsInsertStatusCodeIntoDB = Server.CreateObject("ADODB.Recordset")
	rsInsertStatusCodeIntoDB.CursorLocation = 3 
	
	rsInsertStatusCodeIntoDB.Open SQLInsertStatusCodeIntoDB,cnnInsertStatusCodeIntoDB 
	
	Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Equipment") & " status code: " & statusCodeDesc 
	CreateAuditLogEntry GetTerm("Equipment") & " Status Code Added",GetTerm("Equipment") & " Status Code Added","Minor",0,Description
	
	
	'******************************************************
	'THEN GET THE INTRECID OF THE NEW STATUS CODE
	'******************************************************
	
	SQLInsertStatusCodeIntoDB = "SELECT MAX(InternalRecordIdentifier) AS EXPR1 FROM EQ_StatusCodes"
		 
	Set rsInsertStatusCodeIntoDB = Server.CreateObject("ADODB.Recordset")
	rsInsertStatusCodeIntoDB.CursorLocation = 3 
	rsInsertStatusCodeIntoDB.Open SQLInsertStatusCodeIntoDB,cnnInsertStatusCodeIntoDB 

	If Not rsInsertStatusCodeIntoDB.EOF Then
		newIntRecID = rsInsertStatusCodeIntoDB("EXPR1")
	Else
		newIntRecID = "Error"
	End If
	

	rsInsertStatusCodeIntoDB.Close
	set rsInsertStatusCodeIntoDB = Nothing
	cnnInsertStatusCodeIntoDB.Close	
	set cnnInsertStatusCodeIntoDB = Nothing
	
	
	Response.Write(newIntRecID)
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub AssetTag1ExistsInDB()

	equipIntRecID = Request.Form("intRecID")
	assetTag1 = Request.Form("assetTag1")

	Set cnnAssetTag1ExistsInDB = Server.CreateObject("ADODB.Connection")
	cnnAssetTag1ExistsInDB.open Session("ClientCnnString")

	resultAssetTag1ExistsInDB = "False"
		
	SQLAssetTag1ExistsInDB = " SELECT COUNT(*) AS Expr1 FROM EQ_Equipment WHERE AssetTag1 = '" & assetTag1 & "'"
		 
	Set rsAssetTag1ExistsInDB = Server.CreateObject("ADODB.Recordset")
	rsAssetTag1ExistsInDB.CursorLocation = 3 
	
	
	rsAssetTag1ExistsInDB.Open SQLAssetTag1ExistsInDB,cnnAssetTag1ExistsInDB 

	If Not rsAssetTag1ExistsInDB.EOF Then
		If Not ISNULL(rsAssetTag1ExistsInDB("Expr1")) Then
			If rsAssetTag1ExistsInDB("Expr1") < 2 Then resultAssetTag1ExistsInDB = "False" Else resultAssetTag1ExistsInDB = "True"
		Else
			resultAssetTag1ExistsInDB = "False"
		End If
	Else
		resultAssetTag1ExistsInDB = "False"
	End If

	rsAssetTag1ExistsInDB.Close
	set rsAssetTag1ExistsInDB = Nothing
	cnnAssetTag1ExistsInDB.Close	
	set cnnAssetTag1ExistsInDB = Nothing
	
	Response.Write(resultAssetTag1ExistsInDB)
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************





'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Function RandomAssetTagString()

	all_chars = array("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","S","T","U","V","W","X","Y","Z","0","1","2","3","4","5","6","7","8","9")
	
	Randomize()
	
	for i = 1 to 16
	   	random_index = int(Rnd()*35)
		random_string = random_string & all_chars(random_index) 
	next  

    RandomAssetTagString = random_string

End Function

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GenerateAssetTagForEquipment()

	equipIntRecID = Request.Form("intRecID") 
	
	assetTagIsUnique = False

	Set cnnGenAssetTag = Server.CreateObject("ADODB.Connection")
	cnnGenAssetTag.open Session("ClientCnnString")
	Set rsGenAssetTag = Server.CreateObject("ADODB.Recordset")
	rsGenAssetTag.CursorLocation = 3 


	Do While assetTagIsUnique = False
	
		randomAssetTag = RandomAssetTagString()
		
		SQLGenAssetTag = "SELECT * FROM EQ_Equipment WHERE AssetTag1 = '" & randomAssetTag & "'"
		rsGenAssetTag.Open SQLGenAssetTag,cnnGenAssetTag 
		
		If rsGenAssetTag.EOF Then
			assetTagIsUnique = True			
		End If
	Loop
			
	rsGenAssetTag.Close
	set rsGenAssetTag = Nothing
	cnnGenAssetTag.Close	
	set cnnGenAssetTag = Nothing
	
	Response.Write(randomAssetTag)

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'**************************************************************************************************************************************
'**************************************************************************************************************************************
Sub GetInsightAssetTagByEquipAndModelIntRecID()

	passedEquipIntRecID = Request.Form("equipIntRecID")
	passedModelIntRecID = Request.Form("modelIntRecID")

	Set cnnGetInsightAssetTagByEquipIntRecID = Server.CreateObject("ADODB.Connection")
	cnnGetInsightAssetTagByEquipIntRecID.open Session("ClientCnnString")

	resultGetInsightAssetTagByEquipIntRecID = ""
	Set rsGetInsightAssetTagByEquipIntRecID  = Server.CreateObject("ADODB.Recordset")
	rsGetInsightAssetTagByEquipIntRecID.CursorLocation = 3 
		
	
	SQLGetInsightAssetTagByEquipIntRecID = "SELECT ClassIntRecID, BrandIntRecID, ManufacIntRecID, InsightAssetTagPrefix FROM EQ_Models WHERE InternalRecordIdentifier = " & passedModelIntRecID
	rsGetInsightAssetTagByEquipIntRecID.Open SQLGetInsightAssetTagByEquipIntRecID,cnnGetInsightAssetTagByEquipIntRecID 
	BrandIntRecID = rsGetInsightAssetTagByEquipIntRecID("BrandIntRecID")
	ManufacIntRecID = rsGetInsightAssetTagByEquipIntRecID("ManufacIntRecID")
	ClassIntRecID = rsGetInsightAssetTagByEquipIntRecID("ClassIntRecID")
	ModelInsightAssetTagPrefix = rsGetInsightAssetTagByEquipIntRecID("InsightAssetTagPrefix")
	rsGetInsightAssetTagByEquipIntRecID.Close
	
	
	SQLGetInsightAssetTagByEquipIntRecID = "SELECT InsightAssetTagPrefix FROM EQ_Classes WHERE InternalRecordIdentifier = " & ClassIntRecID
	rsGetInsightAssetTagByEquipIntRecID.Open SQLGetInsightAssetTagByEquipIntRecID,cnnGetInsightAssetTagByEquipIntRecID 
	ClassInsightAssetTagPrefix = rsGetInsightAssetTagByEquipIntRecID("InsightAssetTagPrefix")
	rsGetInsightAssetTagByEquipIntRecID.Close
	

	SQLGetInsightAssetTagByEquipIntRecID = "SELECT InsightAssetTagPrefix FROM EQ_Manufacturers WHERE InternalRecordIdentifier = " & ManufacIntRecID
	rsGetInsightAssetTagByEquipIntRecID.Open SQLGetInsightAssetTagByEquipIntRecID,cnnGetInsightAssetTagByEquipIntRecID 
	ManfInsightAssetTagPrefix = rsGetInsightAssetTagByEquipIntRecID("InsightAssetTagPrefix")
	rsGetInsightAssetTagByEquipIntRecID.Close
	

	SQLGetInsightAssetTagByEquipIntRecID = "SELECT InsightAssetTagPrefix FROM EQ_Brands WHERE InternalRecordIdentifier = " & BrandIntRecID
	rsGetInsightAssetTagByEquipIntRecID.Open SQLGetInsightAssetTagByEquipIntRecID,cnnGetInsightAssetTagByEquipIntRecID 
	BrandInsightAssetTagPrefix = rsGetInsightAssetTagByEquipIntRecID("InsightAssetTagPrefix")
	rsGetInsightAssetTagByEquipIntRecID.Close
	
	If ClassInsightAssetTagPrefix = "" OR IsNull(ClassInsightAssetTagPrefix) Then
		ClassInsightAssetTagPrefix = "***"
	End If
	
	If ManfInsightAssetTagPrefix = "" OR IsNull(ManfInsightAssetTagPrefix) Then 
		ManfInsightAssetTagPrefix = "***"
	End If
	
	If BrandInsightAssetTagPrefix = "" OR IsNull(BrandInsightAssetTagPrefix) Then
		BrandInsightAssetTagPrefix = "***"
	End If
	
	If ModelInsightAssetTagPrefix = "" OR IsNull(ModelInsightAssetTagPrefix) Then
		ModelInsightAssetTagPrefix = "***"
	End If
	
	InsightAssetTagPrefix = ClassInsightAssetTagPrefix & "-" & ManfInsightAssetTagPrefix & "-" & BrandInsightAssetTagPrefix & "-" & ModelInsightAssetTagPrefix & "-" & passedEquipIntRecID

	

	set rsGetInsightAssetTagByEquipIntRecID = Nothing
	cnnGetInsightAssetTagByEquipIntRecID.Close	
	set cnnGetInsightAssetTagByEquipIntRecID = Nothing
	
	If InsightAssetTagPrefix = "" OR Left(InsightAssetTagPrefix,12) ="***-***-***-***" Then InsightAssetTagPrefix = "NO INSIGHT ASSET TAG ABLE TO BE GENERATED"	
	Response.Write(InsightAssetTagPrefix)
	
End Sub
'**************************************************************************************************************************************
'**************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'END ALL AJAX MODAL SUBROUTINES AND FUNCTIONS

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

%>