<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/InSightFuncs_Prospecting.asp"-->
<% If Session("Userno") = "" Then Response.End() %>
<%
Response.ContentType = "application/json"
ProspectIntRecID = Request.QueryString("i") 
If ProspectIntRecID = "" Then Response.End()

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))

If Request.Form("updateAction")="save" Then

	'***************************************************************************************
	'Lookup the record as it exists now so we can fillin the audit trail
	'***************************************************************************************
	SQL = "SELECT * FROM PR_ProspectCompetitors WHERE ProspectRecID='"&ProspectIntRecID &"' AND CompetitorRecID='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		Orig_CompetitorRecID = rs("CompetitorRecID")
		Orig_PrimaryCompetitor = rs("PrimaryCompetitor")
		Orig_BottledWater = rs("BottledWater")
		Orig_FilteredWater = rs("FilteredWater")
		Orig_OCS = rs("OCS")
		Orig_OCS_Supply = rs("OCS_Supply")
		Orig_OfficeSupplies = rs("OfficeSupplies")
		Orig_Vending = rs("Vending")
		Orig_Micromarket = rs("Micromarket")
		Orig_Pantry = rs("Pantry")	
		Orig_Notes = rs("Notes")
	End If

'	If (Orig_PrimaryCompetitor <> "" AND Orig_PrimaryCompetitor <> 0) Then Orig_PrimaryCompetitor = 1 Else Orig_PrimaryCompetitor = 0
	If (Orig_OCS <> "" AND Orig_OCS <> 0) Then Orig_OCS = 1 Else Orig_OCS = 0
	If (Orig_OCS_Supply <> "" AND Orig_OCS_Supply <> 0) Then Orig_OCS_Supply = 1 Else Orig_OCS_Supply = 0
	If (Orig_BottledWater <> "" AND Orig_BottledWater <> 0) Then Orig_BottledWater = 1 Else Orig_BottledWater = 0
	If (Orig_FilteredWater <> "" AND Orig_FilteredWater <> 0) Then Orig_FilteredWater = 1 Else Orig_FilteredWater = 0	
	If (Orig_Vending <> "" AND Orig_Vending <> 0) Then Orig_Vending = 1 Else Orig_Vending = 0
	If (Orig_Micromarket <> "" AND Orig_Micromarket <> 0) Then Orig_Micromarket = 1 Else Orig_Micromarket = 0
	If (Orig_Pantry <> "" AND Orig_Pantry <> 0) Then Orig_Pantry = 1 Else Orig_Pantry = 0
	If (Orig_OfficeSupplies <> "" AND Orig_OfficeSupplies <> 0) Then Orig_OfficeSupplies = 1 Else Orig_OfficeSupplies = 0

	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	
	SQL = "SELECT * FROM PR_Competitors WHERE InternalRecordIdentifier = " & Orig_CompetitorRecID
		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		CompetitorName = rs("CompetitorName")
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	'***************************************************************************************
	'Perform update on record in SQL
	'***************************************************************************************

	Query = "UPDATE PR_ProspectCompetitors SET "
	
	If Request.Form("PrimaryCompetitor") = 1 Then
		Query = Query & "PrimaryCompetitor='1', "
	Else
		Query = Query & "PrimaryCompetitor='0', "
	End If
		
	Query = Query & "Notes='"&EscapeSingleQuotes(Request.Form("Notes"))&"', "

	If Request.Form("BottledWater") = 1 Then
		Query = Query & "BottledWater='1', "
	Else
		Query = Query & "BottledWater='0', "
	End If
	If Request.Form("FilteredWater") = 1 Then
		Query = Query & "FilteredWater='1', "
	Else
		Query = Query & "FilteredWater='0', "
	End If
	If Request.Form("OCS") = 1 Then
		Query = Query & "OCS='1', "
	Else
		Query = Query & "OCS='0', "
	End If
	If Request.Form("OCS_Supply") = 1 Then
		Query = Query & "OCS_Supply='1', "
	Else
		Query = Query & "OCS_Supply='0', "
	End If
	If Request.Form("OfficeSupplies") = 1 Then
		Query = Query & "OfficeSupplies='1', "
	Else
		Query = Query & "OfficeSupplies='0', "
	End If
	If Request.Form("Vending") = 1 Then
		Query = Query & "Vending='1', "
	Else
		Query = Query & "Vending='0', "
	End If
	If Request.Form("MicroMarket") = 1 Then
		Query = Query & "MicroMarket='1', "
	Else
		Query = Query & "MicroMarket='0', "
	End If
	If Request.Form("Pantry") = 1 Then
		Query = Query & "Pantry='1' "
	Else
		Query = Query & "Pantry='0' "
	End If
		
	Query = Query & "WHERE ProspectRecID='"&ProspectIntRecID &"' AND CompetitorRecID='"&Request.Form("updateActionId")&"'"
	cnn.Execute(Query)
	
	If Request.Form("PrimaryCompetitor") = 1 Then
		Query = "UPDATE PR_ProspectCompetitors SET PrimaryCompetitor= 0 WHERE ProspectRecID='"&ProspectIntRecID &"' AND CompetitorRecID <> '"&Request.Form("updateActionId")&"'"
		cnn.Execute(Query)
	End If
	

	'***************************************************************************************
	'After SQL update, record entries in audit trail
	'***************************************************************************************
	
	PrimaryCompetitor		= Request.Form("PrimaryCompetitor")
	competitorNotes			= Request.Form("Notes")
	OCS						= Request.Form("OCS")
	OCS_Supply				= Request.Form("OCS_Supply")
	BottledWater			= Request.Form("BottledWater")
	FilteredWater			= Request.Form("FilteredWater")
	Vending					= Request.Form("Vending")
	Micromarket				= Request.Form("Micromarket")
	Pantry 					= Request.Form("Pantry")
	OfficeSupplies 			= Request.Form("OfficeSupplies")

	If (PrimaryCompetitor <> "" AND PrimaryCompetitor <> 0) Then PrimaryCompetitor = 1 Else PrimaryCompetitor = 0
	If (OCS <> "" AND OCS <> 0) Then OCS = 1 Else OCS = 0
	If (OCS_Supply <> "" AND OCS_Supply <> 0) Then OCS_Supply = 1 Else OCS_Supply = 0
	If (BottledWater <> "" AND BottledWater <> 0) Then BottledWater = 1 Else BottledWater = 0
	If (FilteredWater <> "" AND FilteredWater <> 0) Then FilteredWater = 1 Else FilteredWater = 0	
	If (Vending <> "" AND Vending <> 0) Then Vending = 1 Else Vending = 0
	If (Micromarket <> "" AND Micromarket <> 0) Then Micromarket = 1 Else Micromarket = 0
	If (Pantry <> "" AND Pantry <> 0) Then Pantry = 1 Else Pantry = 0
	If (OfficeSupplies <> "" AND OfficeSupplies <> 0) Then OfficeSupplies = 1 Else OfficeSupplies = 0
		
	
	'***********************************************************************
	'End Lookup the record as it exists now so we can fillin the audit trail
	'***********************************************************************

	Description = ""
	
	If Orig_PrimaryCompetitor <> PrimaryCompetitor Then
	
		If (PrimaryCompetitor  = 1 OR PrimaryCompetitor  = vbTrue) Then
			Description = CompetitorName & " was set to be the primary competitor for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " primary competitor change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = " The primary competitor was changed to " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = CompetitorName & " was un-set as the primary competitor for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " primary competitor change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = CompetitorName & " is no longer the primary competitor."
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If
	End If
	
	
	If Orig_Notes  <> competitorNotes Then
	
		Description =  "Competitor notes changed from " & Orig_Notes  & " to " & competitorNotes & " for competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " competitor notes change ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The notes for the competitor " & CompetitorName & " changed from: <em><strong> " & Orig_Notes  & "</em></strong> to: <em><strong>" & competitorNotes & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	
	

	If Orig_BottledWater <> BottledWater Then
	
		If (BottledWater = 1 OR BottledWater = vbTrue) Then
			Description = "Bottled water was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Bottled water was added as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = "Bottled water was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Bottled water was removed as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If

	End If
	
	
	
	
	
	If Orig_FilteredWater <> FilteredWater Then
	
		If (FilteredWater = 1 OR FilteredWater = vbTrue) Then
			Description = "Filtered water was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Filtered water was added as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = "Filtered water was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Filtered water was removed as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If

	End If
	
	
	If Orig_OCS <> OCS Then
	
		If (OCS = 1 OR OCS = vbTrue) Then
			Description = "OCS was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "OCS was added as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = "OCS was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "OCS was removed as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If

	End If
	
	
	If Orig_OCS_Supply <> OCS_Supply Then
	
		If (OCS_Supply = 1 OR OCS_Supply = vbTrue) Then
			Description = "OCS Supply was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "OCS Supply was added as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = "OCS Supply was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "OCS Supply was removed as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If

	End If


	If Orig_OfficeSupplies <> OfficeSupplies Then
	
		If (OfficeSupplies = 1 OR OfficeSupplies = vbTrue) Then
			Description = "Office Supplies was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Office Supplies was added as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = "Office Supplies was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Office Supplies was removed as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If

	End If
	
	
	If Orig_Vending <> Vending Then
	
		If (Vending = 1 OR Vending = vbTrue) Then
			Description = "Vending was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Vending was added as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = "Vending was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Vending was removed as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If

	End If
	
	
	If Orig_Micromarket <> Micromarket Then
	
		If (Micromarket = 1 OR Micromarket = vbTrue) Then
			Description = "Micromarket was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Micromarket was added as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = "Micromarket was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Micromarket was removed as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If

	End If
	
	
	If Orig_Pantry <> Pantry Then
	
		If (Pantry = 1 OR Pantry = vbTrue) Then
			Description = "Pantry was added as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Pantry was added as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		Else
			Description = "Pantry was removed as a competitor category being offered by the competitor " & CompetitorName & " for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
			CreateAuditLogEntry GetTerm("Prospecting")& " competitor category offering change ",GetTerm("Prospecting"),"Minor",0,Description
			
			Description = "Pantry was removed as a competitor category being offered by the competitor " & CompetitorName 
			Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
		End If
	End If
	
	

		
End If







If Request.Form("updateAction")="insert" Then

	If Request.Form("PrimaryCompetitor") = 1 Then
		PrimaryCompetitor=1
	Else
		PrimaryCompetitor=0
	End If
	
	If Request.Form("BottledWater") = 1 Then
		BottledWater=1
	Else
		BottledWater=0
	End If
	If Request.Form("FilteredWater") = 1 Then
		FilteredWater=1
	Else
		FilteredWater=0
	End If
	If Request.Form("OCS") = 1 Then
		OCS=1
	Else
		OCS=0
	End If
	If Request.Form("OCS_Supply") = 1 Then
		OCS_Supply=1
	Else
		OCS_Supply=0
	End If
	If Request.Form("OfficeSupplies") = 1 Then
		OfficeSupplies=1
	Else
		OfficeSupplies=0
	End If
	If Request.Form("Vending") = 1 Then
		Vending=1
	Else
		Vending=0
	End If
	If Request.Form("MicroMarket") = 1 Then
		MicroMarket=1
	Else
		MicroMarket=0
	End If	
	If Request.Form("Pantry") = 1 Then
		Pantry=1
	Else
		Pantry=0
	End If

	Notes= EscapeSingleQuotes(Request.Form("Notes"))
	CompetitorRecID = Request.Form("CompetitorRecID")

	Query = "INSERT INTO PR_ProspectCompetitors (ProspectRecID, CompetitorRecID, PrimaryCompetitor, Notes, OCS, OCS_Supply, BottledWater, FilteredWater, OfficeSupplies, Vending, Micromarket, Pantry) "
	Query = Query & " VALUES "
	Query = Query & "(" & ProspectIntRecID & ",'" & CompetitorRecID & "'," & PrimaryCompetitor &",'" & Notes & "', "
	Query = Query & OCS & "," & OCS_Supply & "," & BottledWater & "," & FilteredWater & "," & OfficeSupplies & "," & Vending & "," & Micromarket & "," & Pantry & ") "
	cnn.Execute(Query)
	
	If Request.Form("PrimaryCompetitor") = 1 Then
		Query = "UPDATE PR_ProspectCompetitors SET PrimaryCompetitor= 0 WHERE ProspectRecID='"& ProspectIntRecID &"' AND CompetitorRecID <> '"& CompetitorRecID &"'"
		cnn.Execute(Query)
	End If

	If (Request.Form("PrimaryCompetitor") = 1 OR Request.Form("PrimaryCompetitor") = vbTrue) Then
	
		Description = GetCompetitorByNum(CompetitorRecID) & " was added to the prospect " & GetProspectNameByNumber(ProspectIntRecID) & " and set to be the primary competitor"
		CreateAuditLogEntry GetTerm("Prospecting") & " competitor added ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = GetCompetitorByNum(CompetitorRecID) & " was added to this prospect and set to be the primary competitor"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
	Else
	
		Description = GetCompetitorByNum(CompetitorRecID) & " was added to the prospect " & GetProspectNameByNumber(ProspectIntRecID) & " as a competitor."
		CreateAuditLogEntry GetTerm("Prospecting") & " competitor added ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = GetCompetitorByNum(CompetitorRecID) & " was added to this prospect as a competitor."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
	End If

	
	
	
End If




If Request.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM PR_ProspectCompetitors WHERE ProspectRecID='"&ProspectIntRecID &"' AND CompetitorRecID='"&Request.Form("updateActionId")&"'"		
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)
		
	If not rs.EOF Then
		PrimaryCompetitor = rs("PrimaryCompetitor")
	End If
	
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing


	If (PrimaryCompetitor = 1 OR PrimaryCompetitor = vbTrue) Then
	
		Description = "The primary competitor " & GetCompetitorByNum(Request.Form("updateActionId")) & " was removed from the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " competitor removed from prospect ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The primary competitor " & GetCompetitorByNum(Request.Form("updateActionId")) & " was removed from this prospect "
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
	Else
	
		Description = "The competitor " & GetCompetitorByNum(Request.Form("updateActionId")) & " was removed from the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Prospecting") & " competitor removed from prospect ",GetTerm("Prospecting"),"Minor",0,Description
		
		Description = "The competitor " &  GetCompetitorByNum(Request.Form("updateActionId")) & " was removed from this prospect."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
	End If


	Query = "DELETE FROM PR_ProspectCompetitors WHERE ProspectRecID= '" & ProspectIntRecID & "' AND CompetitorRecID = '" & Request.Form("updateActionId") & "'"
	cnn.Execute(Query)
	
End If





Query = "SELECT * FROM PR_ProspectCompetitors WHERE ProspectRecID =' " & ProspectIntRecID & "' ORDER BY PrimaryCompetitor DESC, CompetitorRecID"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn.Execute(Query)

Response.Write("[")

If not rs.EOF Then
	sep = ""
	Do While Not rs.EOF
	
			CompetitorRecID = rs("CompetitorRecID")
			PrimaryCompetitor = rs("PrimaryCompetitor")
			competitorNotes = rs("Notes")
			BottledWater = rs("BottledWater")
			FilteredWater = rs("FilteredWater")
			OCS = rs("OCS")
			OCS_Supply = rs("OCS_Supply")
			OfficeSupplies = rs("OfficeSupplies")
			Vending = rs("Vending")
			MicroMarket = rs("MicroMarket")
			Pantry = rs("Pantry")
			
			SQLCompetitorsInner = "SELECT * FROM PR_Competitors WHERE InternalRecordIdentifier = " & CompetitorRecID 
			
			Set cnnCompetitorsInner = Server.CreateObject("ADODB.Connection")
			cnnCompetitorsInner.open (Session("ClientCnnString"))
			Set rsCompetitorsInner = Server.CreateObject("ADODB.Recordset")
			rsCompetitorsInner.CursorLocation = 3 
			Set rsCompetitorsInner = cnnCompetitorsInner.Execute(SQLCompetitorsInner)
			
			If not rsCompetitorsInner.EOF Then
			
				Do While Not rsCompetitorsInner.EOF
			
					CompInternalRecordIdentifier = rsCompetitorsInner("InternalRecordIdentifier")
					CompetitorName = rsCompetitorsInner("CompetitorName")
					AddressInformation = rsCompetitorsInner("AddressInformation")
					
					Response.Write(sep)
					sep = ","
					Response.Write("{")
					Response.Write("""CompetitorRecID"":""" & EscapeQuotes(CompetitorRecID) & """")
					If rs("PrimaryCompetitor") = vbTrue Then
						Response.Write(",""PrimaryCompetitor"":1")
					Else
						Response.Write(",""PrimaryCompetitor"":0")
					End If
					Response.Write(",""CompetitorName"":""" & EscapeQuotes(rsCompetitorsInner("CompetitorName")) & """")
					Response.Write(",""AddressInformation"":""" & EscapeQuotes(rsCompetitorsInner("AddressInformation")) & """")
					Response.Write(",""Notes"":""" & EscapeQuotes(rs("Notes")) & """")

					If rs("BottledWater") = vbTrue Then
						Response.Write(",""BottledWater"":1")
					Else
						Response.Write(",""BottledWater"":0")
					End If
					If rs("FilteredWater") = vbTrue Then
						Response.Write(",""FilteredWater"":1")
					Else
						Response.Write(",""FilteredWater"":0")
					End If
					If rs("OCS") = vbTrue Then
						Response.Write(",""OCS"":1")
					Else
						Response.Write(",""OCS"":0")
					End If
					If rs("OCS_Supply") = vbTrue Then
						Response.Write(",""OCS_Supply"":1")
					Else
						Response.Write(",""OCS_Supply"":0")
					End If
					If rs("OfficeSupplies") = vbTrue Then
						Response.Write(",""OfficeSupplies"":1")
					Else
						Response.Write(",""OfficeSupplies"":0")
					End If
					If rs("Vending") = vbTrue Then
						Response.Write(",""Vending"":1")
					Else
						Response.Write(",""Vending"":0")
					End If
					If rs("MicroMarket") = vbTrue Then
						Response.Write(",""MicroMarket"":1")
					Else
						Response.Write(",""MicroMarket"":0")
					End If
					If rs("Pantry") = vbTrue Then
						Response.Write(",""Pantry"":1")
					Else
						Response.Write(",""Pantry"":0")
					End If
					
					Response.Write("}")
			
	 				
	 			rsCompetitorsInner.MoveNext
	 			Loop
	 		End If
	 		
			Set rsCompetitorsInner = Nothing
			cnnCompetitorsInner.Close
			Set cnnCompetitorsInner = Nothing
					
		rs.MoveNext						
	Loop
End If
					
Response.Write("]")
Set rs = Nothing
cnn.Close
Set cnn = Nothing

Function EscapeQuotes(val)
	If val <> "" Then
		EscapeQuotes = Replace(val, """", "\""")
	End If
End Function
Function EscapeSingleQuotes(val)
	If val <> "" Then
		EscapeSingleQuotes = Replace(val, "'", "''")
	End If
End Function

%> 
