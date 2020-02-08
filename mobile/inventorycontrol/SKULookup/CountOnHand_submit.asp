<!--#include file="../../../inc/subsandfuncs.asp"-->
<!--#include file="../../../inc/InSightFuncs.asp"-->
<!--#include file="../../../inc/mail.asp"-->
<%

prodSKU = Request.Form("txtprodSKU")
CasesCounted = Request.Form("txtCasesCounted")
UnitsCounted = Request.Form("txtUnitsCounted")
BinLocation = Request.Form("txtBinLocation")
UnitBinLocation = Request.Form("txtUnitBinLocation")
CaseBinLocation = Request.Form("txtCaseBinLocation")

sURL = Request.ServerVariables("SERVER_NAME")
'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)


' Write Audit trail first, then post
	
Set cnnPostOnHandReplacementToBackend = Server.CreateObject("ADODB.Connection")
cnnPostOnHandReplacementToBackend.open (Session("ClientCnnString"))

Set rsRepost = Server.CreateObject("ADODB.Recordset")
rsRepost.CursorLocation = 3 
		
SQLrsRepost = "SELECT * FROM IC_Product WHERE prodSKU = '" & prodSKU & "'"

Response.Write(SQLrsRepost )

Set rsRepost = cnnPostOnHandReplacementToBackend.Execute(SQLrsRepost)

If Not rsRepost.Eof Then
		
	SERNO = GetPOSTParams("SERNO")
	
	Select Case rsRepost("prodCasePricing")
		Case "N"
			UnitBinLocation = BinLocation 		
	End Select
	
	'Construct xml fields based on record
	xmlData = "<DATASTREAM>"
	xmlData = xmlData & "<IDENTITY>Pm8316wyc011</IDENTITY>"
			
	xmlData = xmlData & "<MODE>" & GetPOSTParams("InventoryWebAppPostOnHandMode") & "</MODE>"
			
	xmlData = xmlData & "  <RECORD_TYPE>INVENTORY</RECORD_TYPE>"
	xmlData = xmlData & "  <RECORD_SUBTYPE>REPLACE_ONHAND</RECORD_SUBTYPE>"
			
	xmlData = xmlData & "<SERNO>" & SERNO & "</SERNO>"
			
	xmlData = xmlData & "  <NUM_REPLACE_ONHAND_LINES>1</NUM_REPLACE_ONHAND_LINES>"

	xmlData = xmlData & "  <REPLACE_ONHAND_LINES>"
				
	xmlData = xmlData & " <REPLACE_ONHAND>"
	xmlData = xmlData & "        <DETAIL_NUM>1</DETAIL_NUM>"
	xmlData = xmlData & "        <PROD_ID>" & rsRepost("prodSKU") & "</PROD_ID>"
	xmlData = xmlData & "        <CASES_COUNTED>" & CasesCounted & "</CASES_COUNTED>"
	xmlData = xmlData & "        <UNITS_COUNTED>" & UnitsCounted & "</UNITS_COUNTED>"
	xmlData = xmlData & "        <UNIT_BIN>" & UnitBinLocation & "</UNIT_BIN>"
	xmlData = xmlData & "        <CASE_BIN>" & CaseBinLocation & "</CASE_BIN>"
	xmlData = xmlData & "        <UNIT_UPC></UNIT_UPC>"
	xmlData = xmlData & "        <CASE_UPC></CASE_UPC>"
	xmlData = xmlData & " </REPLACE_ONHAND>"
		 
	xmlData = xmlData & "  </REPLACE_ONHAND_LINES>"
	xmlData = xmlData & "</DATASTREAM>"

	
	Set rsRepost = Nothing
	cnnPostOnHandReplacementToBackend.Close
	Set cnnPostOnHandReplacementToBackend = Nothing
		
			
	xmlDataForDisp = Replace(xmlData,"<","[")
	xmlDataForDisp = Replace(xmlDataForDisp ,">","]")
	xmlDataForDisp = Replace(xmlDataForDisp ,"][","]<br>[")
	xmlDataForDisp = Replace(xmlDataForDisp ,"[","</b>[")
	xmlDataForDisp = Replace(xmlDataForDisp ,"]","]<b>")

	Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP")

	Response.Write(GetPOSTParams("InventoryWebAppPostOnHandURL"))
	
	httpRequest.Open "POST", GetPOSTParams("InventoryWebAppPostOnHandURL"), False
	httpRequest.SetRequestHeader "Content-Type", "text/xml"
	
	xmlData = Replace(xmlData,"&","&amp;")
	xmlData = Replace(xmlData,chr(34),"")			
	httpRequest.Send xmlData

	data = xmlData

	Response.Write("API Response:" & httpRequest.responseText & "<br><br><br>")

	If (Err.Number <> 0 ) Then
		emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>REPLACE ONHAND"& "<br><br>"
		emailBody = "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
		emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
		emailBody = emailBody & "Posted to " & GetPOSTParams("INVENTORYAPIPOSTNHANDURL") & "<br><br>"
		emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
		emailBody = emailBody & "SERNO: " & SERNO & "<br>"
		SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",SERNO & " Post Error Inventory Replace On Hand",emailBody, "Inventory API", "Inventory API"
	
		Description = emailBody 
		Write_API_AuditLog_Entry Request.ServerVariables("SERVER_NAME"),Description,GetPOSTParams("INVENTORYAPIREPOSTMODE"),SERNO,SERNO,"Inventory API"
	End If

	If httpRequest.status = 200 THEN 
	
		If Instr(httpRequest.responseText,"success") <> 0 Then ' Success
	
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>REPLACE ONHAND<"& "<br><br>"
			emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetPOSTParams("INVENTORYAPIPOSTNHANDURL") & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & SERNO & "<br>"
			SendMail "mailsender@" & maildomain ,"insight@ocsaccess.com", SERNO & " Good Post Inventory Replace On Hand",emailBody, "Inventory API", "Inventory API"
			
			Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("INVENTORYAPIREPOSTMODE"),"CountOnHand_submit.asp")
			
		Else
			'FAILURE
			emailbody="httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>REPLACE ONHAND<"& "<br><br>"
			emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetPOSTParams("INVENTORYAPIPOSTNHANDURL") & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & SERNO & "<br>"
			SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",SERNO & " Post Error Inventory Replace On Hand",emailBody, "Inventory API", "Inventory API"
		
			Call Write_API_AuditLog_Entry(Identity ,emailBody ,GetPOSTParams("INVENTORYAPIREPOSTMODE"),"CountOnHand_submit.asp")
			
		End If
		
	Else
	
			'FAILURE
			emailbody="NON 200 Response - httpRequest.status returned " & httpRequest.status & " when posting <RECORD_TYPE>INVENTORY and <RECORD_SUBTYPE>REPLACE ONHAND<"& "<br><br>"
			emailBody = emailBody & "httpRequest.responseText:" & httpRequest.responseText & "<br><br>"
			emailBody = emailBody & "PAGE: " & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "<br>"
			emailBody = emailBody & "Posted to " & GetPOSTParams("INVENTORYAPIPOSTNHANDURL") & "<br><br>"
			emailBody = emailBody & "POSTED DATA:<br>" & xmlDataForDisp & "<br><br>"
			emailBody = emailBody & "SERNO: " & SERNO & "<br>"
			SendMail "mailsender@" & maildomain ,"support@mdsinsight.com",SERNO & " Post Error Inventory Replace On Hand",emailBody, "Inventory API", "Inventory API"
		
			Call Write_API_AuditLog_Entry(Identity ,emailBody,GetPOSTParams("INVENTORYAPIREPOSTMODE"),"CountOnHand_submit.asp")

	End If

End If





Response.Redirect("skulookup.asp")	
	
'	Response.Write("<br>XX" &  data & "XX<br>")
'	Response.Write("<br>XX" &  postResponse & "XX<br>")
'	Response.end
	


%>















