<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->

<%
InternalRecordIdentifier = Request.Form("txtInternalRecordIdentifier")

'*******************************************************************
'Lookup the record as it exists now so we can fillin the audit trail

SQL = "SELECT * FROM PR_Competitors where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	Orig_Competitor = rs("CompetitorName")
	Orig_CompetitorAddressInfo = rs("AddressInformation")
	Orig_CompetitorWebsite = rs("CompetitorWebsite")
	Orig_CompetitorAdditionalNotes = rs("AdditionalNotes")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now so we can fillin the audit trail
'***********************************************************************

competitorName 			= Request.Form("txtCompetitorName")
competitorAddressInfo	= Request.Form("txtCompetitorAddressInfo")
txtCompetitorWebsite = Request.Form("txtCompetitorWebsite")
txtCompetitorAdditionalNotes = Request.Form("txtCompetitorAdditionalNotes")

'check if fields are not empty
If competitorName<>"" Then
	competitorName = Hacker_Filter2(competitorName)
End If
If competitorAddressInfo<>"" Then
	competitorAddressInfo = Hacker_Filter2(competitorAddressInfo)
End If
If txtCompetitorWebsite<>"" Then
	txtCompetitorWebsite = Hacker_Filter2(txtCompetitorWebsite)
End If
If txtCompetitorAdditionalNotes<>"" Then
	txtCompetitorAdditionalNotes = Hacker_Filter2(txtCompetitorAdditionalNotes)
End If


SQL = "UPDATE PR_Competitors SET "
SQL = SQL &  " CompetitorName = '" & CompetitorName & "', AddressInformation= '" & competitorAddressInfo & "'"
SQL = SQL &  ",CompetitorWebsite = '" & txtCompetitorWebsite & "', AdditionalNotes= '" & txtCompetitorAdditionalNotes & "'"
SQL = SQL & " WHERE InternalRecordIdentifier = " & InternalRecordIdentifier


Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing


Description = ""
If Orig_Competitor  <> CompetitorName  Then
	Description = Description & GetTerm("Prospecting") & " Competitor changed from " & Orig_Competitor & " to " & CompetitorName
End If
If Orig_CompetitorAddressInfo  <> competitorAddressInfo Then
	Description = Description & GetTerm("Prospecting") & " Competitor address changed from " & Orig_CompetitorAddressInfo  & " to " & competitorAddressInfo & " for competitor " & CompetitorName
End If
If Orig_CompetitorWebsite  <> txtCompetitorWebsite Then
	Description = Description & GetTerm("Prospecting") & " Competitor web site changed from " & Orig_CompetitorWebsite  & " to " & txtCompetitorWebsite & " for competitor " & CompetitorName
End If
If Orig_CompetitorAdditionalNotes  <> txtCompetitorAdditionalNotes Then
	Description = Description & GetTerm("Prospecting") & " Competitor additional notes changed from " & Orig_CompetitorAdditionalNotes  & " to " & txtCompetitorAdditionalNotes & " for competitor " & CompetitorName
End If


CreateAuditLogEntry GetTerm("Prospecting") & " Competitor edited",GetTerm("Prospecting") & " Competitor edited","Minor",0,Description

Response.Redirect("main.asp")

%>















