<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%
competitorName = Request.Form("txtCompetitorName")
competitorAddressInfo = Request.Form("txtCompetitorAddressInfo")
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


SQL = "INSERT INTO PR_Competitors (CompetitorName, AddressInformation,CompetitorWebsite,AdditionalNotes) VALUES ('"  & competitorName & "','"  & competitorAddressInfo & "','"&txtCompetitorWebsite&"','"&txtCompetitorAdditionalNotes&"')"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))

Set rs8 = Server.CreateObject("ADODB.Recordset")
rs8.CursorLocation = 3 
Response.Write(SQL)
Set rs8 = cnn8.Execute(SQL)
set rs8 = Nothing

Description = GetUserDisplayNameByUserNo(Session("Userno")) & " added the " & GetTerm("Prospecting") & " Competitor: " & CompetitorName 
CreateAuditLogEntry GetTerm("Prospecting") & " Competitor added",GetTerm("Prospecting") & " Competitor added","Minor",0,Description

Response.Redirect("main.asp")

%>















