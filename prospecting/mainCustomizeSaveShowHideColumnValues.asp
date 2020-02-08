<%

showhideCol_address		= Request.Form("chkCol_address")
showhideCol_city		= Request.Form("chkCol_city")
showhideCol_state		= Request.Form("chkCol_state")
showhideCol_zip			= Request.Form("chkCol_zip")
showhideCol_leadsource	= Request.Form("chkCol_leadsource")
showhideCol_stage		= Request.Form("chkCol_stage")
showhideCol_industry	= Request.Form("chkCol_industry")
showhideCol_numemployees= Request.Form("chkCol_numemployees")
showhideCol_owner		= Request.Form("chkCol_owner")
showhideCol_createddate	= Request.Form("chkCol_createddate")
showhideCol_createdby	= Request.Form("chkCol_createdby")
showhideCol_telemarketer= Request.Form("chkCol_telemarketer")
showhideCol_numpantries	= Request.Form("chkCol_numpantries")
showhideCol_prospectid	= Request.Form("chkCol_prospectid")



columnsData = showhideCol_address & "," & showhideCol_city & "," & showhideCol_state & "," & showhideCol_zip & "," & showhideCol_leadsource
columnsData = columnsData & showhideCol_stage & "," & showhideCol_industry & "," & showhideCol_numemployees & "," & showhideCol_owner & "," & showhideCol_createddate
columnsData = columnsData & showhideCol_createdby & "," & showhideCol_telemarketer & "," & showhideCol_numpantries & "," & showhideCol_prospectid

'Response.write("columnsData : " & columnsData & "<br>")

' Splitting based on delimiter comma ','

columnsDataArray = Split(columnsData,",")
upperBound = ubound(columnsDataArray)
columnsDataSQL = "" 

For i=0 to upperBound
   'Response.write("The value of array in " & i & " is :"  & columnsDataArray(i) & "<br />")
   If columnsDataArray(i) <> "" Then
   		columnsDataSQL = columnsDataSQL & columnsDataArray(i) & ","
   End If
Next

If columnsDataSQL <> "" Then
	columnsDataSQL = Left(columnsDataSQL, Len(columnsDataSQL)-1)
End If


SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = 'Current'"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)

'Rec does not exist yet, make it quick but empty, update it later
If rs.EOF Then
	SQL = "Insert into Settings_Reports (ReportNumber, UserNo, PoolForProspecting, UserReportName) Values (1400, " & Session("userNo") & ",'Live','Current')"
	rs.Close
	Set rs= cnn8.Execute(SQL)
End If

'Now update the table with the values
SQL = "Update Settings_Reports Set ReportSpecificData1 = '" & columnsDataSQL & "' WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = 'Current'"
Set rs= cnn8.Execute(SQL)

cnn8.Close

Set rs = Nothing
Set cnn8 = Nothing

	%>
<form id="frmSetColumnView" name="frmSetColumnView" method="POST" action="main.asp">
	<input type="hidden" name="selectFilteredView" id="selectFilteredView" value="Current">
</form>

<script type="text/javascript">
  document.forms['frmSetColumnView'].submit();
</script>	

 
