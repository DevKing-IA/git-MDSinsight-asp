<style type="text/css">
	.row-common{
		border: 1px solid #dbdece;
		padding-top: 10px;
		padding-bottom: 10px;
		margin-bottom: 10px;
		font-size: 12px;
	}
</style>

<% ' Lookup the customer record to get the other stuff we need

SQL = "SELECT * FROM " & MUV_Read("SQL_Owner")  & ".FS_ServiceMemos WHERE MemoNumber = '" & SelectedMemoNumber & "' AND CurrentStatus='OPEN'"
						
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.Eof Then
	ProbLocation = rs("ProblemLocation")
	ProbDescription = rs("ProblemDescription")
	SubmittedBy = rs("SubmittedByName")
End IF
Set rs = Nothing
cnn8.close
Set cnn8 = Nothing
%>


<!-- row !-->
 	<div class="col-lg-6 box">
 		<table style="width:100%;">
			<tr><td align="left"><b>Location of problem:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=ProbLocation%>
			<tr><td align="left"><b>Problem:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=ProbDescription%></td></tr>
			<tr><td align="left"><b>Submitted By:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=SubmittedBy%></td></tr>
		</table>
 	</div>
 
