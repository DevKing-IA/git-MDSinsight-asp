<%
'Make entires of invoice #s to skip in zReportConsolidatedInvoiceOmit_

ivshistsequence = Request.Form("ivshistsequence")

If ivshistsequence <> "" Then 

	Set cnnTmpTable = Server.CreateObject("ADODB.Connection")
	cnnTmpTable.open (Session("ClientCnnString"))
	Set rsTmpTable = Server.CreateObject("ADODB.Recordset")
	rsTmpTable.CursorLocation = 3 
	
	SQLTmpTable = "INSERT INTO zReportConsolidatedInvoiceOmit_" & Trim(Session("UserNo")) & " (IvsHistSequence) VALUES ('" & IvsHistSequence & "')"	
	
	Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)
	
	set rsTmpTable = Nothing
	cnnTmpTable.close
	set cnnTmpTable = Nothing

End If	
%>
