<%
'Remove entires from zReportConsolidatedInvoiceOmit_

	Set cnnTmpTable = Server.CreateObject("ADODB.Connection")
	cnnTmpTable.open (Session("ClientCnnString"))
	Set rsTmpTable = Server.CreateObject("ADODB.Recordset")
	rsTmpTable.CursorLocation = 3 
	
	SQLTmpTable = "DELETE FROM zReportConsolidatedInvoiceOmit_" & Trim(Session("UserNo"))
	
	Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)
	
	set rsTmpTable = Nothing
	cnnTmpTable.close
	set cnnTmpTable = Nothing


%>
