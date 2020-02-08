<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_Routing.asp"-->

<%
	clientKey = Request.Form("clientKey")
	
	'*****************************************************************************************************************
	'Get the database login information from tblServerInfo for the passed Client Key
	'*****************************************************************************************************************

	SQL = "SELECT * FROM tblServerInfo WHERE ClientKey = '" & clientKey & "'"
	
	'Response.write(SQL & "<br><br>")
	
	Set TopConnection = Server.CreateObject("ADODB.Connection")
	Set TopRecordset = Server.CreateObject("ADODB.Recordset")
	TopConnection.Open InsightCnnString
		
	'Open the recordset object executing the SQL statement and return records
	TopRecordset.Open SQL,TopConnection,3,3
	
	dbToUpdate = TopRecordset.Fields("dbCatalog")
	dbOwnerToUpdate = TopRecordset.Fields("dbLogin")

	'*****************************************************************************************************************
	'Called the stored procedure in _MDSInsight to run the Delivery Board Nightly Process
	'*****************************************************************************************************************
		              	
	Set cnnUpdateDelBoard = Server.CreateObject("ADODB.Connection")
	cnnUpdateDelBoard.open (InsightCnnString)
	Set rsUpdateDelBoard = Server.CreateObject("ADODB.Recordset")
	rsUpdateDelBoard.CursorLocation = 3 
	
	SQLUpdateDelBoard= "EXEC dbo.RT_DeliveryBoardNightlyProcess @dbown = " & dbOwnerToUpdate & ", @db = " & dbToUpdate
	
	'Response.write(SQLUpdateDelBoard & "<br><br>")
	
	Set rsUpdateDelBoard = cnnUpdateDelBoard.Execute(SQLUpdateDelBoard)

	'*****************************************************************************************************************
	'Close the recordsets and connections
	'*****************************************************************************************************************

	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing

	Set rsUpdateDelBoard = Nothing
	cnnUpdateDelBoard.Close
	Set cnnUpdateDelBoard = Nothing

	
	%>
