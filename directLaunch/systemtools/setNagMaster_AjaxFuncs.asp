<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
 
'Sub TurnOnMasterNagAlertsForClientID()
'Sub TurnOffMasterNagAlertsForClientID()

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
	Case "TurnOnMasterNagAlertsForClientID"
		TurnOnMasterNagAlertsForClientID()
	Case "TurnOffMasterNagAlertsForClientID"
		TurnOffMasterNagAlertsForClientID()
End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub TurnOnMasterNagAlertsForClientID() 

	clientKey = Request.Form("clientKey")
	
	'*****************************************************************************************************************
	'Get the database login information from tblServerInfo for the passed Client Key
	'*****************************************************************************************************************

	Call SetClientCnnString(clientKey)
	
	Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this

	'*****************************************************************************************************************
	'Get the database login information from tblServerInfo for the passed Client Key
	'*****************************************************************************************************************


	Set cnnGlobalMasterNagAlert = Server.CreateObject("ADODB.Connection")
	cnnGlobalMasterNagAlert.open (Session("ClientCnnString"))
	Set rsGlobalMasterNagAlert = Server.CreateObject("ADODB.Recordset")
	rsGlobalMasterNagAlert.CursorLocation = 3 
	
	SQLGlobalMasterNagAlert = "UPDATE Settings_Global SET MasterNagMessageONOFF = 1"
	
	Set rsGlobalMasterNagAlert = cnnGlobalMasterNagAlert.Execute(SQLGlobalMasterNagAlert)
	
	set rsGlobalMasterNagAlert = Nothing
	cnnGlobalMasterNagAlert.close
	set cnnGlobalMasterNagAlert = Nothing
	
	
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub TurnOffMasterNagAlertsForClientID() 

	clientKey = Request.Form("clientKey")
	
	'*****************************************************************************************************************
	'Get the database login information from tblServerInfo for the passed Client Key
	'*****************************************************************************************************************

	Call SetClientCnnString(clientKey)
	
	Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this

	'*****************************************************************************************************************
	'Get the database login information from tblServerInfo for the passed Client Key
	'*****************************************************************************************************************


	Set cnnGlobalMasterNagAlert = Server.CreateObject("ADODB.Connection")
	cnnGlobalMasterNagAlert.open (Session("ClientCnnString"))
	Set rsGlobalMasterNagAlert = Server.CreateObject("ADODB.Recordset")
	rsGlobalMasterNagAlert.CursorLocation = 3 
	
	SQLGlobalMasterNagAlert = "UPDATE Settings_Global SET MasterNagMessageONOFF = 0"
	
	Set rsGlobalMasterNagAlert = cnnGlobalMasterNagAlert.Execute(SQLGlobalMasterNagAlert)
	
	set rsGlobalMasterNagAlert = Nothing
	cnnGlobalMasterNagAlert.close
	set cnnGlobalMasterNagAlert = Nothing

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'END ALL AJAX MODAL SUBROUTINES AND FUNCTIONS

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'************************************************************************************
'************************************************************************************
'Subs and funcs begin here
'************************************************************************************

Sub SetClientCnnString(clientKey)

	dummy=MUV_WRITE("cnnStatus","")

	SQL = "SELECT * FROM tblServerInfo WHERE clientKey= '" & ClientKey & "'"

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString
	
	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection and exit
	If Recordset.recordcount <= 0 then
		Recordset.close
		Connection.close
		set Recordset=nothing
		set Connection=nothing
	Else
		ClientCnnString = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
		ClientCnnString = ClientCnnString & ";Database=" & Recordset.Fields("dbCatalog")
		ClientCnnString = ClientCnnString & ";Uid=" & Recordset.Fields("dbLogin")
		ClientCnnString = ClientCnnString & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
		dummy = MUV_Write("ClientCnnString",ClientCnnString)
		dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
		Session("SQL_Owner") = Recordset.Fields("dbLogin")
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub



'************************************************************************************
'************************************************************************************
'Subs and funcs end here
'************************************************************************************



%>