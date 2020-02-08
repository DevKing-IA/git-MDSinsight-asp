<% @ Language = VBScript %>

<!--#include file="inc/SubsAndFuncs.asp"-->
<!--#include file="inc/InsightFuncs.asp"-->

<% MUV_Init() %>

<%
	ClientKey = Request.QueryString("c")
	RandomLoginValue = Request.QueryString("r")

    'Use the ClientKey to lookup SQL credentials
    
	SQLServerInfo = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"
	Set cnnServerInfo = Server.CreateObject("ADODB.Connection")
	Set rsServerInfo  = Server.CreateObject("ADODB.Recordset")
	cnnServerInfo.Open InsightCnnString

	'Open the recordset object executing the SQL statement and return records
	rsServerInfo.Open SQLServerInfo,cnnServerInfo,3,3

	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and Stop
	If rsServerInfo.EOF then
		Response.Write("<br><br><br><br><center>Your client key could not be found. Please contact your administrator.</center><br>")
		cnnServerInfo.close
		Response.End
	End If

	
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & rsServerInfo.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & rsServerInfo.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & rsServerInfo.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & rsServerInfo.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",rsServerInfo.Fields("dbLogin"))
	Session("SQL_Owner") = rsServerInfo.Fields("dbLogin")
	dummy = MUV_Write("ClientID",rsServerInfo.Fields("clientkey"))
	dummy = MUV_Write("BackendSystem",Recordset.Fields("Backend"))

	If rsServerInfo.Fields("advancedDispatch") = 1 Then advancedDispatch = True Else advancedDispatch = False
	'**********************
	' Load Leftnav Options
	'**********************
	dummy = MUV_Write("prospectingModuleOn",rsServerInfo.Fields("prospectingModule"))
	dummy = MUV_Write("routingModuleOn",rsServerInfo.Fields("routingModule"))
	dummy = MUV_Write("biModuleOn",rsServerInfo.Fields("biModule"))
	dummy = MUV_Write("custServiceOn",rsServerInfo.Fields("ShowCustServiceMenu"))
	dummy = MUV_Write("arModuleOn",rsServerInfo.Fields("arModule"))
	dummy = MUV_Write("nightBatchModuleOn",rsServerInfo.Fields("nightBatchModule"))
	dummy = MUV_Write("OrderAPIModuleOn",rsServerInfo.Fields("OrderAPIModule"))
	dummy = MUV_Write("serviceModuleOn",rsServerInfo.Fields("serviceModule"))
	'**********************
			
	rsServerInfo.close	
	cnnServerInfo.close
	
    
	'Now use the Random Login Value to lookup this user in the NagsSent table

	Set cnnNagsSent = Server.CreateObject("ADODB.Connection")
	cnnNagsSent.Open Session("ClientCnnString")
		
	SQLNegsSent = "SELECT * FROM SC_NagsSent WHERE RandomLoginValueIfApplicable= '" & RandomLoginValue & "' AND RecordCreationDateTime < '" & DateAdd("d",7,Now()) & "'"
	
	Set rsNagsSent = Server.CreateObject("ADODB.Recordset")
	rsNagsSent.CursorLocation = 3 
	Set rsNagsSent = cnnNagsSent.Execute(SQLNegsSent)
	
	If rsNagsSent.Eof Then
		CreateAuditLogEntry "Invalid login link","Invalid login link","Minor",0,"An attempt was made to login via ql_text.asp but the link was invalid or expired. The link value used was: " & RandomLoginValue
		Response.Write("<br><br><br><br><br><br><br>")
		Response.Write("<center><font size='18'>The link you are<br>trying to use<br>is not a<br>valid link.</font></center><br>")
		cnnNagsSent.close
		Response.End
	End If
	
	If NOT rsNagsSent.Eof Then
		If Now() > DateAdd("n",30,rsNagsSent("RecordCreationDateTime")) Then
			MessageText = "An attempt was made to login via ql_text.asp by " & GetUserDisplayNameByUserNo(rsNagsSent("UserNoSentToIfApplicable")) & " but the link was expired. "
			MessageText = MessageText  & "The link value used was: " & RandomLoginValue & "which expired " & DateAdd("n",30,rsNagsSent("RecordCreationDateTime")) 
			CreateAuditLogEntry "Expired login link","Expired login link","Minor",0,MessageText
			Response.Write("<br><br><br><br><br><br><br>")
			Response.Write("<center><font size='18'>The link you are<br>trying to use<br>is expired.</font></center><br>")
			cnnNagsSent.close
			Response.End
		End If
	End If
	
	'OK, link is valid & not expired
	'Grab the user & log them in
	
	UserNoToProcess = rsNagsSent("UserNoSentToIfApplicable")
	
	Set rsNagsSent = Nothing
	cnnNagsSent.Close
	Set cnnNagsSent = Nothing
	

	'Gather all the needed fields & submit it to the normal action_login page
	
	
	Set cnnUserLogin = Server.CreateObject("ADODB.Connection")
	cnnUserLogin.Open Session("ClientCnnString")
  	Set rsUserLogin = Server.CreateObject("ADODB.Recordset")
  	
	SQL = "SELECT * FROM tblUsers where UserNo = " & UserNoToProcess 
	'Open the recordset object executing the SQL statement and return records
	rsUserLogin.Open SQL,cnnUserLogin,3,3
   
	If rsUserLogin.EOF Then
		Response.Write("<BR><BR>Unable to login via link<BR><BR>")	
		Response.End
	End If

	'Create the hidden login form
	Response.Write("<form action='" & BaseURL & "action_login.asp' method='POST' name='frmtextLoginForm' id='frmtextLoginForm'>")
		Response.Write("<input type='hidden' name='txtUsername' id='txtUsername' class='form-control' value='" & rsUserLogin("userEmail") & "'>")	
		Response.Write("<input type='hidden' name='txtQuickLogin' id='txtQuickLogin' class='form-control' value='true'>")
		Response.Write("<input type='hidden' name='txtUserNo' id='txtUserNo' class='form-control' value='" & UserNoToProcess & "'>")
		Response.Write("<input type='hidden' name='txtDestinationURL' id='txtDestinationURL' class='form-control' value=''>")	
		Response.Write("<input type='hidden' name='txtClientKey' id='txtClientKey' class='form-control' value='" & ClientKey & "'>")
		Response.Write("<input type='hidden' class='input' name='txtPassword' id='txtPassword' value='" & rsUserLogin("userPassword") & "'>")
	Response.Write("</form>")
	
	'Submit the form %>
	
	<script type="text/javascript">
		document.forms['frmtextLoginForm'].submit();
	</script>
