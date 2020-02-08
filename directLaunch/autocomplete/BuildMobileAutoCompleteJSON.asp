<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
<%
Server.ScriptTimeout = 2500
'Delivery Board Alert processing page
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/autocomplete/BuildMobileAutoCompleteJSON.asp?runlevel=run_now

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)
If Request.QueryString("runlevel") <> "run_now" then response.end

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles autocompletes for all databases
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and exit
If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF

		PassPhrase = TopRecordset.Fields("directLaunchPassphrase")
		ClientKey = TopRecordset.Fields("clientkey")

		Response.Write("******** Processing " & ClientKey  & "************<br>")
		Call SetClientCnnString
		
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
		
		If MUV_READ("cnnStatus") = "OK" Then ' else it loops and excludes dev client keys

			serverName = Request.ServerVariables("SERVER_NAME")
			
			If serverName = "www.mdsinsight.com" Then serverName = "mdsinsight.com"
			
			Response.Write("serverName: " & serverName & "<br>")
			
			
			If serverName <> "mdsinsight.com" OR (serverName = "mdsinsight.com" AND UCASE(RIGHT(ClientKey,1)) <> "D") Then 
				
				'Each autocomplete is handled individually as every customer has different accounts
				
				'****************************************
				'Begin Build Autocomplete JSON Files
				'****************************************
				 Response.Write("Begin Build Auto Complete JSON<br>")
				 
				'******************************************


				'*********************************************************
				' Begin Mobile Auto Complete Product List
				'*********************************************************
				 Response.Write("Begin Build Mobile Auto Complete JSON Product List <br>")
				'******************************************

				
				Response.Write("[")
				SQL = "SELECT Distinct prodSKU,prodDescription FROM IC_Product where prodSKU <> '' order by prodSKU"
				Set cnnAutoComplete = Server.CreateObject("ADODB.Connection")
				cnnAutoComplete.open (Session("ClientCnnString"))
				Set rsAutoComplete = Server.CreateObject("ADODB.Recordset")
				rsAutoComplete.CursorLocation = 3 
				Set rsAutoComplete = cnnAutoComplete.Execute(SQL)
				
				If not rsAutoComplete.EOF Then
				strAuto = "["
				Do While Not rsAutoComplete.EOF
				
					 ' {
					  '  "year": "1961",
					  '  "value": "West Side Story",
					   ' "tokens": [
					   '   "West",
					   '   "Side",
					   '   "Story"
					   ' ]
					 ' },
					  
  					tokenList = ""
  					tokensForProduct = ""
  					tokens = ""
  					tokensForProduct = rsAutoComplete("prodSKU") & " " & rsAutoComplete("prodDescription")
  					tokens = split(tokensForProduct, " ")
  					
					for each token in tokens
					    tokenList = tokenList & """" & token & ""","
					next  
					
					If right(tokenList,1)= "," Then tokenList = left(tokenList,len(tokenList)-1)	
									
					strAuto = strAuto & "{" & chr(13) & chr(10)
					strAuto = strAuto & """value"":""" & rsAutoComplete("prodSKU") & """," & chr(13) & chr(10)
				    strAuto = strAuto & """description"":""" & rsAutoComplete("prodDescription") & """," & chr(13) & chr(10)
				    strAuto = strAuto & """display"":""" & rsAutoComplete("prodSKU") & " " & rsAutoComplete("prodDescription") & """," & chr(13) & chr(10)
				    strAuto = strAuto & """tokens"":[" & tokenList & "]" & chr(13) & chr(10)
				    

				    rsAutoComplete.MoveNext
				    
				    If rsAutoComplete.EOF Then
				    	strAuto = strAuto & "}" & chr(13) & chr(10)
				    Else
				    	strAuto = strAuto & "}," & chr(13) & chr(10)
				    End IF
				    
				Loop
				End If
				
				If right(strAuto,1)= "," Then strAuto = left(strAuto,len(strAuto)-1) 
				
				strAuto = trim(strAuto) & "]"
				
				Response.Write("]")
				
				ClientKeyForFileName = ClientKey

			
				set fs=Server.CreateObject("Scripting.FileSystemObject")
				set fs2=Server.CreateObject("Scripting.FileSystemObject")
				
				set tfile=fs.CreateTextFile(Server.MapPath("..\..\") & "\clientfiles\"  & ClientKeyForFileName & "\autocomplete\product_list_mobile_" & ClientKeyForFileName & ".json")
				tfile.WriteLine(strAuto)
				tfile.close
				set tfile=nothing
				set fs=nothing
				
				Set rsAutoComplete = Nothing
				cnnAutoComplete.Close
				Set AutoComplete = nothing

				
				'*********************************************************
				' END Mobile Auto Complete Product List
				'*********************************************************
					
				'******************************************	
				Response.Write("******** DONE Processing " & ClientKey  & "************<br>")
			
		End If
	End If				
	TopRecordset.movenext
	
	Loop

	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")	
'Response.End
'*************************
'*************************
'Subs and funcs begin here


Sub SetClientCnnString

	dummy=MUV_WRITE("cnnStatus","")

	SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString
	
	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	
	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and exit
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
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub



%>