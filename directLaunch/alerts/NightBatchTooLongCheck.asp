<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/mailDirectlaunch.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<script type="text/javascript">
    function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>


<%
'Alert processing page
'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will check the alerts db and take the appropriate actions
'Usage = "http://{xxx}.{domain}.com/directLaunch/alerts/NightBatchToLongCheck.asp?runlevel=run_now

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then response.end

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles alerts for ALL clients
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1"
If Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 Then ' Good 'ol dev failsafe
	'SQL = "SELECT * FROM tblServerInfo where dbLogin like '%dev'"
End If

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

		
		Response.Write("******** Setting Stopmail Vars for this client: " & ClientKey  & "************<br>")
		
		If Session("ClientCnnString") <> ""Then
			'SEE IF MAIL IS ON OR OFF
			SQLtoggle = "Select STOPALLEMAIL from " & MUV_Read("SQL_Owner") & ".Settings_Global"
			
			Response.write("ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ     " & SQLtoggle & "     ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ")
			Set cnntoggle = Server.CreateObject("ADODB.Connection")
			cnntoggle.open (Session("ClientCnnString"))
			Set rstoggle = Server.CreateObject("ADODB.Recordset")
			rstoggle.CursorLocation = 3 
			Set rstoggle = cnntoggle.Execute(SQLtoggle)
			If rstoggle.Eof Then 
				Session("MAILOFF") = 1 ' If eof then set email to off
				Response.Write("<br><font color='red'>MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-</font>")
			Else
				Session("MAILOFF") = rstoggle("STOPALLEMAIL")
				If Session("MAILOFF") = 1 Then
					Response.Write("<br><font color='red'>MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-MAIL OFF-</font>")				
				Else
				Response.Write("<br><font color='green'>MAIL ON-MAIL ON-MAIL ON-MAIL ON-MAIL ON-MAIL ON-</font>")				
				End IF
			End If
			set rstoggle = Nothing
			cnntoggle.close
			set cnntoggle = Nothing
		Else
			Session("MAILOFF") = 0 ' There was no valid ccn string, so assume it is on
		End If

		
		If MUV_READ("cnnStatus") = "OK" Then ' else it loops
		
			'Each alert is handled individually since we never know what's going to be added
			
			'****************************************
			'Begin Night Batch Running Too Long Alert
			'****************************************
			Response.Write("Begin Night Batch Running Too Long Alert<br>")
			
'**********************************************
			'Delete alerts sent more than 20 hours ago
			'If we don't the system always thinks the alerts
			'have been sent & the max send thresholds have been met
			'ONLY delete night batch, service will do it's own
			Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
			cnnAlertsSent.open (MUV_Read("ClientCnnString"))
			Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
			rsSCAlertsSent.CursorLocation = 3 
			''''''''on error resume next ' In case they dont have this table
			
			rsSCAlertsSent = "SELECT * FROM SC_AlertsSent WHERE DateTimeSent < dateadd(hh,-48,getdate())" ' All cleared after 20hrs
			Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)
			on error goto 0
			
			If Not rsSCAlertsSent.EOF Then
			
				Set rsForDeletes = Server.CreateObject("ADODB.Recordset")
				rsForDeletes.CursorLocation = 3 

				Do While Not rsSCAlertsSent.EOF
					If GetAlertType(rsSCAlertsSent("Alert_InternalAlertRecNumber")) = "NightBatch" Then
						'OK, it is an old night batch alert so we can delete it					
						SQLForDeletes = "DELETE FROM SC_AlertsSent WHERE UniqueRecordIdentifier = " & rsSCAlertsSent("UniqueRecordIdentifier")
						Set rsForDeletes = cnnAlertsSent.Execute(SQLForDeletes)
					End If
					rsSCAlertsSent.MoveNext
				Loop
				
				Set rsForDeletes = Nothing
			End If
			
			Set rsSCAlertsSent = Nothing
			cnnAlertsSent.Close
			Set cnnAlertsSent = Nothing
'**********************************************
			'Also, if nightbatch has been running > 20 hours
			'kill the running file, prob a mistake
			Set cnnAlerts = Server.CreateObject("ADODB.Connection")
			cnnAlerts.open (MUV_Read("ClientCnnString"))
			
			
			'First see if it's even running, if not, nothing to do
			Set rsRunning = Server.CreateObject("ADODB.Recordset")
			rsRunning.CursorLocation = 3 
			SQLRunning = "SELECT * FROM sysobjects WHERE name = 'zNightBatchRunning'"
			Set rsRunning = cnnAlerts.Execute(SQLRunning)
			
			If Not rsRunning.EOF Then ' table exists, so we think it is running
			
				SQLRunning = "SELECT * from zNightBatchRunning"
				Set rsRunning = cnnAlerts.Execute(SQLRunning)

				If Not rsRunning.EOF Then
				
					TimeDateCompare = cDate(Left(rsRunning("StartDate"),10) & " "  & right(rsRunning("StartTime"),11))
					
					If datediff("n",TimeDateCompare,Now()) > 1200 Then 'kill the running file, prob a mistake
					
						Response.Write("<font color='blue'>TimeDateCompare " & TimeDateCompare  & "</font><br>")
						Response.Write("<font color='blue'>Datediff " & datediff("n",TimeDateCompare,Now()) & "</font><br>")
						Response.Write("<font color='blue'>*************Table exists for more than 20 hours, killing table*************</font><br>")
						
						SQLRunning = "DROP TABLE zNightBatchRunning"
						Set rsRunning = cnnAlerts.Execute(SQLRunning)
				
					End If
				
				End If
			
			End If
			
			Set rsRunning = Nothing
			cnnAlerts.Close
			Set cnnAlerts = Nothing
			
			
'**********************************************			
			Set cnnAlerts = Server.CreateObject("ADODB.Connection")
			cnnAlerts.open (MUV_Read("ClientCnnString"))
			
			
			'First see if it's even running, if not, nothing to do
			Set rsRunning = Server.CreateObject("ADODB.Recordset")
			rsRunning.CursorLocation = 3 
					
			SQLRunning = "SELECT * FROM sysobjects WHERE name = 'zNightBatchRunning'"
			Set rsRunning = cnnAlerts.Execute(SQLRunning)
			
			If Not rsRunning.EOF Then ' table exists, so we think it is running
			
				SQLRunning = "SELECT * from zNightBatchRunning"
				Set rsRunning = cnnAlerts.Execute(SQLRunning)
				
				TimeDateCompare = cDate(Left(rsRunning("StartDate"),10) & " "  & right(rsRunning("StartTime"),11))
				Response.Write("TimeDateCompare : " & TimeDateCompare & "<br>")
				
				Set rsRunning = Nothing
				
				SQLAlerts = "SELECT * FROM SC_Alerts Where AlertType='NightBatch' and Condition='Running' And Enabled = 1" 
				
				Set rsAlert = Server.CreateObject("ADODB.Recordset")
				rsAlert.CursorLocation = 3 
				Set rsAlert = cnnAlerts.Execute(SQLAlerts)
			
				If not rsAlert.EOF Then
						
					Do While Not rsAlert.EOF
				
						Response.Write("<b>Found Alert Named: " & rsAlert("AlertName") &"</b><br>") 
						Response.Write("Runnning for minutes is: " & datediff("n",TimeDateCompare,Now()) & "<br>") 
			
						If OKtoSendAlert(rsAlert("InternalAlertRecNumber")) = True Then 
			
							'OK, if more that set amount, send the alert
							If datediff("n",TimeDateCompare,Now()) > rsAlert("NBMinutes") Then
				
									Send_To = ""
									'*************************
									'First do the email alerts
									'*************************
									'Get user based emails
									If Not IsNull(rsAlert("EmailToUserNos")) Then
										If rsAlert("EmailToUserNos") <> "" And rsAlert("EmailToUserNos") <> "0" Then
											UserNoList = Split(rsAlert("EmailToUserNos"),",")
											For x = 0 To UBound(UserNoList)
												Send_To = Send_To & GetUserEmailByUserNo(UserNoList(x)) & ";"
											Next
										End If
									End If
									
																	
									'Get additional emails if there are any
									If rsAlert("AdditionalEmails") <> "" and not IsNull(rsAlert("AdditionalEmails")) Then
										tmpSendAlertToAdditionalEmails = trim(rsAlert("AdditionalEmails"))		
										If Len(tmpSendAlertToAdditionalEmails) > 1 Then
											If Right(tmpSendAlertToAdditionalEmails,1) <> ";" Then tmpSendAlertToAdditionalEmails = tmpSendAlertToAdditionalEmails & ";"
											Send_To = Send_To & tmpSendAlertToAdditionalEmails
										End If
									End If
									
					
									'Only do this if there are actually emails to send
									If Send_To <> "" Then
								
										If rsAlert("EmailVerbiage") <> "" Then
											If Not IsNull(rsAlert("EmailVerbiage")) Then
												AdditionalEmailVerbiage = rsAlert("EmailVerbiage")
											End If
										End If
									
										emailBody = "Night Batch has been running for more than " & rsAlert("NBMinutes") & " minutes. Running minutes so far: " & datediff("n",TimeDateCompare,Now()) & "<br>"
										emailBody = emailBody & AdditionalEmailVerbiage & "<br>"
										emailBody = emailBody & "(Alert name: " & rsAlert("AlertName")& ")"
									
										'Now Send the emails
										'Got all the addresses so now break them up
										Send_To_Array = Split(Send_To,";")
						
										For x = 0 to Ubound(Send_To_Array) -1
											Send_To = Send_To_Array(x)
											Response.Write("<font color='green'><b>Sending alerts to Send_To: " & Send_To & "</b></font><br>")
											'Failsafe for dev

											SendMail "mailsender@" & maildomain,Send_To,"Night Batch Running Too Long " & ClientKey ,emailBody,GetTerm("Alerts"),"Night Batch","MDS Insight"
											
											Description = " A night batch running to long alert was trigerred. (Alert name: " & rsAlert("AlertName")& ") An email was sent to " & Send_To
											CreateAuditLogEntry "Alert Sent","Night batch running to long" ,"Major",0,Description
											
											dummy = MUV_WRITE("ANY_ALERT_SENT","1")
										Next 
									
									End If
								
									Send_To=""
									'**********************
									'Now do the text alerts
									'**********************
									'Get user based Texts
									If Not IsNull(rsAlert("TextToUserNos")) Then
										If rsAlert("TextToUserNos") <> "" And rsAlert("TextToUserNos") <> "0" Then
											UserNoList = Split(rsAlert("TextToUserNos"),",")
											For x = 0 To UBound(UserNoList)
												Send_To = Send_To & getUserCellNumber(UserNoList(x)) & ","
											Next
										End If
									End If
					
									'Get additional texts if there are any
									If rsAlert("AdditionalText") <> "" and not IsNull(rsAlert("AdditionalText")) Then
										tmpSendAlertToAdditionalTexts = trim(rsAlert("AdditionalText"))		
										If Len(tmpSendAlertToAdditionalTexts) > 1 Then
											If Right(tmpSendAlertToAdditionalTexts,1) <> "," Then tmpSendAlertToAdditionalTexts = tmpSendAlertToAdditionalTexts & ","
											 Send_To = Send_To & tmpSendAlertToAdditionalTexts
										End If
									End If
							
									'Only do this if there are actually texts to send
									If Send_To <> "" Then
																
										Send_To = Replace(Send_To,"-","") ' EZ Texting doesn't like dashes
								
										If rsAlert("TextVerbiage") <> "" Then
											If Not IsNull(rsAlert("TextVerbiage")) Then
												AdditionalTextVerbiage = rsAlert("TextVerbiage")
											End If
										End If
									
								
										txtSubject = "NB Too Long"
										txtMessage = ClientKey  & "  Night batch running " & datediff("n",TimeDateCompare,Now()) & " minutes so far. "
										txtMessage = txtMessage & AdditionalTextVerbiage & " "
										txtMessage = txtMessage & "(Alert name: " & rsAlert("AlertName")& ")"
						
										
										'*****Text numbers don't get split into an array, the php takes multiple #'s seprated by commas	
											
										If Right(Send_To,1) = "," Then Send_To = Left(Send_To,Len(Send_To)-1)
									
										TEXT_TO = Send_To
										
										'Split Text_To for recording the alerts sent
										TextNumberArray = Split(TEXT_TO&",",",")
										
										For x = 0 to ubound(TextNumberArray) -1
											Response.Write("TextNumberArray (x):" & TextNumberArray (x) & "<br>")
										Next
		
										Response.Write("<font color='green'><b>Sending text to: "& TEXT_TO & "</b></font><br>")
										
										Description = " A night batch running to long alert was trigerred. (Alert name: " & rsAlert("AlertName")& ") Text messages where sent to " & TEXT_TO
										CreateAuditLogEntry "Alert Sent","Night batch running to long" ,"Major",0,Description

										
										Response.Write("POST TO: " & BaseURL & "inc/sendtext_post.php<br>")
								
										str_data="txtSubject=" & txtSubject  & "&txtMessage=" & txtMessage & "&txtTEXT_TO=" & TEXT_TO
						
										str_data = str_data & "&txtu1=" & EzTextingUserID() & "&txtu2=" & EzTextingPassword()
										
										str_data = str_data & "&txtCountry=" & GetCompanyCountry()
										
										Response.Write("str_data= " & str_data & "<br>")
							
										Set obj_post=Server.CreateObject("Msxml2.SERVERXMLHTTP")
										obj_post.Open "POST", BaseURL & "inc/sendtext_post.php",False
										obj_post.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
										obj_post.Send str_data
										
										dummy = MUV_WRITE("ANY_ALERT_SENT","1")
						
										Response.Write("obj_post.responseText: " & obj_post.responseText & "<br>")
									
									End If
									
									If dummy = MUV_INSPECT("ANY_ALERT_SENT") <> "" Then
										Call WriteAlertSentRecord (rsAlert("InternalAlertRecNumber"))
										MUV_REMOVE("ANY_ALERT_SENT")
									End If
							End If
						Else
							Response.Write("<font color='red'>*************Skipping Alert, NOT OK to send*************</font><br>")
						End IF
						rsAlert.Movenext
					Loop
			
				Else
					Response.Write("Found No Enabled Night Batch Running Too Long Alerts<br>")	
				End If
			
				Set rsAlert = Nothing
				cnnAlerts.Close
				Set cnnAlerts = Nothing
				
			Else
				Response.Write("No table, nightbatch is not running<br>") 
				Set rsRunning = Nothing
				cnnAlerts.Close
				Set cnnAlerts = Nothing
			End If
				
			
			Response.Write("End Night Batch Running Too Long Alert<br>")
			'****************************************
			'End Night Batch Running Too Long Alert
			'****************************************
		End If
		
		Response.Write("******** DONE Processing " & ClientKey  & "************<br>")
				
		TopRecordset.movenext
		
	Loop
	
	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

'Response.write("<script type=""text/javascript"">closeme();</script>")	

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
		Session("SQL_Owner") = Recordset.Fields("dbLogin")
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub


Sub WriteAlertSentRecord (PassedAlertRecIdentifier)

		Set cnnAlertsSent = Server.CreateObject("ADODB.Connection")
		cnnAlertsSent.open (MUV_Read("ClientCnnString"))
		Set rsSCAlertsSent = Server.CreateObject("ADODB.Recordset")
		rsSCAlertsSent.CursorLocation = 3 

		rsSCAlertsSent = "Insert Into SC_AlertsSent (Alert_InternalAlertRecNumber) Values (" &  PassedAlertRecIdentifier & ")"

		Set rsSCAlertsSent = cnnAlertsSent.Execute(rsSCAlertsSent)


		Set rsSCAlertsSent = Nothing
		cnnAlertsSent.Close
		Set cnnAlertsSent = Nothing

End Sub

Function NumberOfNumberOfAlertsSent  (PassedAlertRecIdentifier)

		resultNumberOfNumberOfAlertsSent  = 0 
		
		Set cnnNumberOfAlertsSent  = Server.CreateObject("ADODB.Connection")
		cnnNumberOfAlertsSent.open (MUV_Read("ClientCnnString"))
		Set rsSCNumberOfAlertsSent  = Server.CreateObject("ADODB.Recordset")
		rsSCNumberOfAlertsSent.CursorLocation = 3 

		'Total number sent
		rsSCNumberOfAlertsSent  = "Select Count(*) As TotSent from SC_AlertsSent  Where Alert_InternalAlertRecNumber = " & PassedAlertRecIdentifier 
		Set rsSCNumberOfAlertsSent  = cnnNumberOfAlertsSent.Execute(rsSCNumberOfAlertsSent )
		If not rsSCNumberOfAlertsSent.Eof Then 
			resultNumberOfNumberOfAlertsSent  = rsSCNumberOfAlertsSent ("TotSent")
			Response.Write("NumberOfNumberOfAlertsSent  :" & resultNumberOfNumberOfAlertsSent  & "<br>")
		End If

		Set rsSCNumberOfAlertsSent  = Nothing
		cnnNumberOfAlertsSent.Close
		Set cnnNumberOfAlertsSent  = Nothing
		
		NumberOfNumberOfAlertsSent  = resultNumberOfNumberOfAlertsSent  

End Function 

Function FirstAlertSent  (PassedAlertRecIdentifier)

		resultFirstAlertSent  = ""
		
		Set cnnNumberOfAlertsSent  = Server.CreateObject("ADODB.Connection")
		cnnNumberOfAlertsSent.open (MUV_Read("ClientCnnString"))
		Set rsSCNumberOfAlertsSent  = Server.CreateObject("ADODB.Recordset")
		rsSCNumberOfAlertsSent.CursorLocation = 3 

		'First one sent
		rsSCNumberOfAlertsSent  = "Select Top 1 DateTimeSent from SC_AlertsSent  Where Alert_InternalAlertRecNumber = " & PassedAlertRecIdentifier & " Order by DateTimeSent" 
		Set rsSCNumberOfAlertsSent  = cnnNumberOfAlertsSent.Execute(rsSCNumberOfAlertsSent )
		If not rsSCNumberOfAlertsSent.Eof Then FirstAlertSent = rsSCNumberOfAlertsSent ("DateTimeSent") 

		Set rsSCNumberOfAlertsSent  = Nothing
		cnnNumberOfAlertsSent.Close
		Set cnnNumberOfAlertsSent  = Nothing
		
		FirstAlertSent  = resultFirstAlertSent  

End Function 

Function LastAlertSent  (PassedAlertRecIdentifier)

		resultLastAlertSent  = ""
		
		Set cnnNumberOfAlertsSent  = Server.CreateObject("ADODB.Connection")
		cnnNumberOfAlertsSent.open (MUV_Read("ClientCnnString"))
		Set rsSCNumberOfAlertsSent  = Server.CreateObject("ADODB.Recordset")
		rsSCNumberOfAlertsSent.CursorLocation = 3 

		'First one sent
		rsSCNumberOfAlertsSent  = "Select Top 1 DateTimeSent from SC_AlertsSent  Where Alert_InternalAlertRecNumber = " & PassedAlertRecIdentifier & " Order by DateTimeSent Desc" 
		Set rsSCNumberOfAlertsSent  = cnnNumberOfAlertsSent.Execute(rsSCNumberOfAlertsSent )
		If not rsSCNumberOfAlertsSent.Eof Then
			resultLastAlertSent  = rsSCNumberOfAlertsSent ("DateTimeSent") 
			Response.Write("LastAlertSent :" & resultLastAlertSent  & "<br>")
		End If

		Set rsSCNumberOfAlertsSent  = Nothing
		cnnNumberOfAlertsSent.Close
		Set cnnNumberOfAlertsSent  = Nothing
		
		LastAlertSent  = resultLastAlertSent  

End Function 

Function OKtoSendAlert (PassedAlertRecIdentifier)

	resultOKtoSendAlert = True
	
	'Check Max # Sends
	If NumberOfNumberOfAlertsSent(PassedAlertRecIdentifier) => AlertMaxSends(PassedAlertRecIdentifier) Then resultOKtoSendAlert = False

	'Check velocity & make sure we can send another
	If LastAlertSent(PassedAlertRecIdentifier) <> "" Then
		If cint(datediff("n",LastAlertSent(PassedAlertRecIdentifier),Now())) < cint(AlertLimitMinutes(PassedAlertRecIdentifier)) Then resultOKtoSendAlert = False 'not time yet
		Response.Write("Last alert sent " & datediff("n",LastAlertSent(PassedAlertRecIdentifier),Now()) & " minutes ago<br>")
	End If
	
	OKtoSendAlert = resultOKtoSendAlert 

End Function 

Function AlertMaxSends(PassedAlertRecIdentifier)

		resultAlertMaxSends  = 1
		
		Set cnnAlertMaxSends  = Server.CreateObject("ADODB.Connection")
		cnnAlertMaxSends.open (MUV_Read("ClientCnnString"))
		Set rsSCAlertMaxSends  = Server.CreateObject("ADODB.Recordset")
		rsSCAlertMaxSends.CursorLocation = 3 

		rsSCAlertMaxSends  = "Select NBLimitMaxTimes from SC_Alerts  Where InternalAlertRecNumber = " & PassedAlertRecIdentifier
		Set rsSCAlertMaxSends  = cnnAlertMaxSends.Execute(rsSCAlertMaxSends )
		
		If not rsSCAlertMaxSends.Eof Then 
			resultAlertMaxSends  = rsSCAlertMaxSends ("NBLimitMaxTimes") 
			Response.Write("AlertMaxSends:" & resultAlertMaxSends & "<br>")
		End If

		Set rsSCAlertMaxSends  = Nothing
		cnnAlertMaxSends.Close
		Set cnnAlertMaxSends  = Nothing
		
		AlertMaxSends  = cint(resultAlertMaxSends)

End Function 

Function AlertLimitMinutes(PassedAlertRecIdentifier)

		resultAlertLimitMinutes  = 1
		
		Set cnnAlertLimitMinutes  = Server.CreateObject("ADODB.Connection")
		cnnAlertLimitMinutes.open (MUV_Read("ClientCnnString"))
		Set rsSCAlertLimitMinutes  = Server.CreateObject("ADODB.Recordset")
		rsSCAlertLimitMinutes.CursorLocation = 3 

		rsSCAlertLimitMinutes  = "Select NBLimitMiniutes from SC_Alerts  Where InternalAlertRecNumber = " & PassedAlertRecIdentifier
		Set rsSCAlertLimitMinutes  = cnnAlertLimitMinutes.Execute(rsSCAlertLimitMinutes )
		
		If not rsSCAlertLimitMinutes.Eof Then 
			resultAlertLimitMinutes  = rsSCAlertLimitMinutes ("NBLimitMiniutes") 
			Response.Write("AlertLimitMinutes:" & resultAlertLimitMinutes & "<br>")
		End If

		Set rsSCAlertLimitMinutes  = Nothing
		cnnAlertLimitMinutes.Close
		Set cnnAlertLimitMinutes  = Nothing
		
		AlertLimitMinutes  = cint(resultAlertLimitMinutes)

End Function 
%>