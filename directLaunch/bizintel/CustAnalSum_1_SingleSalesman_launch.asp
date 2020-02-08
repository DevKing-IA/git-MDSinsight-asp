<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/mailDirectLaunch.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_BizIntel.asp"--> 
<!--#include file="../../inc/InSightFuncs_Equipment.asp"--> 

<script type="text/javascript">
    function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>

<%
'Response.End
' This page processes the customer analysis single salesman emails
'Designed to be launched via a scheduled process
'Self contained page will check both global settings to see if reports need to be sent out
'Usage = "http://{xxx}.{domain}.com/directLaunch/bizintel/CatAnalSum1_SingleSalesman.asp?runlevel=run_now
Server.ScriptTimeout = 2500

TestMode = False
If Request.QueryString("m") = "test" then
	TestMode = True ' Sends all emails to rich
End IF 



Dim EntryThread

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then
	Response.Write("Improper usage, no run level was specified in the query string")	
	response.end
End IF 


'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles alerts for ALL clients
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1 "

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
	
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3


If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF
	
		ClientKey = TopRecordset.Fields("clientkey")
		PROCESS_REPORT = True
	
		'To begin with, see if this client uses the bizintel module 
		'If they don't then don't bother checking for Nags
		If TopRecordset.Fields("biModule") = "Enabled" Then
	
			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then 
	
				Call SetClientCnnString
				
				Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
				
				If MUV_READ("cnnStatus") = "OK" Then ' else it loops

					'**************************************************************
					'Get next Entry Thread for use in the SC_AuditLogDLaunch table
					On Error Goto 0
					Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
					cnnAuditLog.open MUV_READ("ClientCnnString") 
					Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
					rsAuditLog.CursorLocation = 3 
					Set rsAuditLog = cnnAuditLog.Execute("Select TOP 1 * from SC_AuditLogDLaunch order by EntryThread desc")
					If Not rsAuditLog.EOF Then 
						If IsNull(rsAuditLog("EntryThread")) Then EntryThread =1 Else EntryThread = rsAuditLog("EntryThread") + 1
					Else
						EntryThread = 1
					End If
					set rsAuditLog = nothing
					cnnAuditLog.close
					set cnnAuditLog = nothing

				
					' Before we do anything, see if this report is turned on for this client
					Set cnnBizIntel = Server.CreateObject("ADODB.Connection")
					cnnBizIntel.open MUV_READ("ClientCnnString") 
					Set rsBizIntel = Server.CreateObject("ADODB.Recordset")
					rsBizIntel.CursorLocation = 3 
					Set rsBizIntel = cnnBizIntel.Execute("Select * from Settings_BizIntel")
					If Not rsBizIntel.EOF Then 
						If rsBizIntel("CustAnalSum1OnOff") <> 1 or IsNUll(rsBizIntel("CustAnalSum1OnOff")) Then
							PROCESS_REPORT = False
							WriteResponse ("Report turned off for client " & ClientKey & " in Settings_BizIntel<BR>")
						End If
					Else
						PROCESS_REPORT = False ' If eof, they dont even have the record, so dont run the report
						WriteResponse ("No Record Found for client " & ClientKey & " in Settings_BizIntel<BR>")
					End If
					set rsBizIntel = nothing
					cnnBizIntel.close
					set cnnBizIntel = nothing
				
					If PROCESS_REPORT = True Then
		
						CreateAuditLogEntry "BizIntel Customer Analysis Single Salesman","BizIntel Customer Analysis Single Salesman","Minor",0,"BizIntel Customer Analysis Single Salesman ran."					
		
						WriteResponse "<font color='purple' size='24'>Start processing " & ClientKey  & "</font><br>"
		
						WriteResponse ("Setting Stopmail vars for " & ClientKey  & "<br>")
						
						If Session("ClientCnnString") <> ""Then
							'SEE IF MAIL IS ON OR OFF
							SQLtoggle = "Select STOPALLEMAIL from " & MUV_Read("SQL_Owner") & ".Settings_Global"
							
							WriteResponse (SQLtoggle & "<br>")
							Set cnntoggle = Server.CreateObject("ADODB.Connection")
							cnntoggle.open (Session("ClientCnnString"))
							Set rstoggle = Server.CreateObject("ADODB.Recordset")
							rstoggle.CursorLocation = 3 
							Set rstoggle = cnntoggle.Execute(SQLtoggle)
							If rstoggle.Eof Then 
								Session("MAILOFF") = 1 ' If eof then set email to off
								WriteResponse ("<font color='red'>MAIL OFF</font><br>")
							Else
								Session("MAILOFF") = rstoggle("STOPALLEMAIL")
								If Session("MAILOFF") = 1 Then
									WriteResponse ("<font color='red'>MAIL OFF<br>-</font>")				
								Else
								WriteResponse ("<font color='green'>MAIL ON<br></font>")				
								End IF
							End If
							set rstoggle = Nothing
							cnntoggle.close
							set cnntoggle = Nothing
						Else
							Session("MAILOFF") = 0 ' There was no valid ccn string, so assume it is on
						End If
					
					End If
				
				
					If PROCESS_REPORT = True Then
					
						WriteResponse ("Begin processing report<br>")
						
						'***********************************
						' Here is where the real work begins
						'***********************************
						
						SQLUserList = "SELECT * FROM Settings_BizIntel"
						Set cnnUserList = Server.CreateObject("ADODB.Connection")
						cnnUserList.open (Session("ClientCnnString"))
						Set rsUserList = Server.CreateObject("ADODB.Recordset")
						rsUserList.CursorLocation = 3 
						Set rsUserList = cnnUserList.Execute(SQLUserList)

						UserList = rsUserList("CustAnalSum1EmailToUserNos")
						CustAnalSum1UserNosToCC = rsUserList("CustAnalSum1UserNosToCC")
						CustAnalSum1EmailAddressesToCC = rsUserList("CustAnalSum1EmailAddressesToCC")
						
						Set rsUserList  = Nothing
						cnnUserList.Close
						Set cnnUserList  = Nothing
						
						UserListArray = Split(UserList,",")
						
						Response.Write("UserListArray: " & UserList & "<br>")

						'***********************************
						' PRIMARY salesman logic starts here
						'***********************************	
						SecondarySalesMan = "" 'To blank out the seoondaries
											
						SQLsalesman = "SELECT DISTINCT SalesMan FROM AR_CUSTOMER WHERE Salesman <> 0 ORDER BY Salesman"
						
						
						Set cnnsalesman = Server.CreateObject("ADODB.Connection")
						cnnsalesman.open (Session("ClientCnnString"))
						Set rssalesman = Server.CreateObject("ADODB.Recordset")
						rssalesman.CursorLocation = 3 
						Set rssalesman = cnnsalesman.Execute(SQLsalesman)

						If Not rssalesman.EOF Then
						
							Do While NOT rssalesman.Eof
							
								Salesman = rssalesman("SalesMan")
								
								ProcessThisSalesPerson = False
								UserNoToProcess = ""
								
								For x = 0 To Ubound(UserListArray)
								
									If GetSalesPersonNoByUserNo(UserListArray(x)) = Salesman Then ' ok, they are a salesman
										ProcessThisSalesPerson = True
										UserNoToProcess = UserListArray(x)
										Exit For
									End If
								Next
								
								If ProcessThisSalesPerson = True Then %>
 
									<!--#include file="CustAnalSum_1_SingleSalesman_EmailWithTextAndLink.asp"-->
									<%
									
									If TotalCustsReported > 0 Then ' This var gets set in the include file
			
										WriteResponse("Sending email to salesperson: " & GetUserDisplayNameByUserNo(UserNoToProcess) & "<br>")
										WriteResponse("Email for this salesperson is : " &  GetUserEmailByUserNo(UserNoToProcess) & "<br>")
										
										Send_To = GetUserEmailByUserNo(UserNoToProcess)
										
	
										'***************
										'Send the email
										'**************
										emailSubject = GetTerm("Customer") & " analysis for " & GetUserDisplayNameByUserNo(UserNoToProcess) 
										
										AllCCEmailAddresses = ""
										If CustAnalSum1EmailAddressesToCC <> "" Then AllCCEmailAddresses = AllCCEmailAddresses & CustAnalSum1EmailAddressesToCC 
	
										If CustAnalSum1UserNosToCC <> "" Then
			
											CustAnalSum1UserNosToCCArray = Split(CustAnalSum1UserNosToCC,",")
											
											For i = 0 to Ubound(CustAnalSum1UserNosToCCArray )
	
												If GetUserEmailByUserNo(CustAnalSum1UserNosToCCArray(i)) <> "" Then
												
													AllCCEmailAddresses = AllCCEmailAddresses & ";" & GetUserEmailByUserNo(CustAnalSum1UserNosToCCArray(i))
												
												End If
											Next
											
	
										End If
										
										If TestMode = True Then AllCCEmailAddresses = "rsmith@ocsaccess.com;rich@ocsaccess.com"
										If AllCCEmailAddresses <> "" Then 
											WriteResponse("AllCCEmailAddresses: " &  AllCCEmailAddresses  & "<br>")
											If TestMode = True Then Send_To = "rsmith@ocsaccess.com"
											SendMailWithCCs "mailsender@" & maildomain,Send_To,emailSubject,emailBody,AllCCEmailAddresses,"","Biz Intel","Biz Intel","MDS Insight"
											Response.Write("Email sent<br>")
										Else
											If TestMode = True Then Send_To = "rsmith@ocsaccess.com"
											SendMail "mailsender@" & maildomain,Send_To,emailSubject,emailBody,"Biz Intel","Biz Intel","MDS Insight"
											Response.Write("Email sent<br>")
										End If
									
									Else
										
										WriteResponse("NOT NOT NOT Sending email to salesperson: " & GetUserDisplayNameByUserNo(UserNoToProcess) & "<br>")
										WriteResponse("TotalCustsReported < 1 , NO customers to report<br>")
								
									End If ' for TotalCustsReported 
								
								End If ' for ProcessThisSalesPerson 
								
								rssalesman.movenext
							
							Loop
						
						End If

						rssalesman.Close
							
						'*************************************
						' SECONDARY salesman logic starts here
						'*************************************			
						SalesMan = "" ' To blank out the primaries
									
						SQLsalesman = "SELECT DISTINCT SecondarySalesMan FROM AR_CUSTOMER WHERE SecondarySalesMan <> 0 ORDER BY SecondarySalesMan"
						
						rssalesman.CursorLocation = 3 
						Set rssalesman = cnnsalesman.Execute(SQLsalesman)

						If Not rssalesman.EOF Then
						
							Do While NOT rssalesman.Eof
							
								SecondarySalesMan = rssalesman("SecondarySalesMan")
								
								ProcessThisSalesPerson = False
								UserNoToProcess = ""
								
								For x = 0 To Ubound(UserListArray)
								
									If GetSalesPersonNoByUserNo(UserListArray(x)) = SecondarySalesMan Then ' ok, they are a salesman
										ProcessThisSalesPerson = True
										UserNoToProcess = UserListArray(x)
										Exit For
									End If
								Next
								
								If ProcessThisSalesPerson = True Then %>
 
									<!--#include file="CustAnalSum_1_SingleSalesman_EmailWithTextAndLink.asp"-->
									<%
												
									WriteResponse("Sending email to secondary salesperson: " & GetUserDisplayNameByUserNo(UserNoToProcess) & "<br>")
									WriteResponse("Email for this secondary salesperson is : " &  GetUserEmailByUserNo(UserNoToProcess) & "<br>")
									
									Send_To = GetUserEmailByUserNo(UserNoToProcess)
									

									'***************
									'Send the email
									'**************
									emailSubject = GetTerm("Customer") & " analysis for " & GetUserDisplayNameByUserNo(UserNoToProcess) 
									
									AllCCEmailAddresses = ""
									If CustAnalSum1EmailAddressesToCC <> "" Then AllCCEmailAddresses = AllCCEmailAddresses & CustAnalSum1EmailAddressesToCC 

									If CustAnalSum1UserNosToCC <> "" Then
		
										CustAnalSum1UserNosToCCArray = Split(CustAnalSum1UserNosToCC,",")
										
										For i = 0 to Ubound(CustAnalSum1UserNosToCCArray )

											If GetUserEmailByUserNo(CustAnalSum1UserNosToCCArray(i)) <> "" Then
											
												AllCCEmailAddresses = AllCCEmailAddresses & ";" & GetUserEmailByUserNo(CustAnalSum1UserNosToCCArray(i))
											
											End If
										Next
										

									End If
									
									If TestMode = True Then AllCCEmailAddresses = "rsmith@ocsaccess.com;rich@ocsaccess.com"
									If AllCCEmailAddresses <> "" Then 
										WriteResponse("AllCCEmailAddresses: " &  AllCCEmailAddresses  & "<br>")
										If TestMode = True Then Send_To = "rsmith@ocsaccess.com"
										SendMailWithCCs "mailsender@" & maildomain,Send_To,emailSubject,emailBody,AllCCEmailAddresses,"","Biz Intel","Biz Intel","MDS Insight"
										Response.Write("Email sent<br>")
									Else
										If TestMode = True Then Send_To = "rsmith@ocsaccess.com"
										SendMail "mailsender@" & maildomain,Send_To,emailSubject,emailBody,"Biz Intel","Biz Intel","MDS Insight"
										Response.Write("Email sent<br>")
									End If
									
									
								
								End If
								
								rssalesman.movenext
							
							Loop
						
						End If
						
					WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				End If ' From PROCESS_Report
			
			End If
			
		End If	
		
	Else ' is the routing module enabled
		WriteResponse ("Skipping the client " & ClientKey & " because the bi module is not enabled.<BR>")
	End If ' is the routing module enabled
	
	TopRecordset.movenext
	
	Loop
	
	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")	

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


Sub WriteResponse(passedLogEntry)

	response.write(Now() & "&nbsp;&nbsp;&nbsp;" & passedLogEntry)
	
	passedLogEntry = Replace(passedLogEntry,"'","''")
	
	SQL = "INSERT INTO SC_AuditLogDLaunch (EntryThread, DirectLaunchName, DirectLaunchFile, LogEntry)"
	SQL = SQL &  " VALUES (" & EntryThread & ""
	SQL = SQL & ",'Routing Module Nag Check'"
	SQL = SQL & ",'/directlaunch/nags/RoutingModuleNagCheck.asp'"
	SQL = SQL & ",'"  & passedLogEntry & "'"
	SQL = SQL & ")"
	
	Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
	cnnAuditLog.open MUV_READ("ClientCnnString") 
	Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
	rsAuditLog.CursorLocation = 3 

	Set rsAuditLog = cnnAuditLog.Execute(SQL)

	set rsAuditLog = nothing
	cnnAuditLog.close
	set cnnAuditLog = nothing

End Sub


Function GetCurrent_UnpostedTotal_ByCust(passedCustID,passedPeriodBeingEvaluated)

	StartDateToFind = GetPeriodBeginDateBySeq(passedPeriodBeingEvaluated+1)
	EndDateToFind = GetPeriodEndDateBySeq(passedPeriodBeingEvaluated+1)
	
	resultGetCurrent_UnpostedTotal_ByCust = 0

	Set cnnGetCurrent_UnpostedTotal_ByCust = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_UnpostedTotal_ByCust.open Session("ClientCnnString")
		
	'SQLGetCurrent_UnpostedTotal_ByCust = "SELECT SUM(InvoiceTotal-SalesTaxCharge-Deposit) AS TotalForCurrent FROM Telsel WHERE (InvoiceTFlag = 'O' OR InvoiceTFlag = 'T') AND CustNum = " & passedCustID & " AND ("
	'SQLGetCurrent_UnpostedTotal_ByCust = SQLGetCurrent_UnpostedTotal_ByCust & "InvoiceDate >= '" & StartDateToFind & "' AND "
	'SQLGetCurrent_UnpostedTotal_ByCust = SQLGetCurrent_UnpostedTotal_ByCust & "InvoiceDate <= '" & EndDateToFind & "')"

	SQLGetCurrent_UnpostedTotal_ByCust = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE CustID='" & passedCustID & "' AND "
	SQLGetCurrent_UnpostedTotal_ByCust = SQLGetCurrent_UnpostedTotal_ByCust & "PostedOrUnposted = 'U' AND ThisPeriodSeqNumber = " & passedPeriodBeingEvaluated+1
'response.write(SQLGetCurrent_UnpostedTotal_ByCust & "<br>")
	Set rsGetCurrent_UnpostedTotal_ByCust = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_UnpostedTotal_ByCust.CursorLocation = 3 
	Set rsGetCurrent_UnpostedTotal_ByCust = cnnGetCurrent_UnpostedTotal_ByCust.Execute(SQLGetCurrent_UnpostedTotal_ByCust)

	If not rsGetCurrent_UnpostedTotal_ByCust.EOF Then resultGetCurrent_UnpostedTotal_ByCust = rsGetCurrent_UnpostedTotal_ByCust("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_UnpostedTotal_ByCust) Then resultGetCurrent_UnpostedTotal_ByCust = 0 ' In case there are no results
	
	rsGetCurrent_UnpostedTotal_ByCust.Close
	set rsGetCurrent_UnpostedTotal_ByCust= Nothing
	cnnGetCurrent_UnpostedTotal_ByCust.Close	
	set cnnGetCurrent_UnpostedTotal_ByCust= Nothing
	
	GetCurrent_UnpostedTotal_ByCust = resultGetCurrent_UnpostedTotal_ByCust 

End Function



Function GetCurrent_PostedTotal_ByCust(passedCustID,passedPeriodBeingEvaluated)


	StartDateToFind = GetPeriodBeginDateBySeq(passedPeriodBeingEvaluated+1)
	EndDateToFind = GetPeriodEndDateBySeq(passedPeriodBeingEvaluated+1)
	
	resultGetCurrent_PostedTotal_ByCust = 0

	Set cnnGetCurrent_PostedTotal_ByCust = Server.CreateObject("ADODB.Connection")
	cnnGetCurrent_PostedTotal_ByCust.open Session("ClientCnnString")
		
	'SQLGetCurrent_PostedTotal_ByCust = "SELECT SUM(IvsTotalAmt-IvsSalesTax-IvsDepositChg) AS TotalForCurrent FROM InvoiceHistory WHERE CustNum = " & passedCustID & " AND ("
	'SQLGetCurrent_PostedTotal_ByCust = SQLGetCurrent_PostedTotal_ByCust & "IvsDate >= '" & StartDateToFind & "' AND "
	'SQLGetCurrent_PostedTotal_ByCust = SQLGetCurrent_PostedTotal_ByCust & "IvsDate <= '" & EndDateToFind & "')"


	SQLGetCurrent_PostedTotal_ByCust = "SELECT SUM(TotalSales) AS TotalForCurrent FROM BI_PostedUnpostedByCustCatPeriod WHERE CustID='" & passedCustID & "' AND "
	SQLGetCurrent_PostedTotal_ByCust = SQLGetCurrent_PostedTotal_ByCust & "PostedOrUnposted = 'P' AND ThisPeriodSeqNumber = " & passedPeriodBeingEvaluated+1

	Set rsGetCurrent_PostedTotal_ByCust = Server.CreateObject("ADODB.Recordset")
	rsGetCurrent_PostedTotal_ByCust.CursorLocation = 3 
	Set rsGetCurrent_PostedTotal_ByCust = cnnGetCurrent_PostedTotal_ByCust.Execute(SQLGetCurrent_PostedTotal_ByCust)

	If not rsGetCurrent_PostedTotal_ByCust.EOF Then resultGetCurrent_PostedTotal_ByCust = rsGetCurrent_PostedTotal_ByCust("TotalForCurrent")

	If Not IsNumeric(resultGetCurrent_PostedTotal_ByCust) Then resultGetCurrent_PostedTotal_ByCust = 0 ' In case there are no results
	
	rsGetCurrent_PostedTotal_ByCust.Close
	set rsGetCurrent_PostedTotal_ByCust= Nothing
	cnnGetCurrent_PostedTotal_ByCust.Close	
	set cnnGetCurrent_PostedTotal_ByCust= Nothing

	
	GetCurrent_PostedTotal_ByCust = resultGetCurrent_PostedTotal_ByCust

End Function

Function TotalCostByPeriodSeq(passedPeriodSeq,passedCustID)

	resultTotalCostByPeriodSeq = ""

	Set cnnTotalCostByPeriodSeq = Server.CreateObject("ADODB.Connection")
	cnnTotalCostByPeriodSeq.open Session("ClientCnnString")
		
	SQLTotalCostByPeriodSeq = "SELECT Sum(TotalCost) AS PeriodTotCost FROM CustCatPeriodSales_ReportData WHERE CustNum = '" & passedCustID & "' AND ThisPeriodSequenceNumber = " & passedPeriodSeq
 
	Set rsTotalCostByPeriodSeq = Server.CreateObject("ADODB.Recordset")
	rsTotalCostByPeriodSeq.CursorLocation = 3 
	Set rsTotalCostByPeriodSeq = cnnTotalCostByPeriodSeq.Execute(SQLTotalCostByPeriodSeq)

	If not rsTotalCostByPeriodSeq.EOF Then resultTotalCostByPeriodSeq = rsTotalCostByPeriodSeq("PeriodTotCost")

	rsTotalCostByPeriodSeq.Close
	set rsTotalCostByPeriodSeq= Nothing
	cnnTotalCostByPeriodSeq.Close	
	set cnnTotalCostByPeriodSeq= Nothing
	
	TotalCostByPeriodSeq = resultTotalCostByPeriodSeq

End Function

Function TotalTPLYAllCats(passedPeriodSeq,passedCustID)

	resultTotalTPLYAllCats = ""

	Set cnnTotalTPLYAllCats = Server.CreateObject("ADODB.Connection")
	cnnTotalTPLYAllCats.open Session("ClientCnnString")
		
	'SQLTotalTPLYAllCats = "SELECT SUM(TotalSales) AS TPLY FROM CustCatPeriodSales WHERE CustNum = '" & passedCustID & "' AND ThisPeriodSequenceNumber = " & passedPeriodSeq - 12
	
	SQLTotalTPLYAllCats = "SELECT Sum(ThisPeriodLastYearSales) AS TPLY FROM CustCatPeriodSales_ReportData WHERE CustNum = '" & passedCustID & "' AND ThisPeriodSequenceNumber = " & passedPeriodSeq
	
	Set rsTotalTPLYAllCats = Server.CreateObject("ADODB.Recordset")
	rsTotalTPLYAllCats.CursorLocation = 3 
	Set rsTotalTPLYAllCats = cnnTotalTPLYAllCats.Execute(SQLTotalTPLYAllCats)

	If not rsTotalTPLYAllCats.EOF Then resultTotalTPLYAllCats = rsTotalTPLYAllCats("TPLY")

	rsTotalTPLYAllCats.Close
	set rsTotalTPLYAllCats= Nothing
	cnnTotalTPLYAllCats.Close	
	set cnnTotalTPLYAllCats= Nothing
	
	TotalTPLYAllCats = resultTotalTPLYAllCats

End Function


%>