<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_BizIntel.asp"-->
<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->

<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
 
<%
'Response.Buffer = True  <-----
'Response.Expires = 0  <-----	These lines commented purposely. They keep the page from close when launched automatically. Can't use them.
'Response.Clear  <-----


'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page
'Usage = "http://{xxx}.{domain}.com/directLaunch/bizintel/mcs_rebuild_helper_launch.asp?runlevel=run_now
Server.ScriptTimeout = 25000

Dim EntryThread
Dim TotalMCSClients, TotalMCSCommitment, TotalSalesAllMCSCustomers
Dim TotalCustomersUnderButRecovered, TotalCustomersOver, TotalOverDollars, TotalCustomersUnder, TotalUnderButRecoveredDeficitDollars 
Dim TotalUnderDollars, TotalLVFLastMonth, TotalPendingLVF, TotalCustomersZeroSales ,TotalZeroSalesCommitment
Dim ReportDate

ReportDate = Month(Now()) & "/01/" & Year(Now())


'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then
	Response.Write("Improper usage, no run level was specified in the query string")	
	response.end
End IF 

'This single page loops through and handles alerts for ALL clients
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
	
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3


If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF
	
		ClientKey = TopRecordset.Fields("clientkey")
	
		'To begin with, see if this client uses the Biz Intel 
		'If they don't then don't bother running for then
		
		Response.Write("Seeing if we need to run this for " & ClientKey & "<br>")
		
		If TopRecordset.Fields("biModule") = "Enabled" Then
	
			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then

												
				Call SetClientCnnString
				
				Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
				

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
					
	
					CreateAuditLogEntry "MCS Rebuild Helper Launch","MCS Rebuild Helper Launch","Minor",0,"MCS Rebuild Helper Launch ran."					
	
					WriteResponse "<font color='purple' size='24'>Start processing " & ClientKey  & "</font><br>"
	
					
					If MUV_READ("cnnStatus") = "OK" Then ' else it loops
					
						Response.Write("blah,blah,blah blah,blah,blah blah,blah,blah blah,blah,blah blah,blah,blah <br>")
						
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''

							
						dummy = MUV_WRITE("MCSFLAG","0")
					
						Set cnnMCSData = Server.CreateObject("ADODB.Connection")
						cnnMCSData.open (Session("ClientCnnString"))
						Set rsMCSData = Server.CreateObject("ADODB.Recordset")
						Set rsMCSDataForUpdating = Server.CreateObject("ADODB.Recordset")
					
						If passedCustID = "" Then
							SQLMCSData = "DELETE FROM BI_MCSData"
						Else
							SQLMCSData = "DELETE FROM BI_MCSData WHERE CustID = '" & passedCustID & "'"
						End If
						Set rsMCSData= cnnMCSData.Execute(SQLMCSData)
						
						If passedCustID = "" Then
							SQLMCSData = "INSERT INTO BI_MCSData (CustID) SELECT CustNum FROM AR_Customer WHERE MonthlyContractedSalesDollars <> 0 AND AcctStatus='A'" 
						Else
							SQLMCSData = "INSERT INTO BI_MCSData (CustID) SELECT CustNum FROM AR_Customer WHERE MonthlyContractedSalesDollars <> 0 AND AR_Customer.CustNum = '"  & passedCustID & "' AND AcctStatus='A'"
						End If
						Set rsMCSData= cnnMCSData.Execute(SQLMCSData)
						
						'Now begin with all the aggregate numbers
						SQLMCSData = "SELECT * FROM BI_MCSData"
						Set rsMCSData= cnnMCSData.Execute(SQLMCSData)
						
						If NOT rsMCSData.EOF Then
							Do While Not rsMCSData.EOF
								
								Month1Sales_NoRent = 0
								Month2Sales_NoRent = 0
								Month3Sales_NoRent = 0
								Month3Cost_NoRent = 0
								LVFHolder = 0
								LVFHolderCurrent = 0
								TotalEquipmentValue = 0
								CurrentHolder = 0
								RentalHolder = 0
								PendingLVF = 0
								
								PendingLVF = cdbl(PendingLVFByCust(rsMCSData("CustID")))
								
								
								Month3Cost_NoRent = TotalCostByCustByMonthByYear_NoRent(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
								If NOT IsNumeric(Month3Cost_NoRent) Then Month3Cost_NoRent = 0
								Month1Sales_NoRent = TotalSalesByCustByMonthByYear_NoRentals(rsMCSData("CustID"),Month(DateAdd("m",-3,ReportDate)),Year(DateAdd("m",-3,ReportDate)))
								Month2Sales_NoRent = TotalSalesByCustByMonthByYear_NoRentals(rsMCSData("CustID"),Month(DateAdd("m",-2,ReportDate)),Year(DateAdd("m",-2,ReportDate)))
								Month3Sales_NoRent = TotalSalesByCustByMonthByYear_NoRentals(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
					
								' Remove LVF from Monthly sales
								Month1LVF = TotalPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-3,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
								Month2LVF = TotalPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-2,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
								Month3LVF = TotalPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
					
								If Month1LVF > 0 Then Month1Sales_NoRent = Month1Sales_NoRent - Month1LVF 
								If Month2LVF > 0 Then Month2Sales_NoRent = Month2Sales_NoRent - Month2LVF 
								If Month3LVF > 0 Then Month3Sales_NoRent = Month3Sales_NoRent - Month3LVF 
								
								Month1XSF = TotalXSFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-3,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
								Month2XSF = TotalXSFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-2,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
								Month3XSF = TotalXSFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
					
								If Month1XSF > 0 Then Month1Sales_NoRent = Month1Sales_NoRent - Month1XSF 
								If Month2XSF > 0 Then Month2Sales_NoRent = Month2Sales_NoRent - Month2XSF 
								If Month3XSF > 0 Then Month3Sales_NoRent = Month3Sales_NoRent - Month3XSF 
									
								CurrentXSF = 0	
								CurrentXSF = TotalXSFByCustByMonthByYear(rsMCSData("CustID"),Month(ReportDate),Year(ReportDate))	
								
					
								LVFHolder = TotalPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
								
								
								LVFHolderCurrent = TotalPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(ReportDate),Year(ReportDate))
								TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(rsMCSData("CustID"))
								
								' Must subtract any rentals from Current moth$
								CurrentHolder = GetCurrent_PostedTotal_ByCust(rsMCSData("CustID"),PeriodSeqBeingEvaluated) + GetCurrent_UnPostedTotal_ByCust(rsMCSData("CustID"),PeriodSeqBeingEvaluated)
								CurrentRent = TotalSalesByCustByMonthByYear_RentalsOnly(rsMCSData("CustID"),Month(ReportDate),Year(ReportDate))
								CurrentHolder = CurrentHolder - CurrentRent
								
								
								RentalHolder = TotalSalesByCustByMonthByYear_RentalsOnly(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
								'If Month3XSF <> 0 Then RentalHolder = RentalHolder +  Month3XSF 
					
								Month1_Cat21Holder = TotalCat21ByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-3,ReportDate)),Year(DateAdd("m",-1,ReportDate)))			
								Month2_Cat21Holder = TotalCat21ByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-2,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
								Month3_Cat21Holder = TotalCat21ByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
								
								
								'Got em all, update the record
								SQLMCSDataForUpdating = "UPDATE BI_MCSData SET "
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & "Month1Sales_NoRent = " & Month1Sales_NoRent
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month2Sales_NoRent = " & Month2Sales_NoRent
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month3Sales_NoRent = " & Month3Sales_NoRent						
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month3Cost_NoRent = " & Month3Cost_NoRent
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", LVFHolder = " & LVFHolder 
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", LVFHolderCurrent = " & LVFHolderCurrent 
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", TotalEquipmentValue = " & TotalEquipmentValue 
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", CurrentHolder = " & CurrentHolder 	
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", RentalHolder = " & RentalHolder 
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month1Cat21Sales = " & Month1_Cat21Holder 
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month2Cat21Sales = " & Month2_Cat21Holder 
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month3Cat21Sales = " & Month3_Cat21Holder 			
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month1XSF = " & Month1XSF 
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month2XSF = " & Month2XSF 
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", Month3XSF = " & Month3XSF 
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", CurrentXSF = " & CurrentXSF 
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", PendingLVF = " & PendingLVF 
								If GetCustChainIDByCustID(rsMCSData("CustID")) <> "" Then
									SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", ChainID = '" & GetCustChainIDByCustID(rsMCSData("CustID")) & "' "
								End If	
								If GetChainDescByChainNum(GetCustChainIDByCustID(rsMCSData("CustID"))) <> "" Then 
									SQLMCSDataForUpdating = SQLMCSDataForUpdating  & ", ChainName = '" & GetChainDescByChainNum(GetCustChainIDByCustID(rsMCSData("CustID"))) & "' "
								End If
					
								
								
								
								SQLMCSDataForUpdating = SQLMCSDataForUpdating  & " WHERE CustID = " & rsMCSData("CustID")
								
								'Response.Write(SQLMCSDataForUpdating & "<br>")
								
								Set rsMCSDataForUpdating = cnnMCSData.Execute(SQLMCSDataForUpdating)
							
								rsMCSData.MoveNext
							Loop
						
						End If
					
					
					cnnMCSData.Close
					Set rsMCSData = Nothing
					Set cnnMCSData = Nothing




''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
						
	
										
						WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				
				End If
			
		End If	
		
	Else ' is the biz in tel  module enabled
	
		Call SetClientCnnString
				
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
			
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


		WriteResponse ("Skipping the client " & ClientKey & " because the Biz Intel module is not enabled.<BR>")
		
	End If ' is the Service  module enabled
	
	TopRecordset.movenext
	
	Loop
	
	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")

'************************************************************************************
'************************************************************************************
'Subs and funcs begin here
'************************************************************************************

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
	SQL = SQL & ",'MCS Rebuild Helper'"
	SQL = SQL & ",'/directlaunch/bizintel/mcs_rebuild_helper_launch.asp'"
	SQL = SQL & ",'"  & passedLogEntry & "'"
	SQL = SQL & ")"
	
	'Response.write("<BR>" & SQL & "<BR>")
	
	Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
	cnnAuditLog.open Session("ClientCnnString") 
	Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
	rsAuditLog.CursorLocation = 3 
	
	Set rsAuditLog = cnnAuditLog.Execute(SQL)

	set rsAuditLog = nothing
	cnnAuditLog.close
	set cnnAuditLog = nothing

End Sub


Sub DelayResponse(numberOfseconds)
 Dim WshShell
 Set WshShell=Server.CreateObject("WScript.Shell")
 WshShell.Run "waitfor /T " & numberOfSecond & "SignalThatWontHappen", , True
End Sub

Function CheckRemoteURL(fileURL)
    ON ERROR RESUME NEXT
    Dim xmlhttp

    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

    xmlhttp.open "GET", fileURL, False
    xmlhttp.send
    If(Err.Number<>0) then
        Response.Write "Could not connect to remote server"
    else
        Select Case Cint(xmlhttp.status)
            Case 200, 202, 302
                Set xmlhttp = Nothing
                CheckRemoteURL = True
            Case Else
                Set xmlhttp = Nothing
                CheckRemoteURL = False
        End Select
    end if
    ON ERROR GOTO 0
End Function


'************************************************************************************
'************************************************************************************
'Subs and funcs end here
'************************************************************************************


%>