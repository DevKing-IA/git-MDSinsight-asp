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
' This page processes the peoplesoft invoice export
' It was written primarily for CCS
' It uses some very specific valuse so if the client id is not 1071 or 1071d it will fail
'If we need to use it with other people, we will need to work on settings fields
'Designed to be launched via a scheduled process
'Self contained page 
'Usage = "http://{xxx}.{domain}.com/directLaunch/invoicing/peoplesoft_export_launch.asp?runlevel=run_now
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
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1 AND (ClientKey ='1071' OR UPPER(ClientKey) ='1071D')"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
	
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3


If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF
	
		ClientKey = TopRecordset.Fields("clientkey")
		PROCESS_REPORT = True
	
		'To begin with, see if this client uses the invoicing module 
		'If they don't then don't bother checking for Nags
		If TopRecordset.Fields("arModule") = 1 Then
	
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

				
					' COMMENTED OUT FOR NOW BECAUSE THIS WILL ONLY RUN FOR 1071 & 1071D AT THE MOMENT
					' Before we do anything, see if this report is turned on for this client
'					Set cnnBizIntel = Server.CreateObject("ADODB.Connection")
'					cnnBizIntel.open MUV_READ("ClientCnnString") 
'					Set rsBizIntel = Server.CreateObject("ADODB.Recordset")
'					rsBizIntel.CursorLocation = 3 
'					Set rsBizIntel = cnnBizIntel.Execute("Select * from Settings_BizIntel")
'					If Not rsBizIntel.EOF Then 
'						If rsBizIntel("CustAnalSum1OnOff") <> 1 or IsNUll(rsBizIntel("CustAnalSum1OnOff")) Then
'							PROCESS_REPORT = False
'							WriteResponse ("Report turned off for client " & ClientKey & " in Settings_BizIntel<BR>")
'						End If
'					Else
'						PROCESS_REPORT = False ' If eof, they dont even have the record, so dont run the report
'						WriteResponse ("No Record Found for client " & ClientKey & " in Settings_BizIntel<BR>")
'					End If
'					set rsBizIntel = nothing
'					cnnBizIntel.close
'					set cnnBizIntel = nothing
'				
				If PROCESS_REPORT = True Then
		
						CreateAuditLogEntry "Peoplesoft Invoice Export","Peoplesoft Invoice Export","Minor",0,"Automated Peoplesoft Invoicing Export ran."					
		
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
					
						WriteResponse ("Begin processing export<br>")
						

						'This direct launch actually does all the processing that the report page does
						'The logic is largely duplicated
						'Changes to one will require changes to the other
						

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	


DIM var_return
var_return=""
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject") 

DIM filename
filename=""


    filename= "c:\home\clientfilesV\" & ClientKey  &"\ftp\outbound\invoice_psoft_"&Right("0" & Day(Now),2)&+Right("0" & Month(Now),2)&YEAR(Now)&".txt"
    if  fso.FileExists(filename) AND Request.Form("replaceFile")="0"  Then
       var_return="{""result"":""1"",""filename"":""invoice_psoft_"&Right("0" & Day(Now),2)&+Right("0" & Month(Now),2)&YEAR(Now)&".txt""}"
        ELSE
            Dim objConn, strFile
            Dim intCampaignRecipientID
            DIM buffer
            DIM APBU, VedorID,Acct,DistributionDescr
            buffer=array()
            
            Set cnnTmpTable = Server.CreateObject("ADODB.Connection")
            cnnTmpTable.open (Session("ClientCnnString"))
            Set rsTmpTable = Server.CreateObject("ADODB.Recordset")

            SQLTmpTable = "SELECT * FROM settings_Reports WHERE reportNumber=8001" 
            Set rsTmpTable = cnnTmpTable.Execute(SQLTmpTable)
            IF NOT rsTmpTable.EOF THEN
                APBU=rsTmpTable("reportspecificdata1")
                VedorID=rsTmpTable("reportspecificdata2")
                Acct=rsTmpTable("reportspecificdata2a")
                DistributionDescr=rsTmpTable("reportspecificdata2b")
                else
		            Response.Write("ERROR")
		            Response.End
            END IF
            rsTmpTable.close

            set rsTmpTable = Nothing
            cnnTmpTable.close
            set cnnTmpTable = Nothing

		StartDate = DateAdd("d",-6,Date())
		EndDate = Date()
		
		SelectedPeriod = ""
		SkipZeroDollar = False
		SkipLessThanZero = True
		SkipLessThanZeroLines = True
		IncludedType = "GT"
		CustomOrPredefined = "Custom"
		Account = ""
		typeOfAccounts = "Chain"
		Chain = "851"
        DuesDateDaysOrSingleDate = ""
            		
        DueDateSingleDate = ""

        DoNotShowDueDate = "CHECKED"


        Description = "System ran  the automated PeopleSoft Invoice Export for Chain 851"
        CreateAuditLogEntry "Peoplesoft Export","Peoplesoft export","Minor",0 ,Description

        'Now get the actual invoice data
        SQLInvoices = "SELECT * FROM InvoiceHistory Where CustNum"
    
    
        SQLInvoices = SQLInvoices &" IN (SELECT CustNum FROM AR_Customer WHERE ChainNum = "&Chain&")"
    
   
            SQLInvoices = SQLInvoices & " AND IvsDate >= '" & StartDate & "' AND IvsDate <= '" & EndDate & "' "

            If SkipZeroDollar = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt <> 0 "
            If SkipLessThanZero = True Then SQLInvoices = SQLInvoices & "AND IvsTotalAmt > 0 "

            If IncludedType <> "" Then SQLInvoices = SQLInvoices & "AND CHARINDEX(IvsType,'" & IncludedType & "') <> 0 "

            SQLInvoices = SQLInvoices & " order by IvsNum"


'Response.Write(SQLInvoices & "<br>")

            Set cnnInvoices = Server.CreateObject("ADODB.Connection")
            cnnInvoices.open (Session("ClientCnnString"))
            Set rsInvoices = Server.CreateObject("ADODB.Recordset")
            rsInvoices.CursorLocation = 3
    
            Set rsInvoices = cnnInvoices.Execute(SQLInvoices)
            If not rsInvoices.Eof Then
    
    
	            Set cnnTmpTable = Server.CreateObject("ADODB.Connection")
	            cnnTmpTable.open (Session("ClientCnnString"))
	            Set rsTmpTable = Server.CreateObject("ADODB.Recordset")
	            rsTmpTable.CursorLocation = 3 

	            TotalAmt = 0
                TotalInvoicesQty=0
	            Do While not rsInvoices.Eof
                    TotalInvoicesQty=TotalInvoicesQty+1
                    TotalAmt = TotalAmt +rsInvoices("IvsTotalAmt")
                    SQLInvoiceDetails =  "Select * from InvoiceHistoryDetail WHERE "
		            SQLInvoiceDetails = SQLInvoiceDetails & "InvoiceHistoryDetail.IvsHistSequence = " & rsInvoices("IvsHistSequence")
			
		            If SkipLessThanZeroLines = True Then SQLInvoiceDetails = SQLInvoiceDetails & "AND InvoiceHistoryDetail.itemPrice <> 0 " 
			
		            SQLInvoiceDetails = SQLInvoiceDetails & " order by IvsHistDetSequence"
			
		            Set cnnInvoiceDetails = Server.CreateObject("ADODB.Connection")
		            cnnInvoiceDetails.open (Session("ClientCnnString"))
		            Set rsInvoiceDetails = Server.CreateObject("ADODB.Recordset")
		            rsInvoiceDetails.CursorLocation = 3 
		            Set rsInvoiceDetails = cnnInvoiceDetails.Execute(SQLInvoiceDetails)

		            If not rsInvoiceDetails.Eof Then
		                SubTot = 0
			            Do While Not rsInvoiceDetails.eof
			                SubTot = SubTot + rsInvoiceDetails("itemQuantity") * rsInvoiceDetails("itemPrice")										
				            rsInvoiceDetails.movenext
			            Loop
                        rsInvoiceDetails.Close
                
		            End If
        
                    rsInvoices.MoveNext
                Loop
    
                'Make file header
                buffer=AddItem(buffer,"C"&APBU&PadNumber(TotalAmt,16,1,"0",1)&PadNumber(TotalInvoicesQty, 3,2,"0",1))
    
                Dim buffData
                rsInvoices.MoveFirst
                Do While not rsInvoices.Eof
    
                    'Make H record 
                    buffer=AddItem(buffer,"H"&PadNumber(VedorID,10,2,"0",1)&PadNumber(rsInvoices("IvsNum"),16,2," ",1)&PadNumber(Month(rsInvoices("IvsDate")),2,2,"0",1) & PadNumber(Day(rsInvoices("IvsDate")),2,2,"0",1)&PadNumber(Year(rsInvoices("IvsDate")),4,2,"0",1)&PadNumber(rsInvoices("IvsTotalAmt"),16,1,"0",1)&"CORPH")

        
                    SQLInvoiceDetails =  "SELECT * FROM InvoiceHistoryDetail WHERE "
		            SQLInvoiceDetails = SQLInvoiceDetails & "InvoiceHistoryDetail.IvsHistSequence = " & rsInvoices("IvsHistSequence")
			
		            If SkipLessThanZeroLines = True Then SQLInvoiceDetails = SQLInvoiceDetails & "AND InvoiceHistoryDetail.itemPrice <> 0 " 
			
		            SQLInvoiceDetails = SQLInvoiceDetails & " order by IvsHistDetSequence"
			
		            Set cnnInvoiceDetails = Server.CreateObject("ADODB.Connection")
		            cnnInvoiceDetails.open (Session("ClientCnnString"))
		            Set rsInvoiceDetails = Server.CreateObject("ADODB.Recordset")
		            rsInvoiceDetails.CursorLocation = 3 
		            Set rsInvoiceDetails = cnnInvoiceDetails.Execute(SQLInvoiceDetails)

		            If not rsInvoiceDetails.Eof Then
		                SubTot = 0
           
			            Do While Not rsInvoiceDetails.eof
			                SubTot = rsInvoiceDetails("itemQuantity") * rsInvoiceDetails("itemPrice")	
			    
                            'buffer=AddItem(buffer,"L"&PadNumber(Replace(rsInvoiceDetails("prodDescription"),"<",""),30,3," ",2)&PadNumber(SubTot, 16,1,"0",1))

                			buffer=AddItem(buffer,"L"&PadNumber(Replace(Replace(Replace(Replace(rsInvoiceDetails("prodDescription"),"<",""),")",""),"(",""),"&",""),30,3," ",2)&PadNumber(SubTot, 16,1,"0",1))
                                
				            dummyprodvar= " " ' Unitl LIJ tells us what to do                
                            buffer=AddItem(buffer,"D"&Acct&PadNumber(getSpecialData(rsInvoices("CustNum"),"DeptID"),8,3," ",2)&PadNumber(getSpecialData(rsInvoices("CustNum"),"GLBU"),5,3," ",2)&PadNumber(DistributionDescr,29,3," ",2)&PadNumber(SubTot, 16,1,"0",1)&PadNumber(dummyprodvar,5,3," ",2)&PadNumber(getSpecialData(rsInvoices("CustNum"),"Project"),6,3," ",2))

				            rsInvoiceDetails.movenext
			            Loop
            
                
		            End If
                    rsInvoiceDetails.Close

                    rsInvoices.MoveNext
                LOOP
    

            END IF
Response.Write("<br><br><br>" & filename & "<br><br>")            
            SET outputfile=fso.CreateTextFile(filename)
            outputfile.write(JOIN(buffer,CHR(13)&CHR(10)))
            outputfile.Close


            var_return="{""result"":""0""}"
    END IF

  

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''						
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	

						
						
					WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				End If ' From PROCESS_Report
			
			End If
			
		End If	
		
	Else ' is the routing module enabled
		WriteResponse ("Skipping the client " & ClientKey & " because the accounts receivable module is not enabled.<BR>")
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




Function AddItem(arr, val)
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
    AddItem = arr
End Function

function PadNumber(number, width,typeNumber,padSymbol,typePad)
    dim padded
    SELECT CASE typeNumber
        CASE 1
             padded = ""&FormatNumber(number,2,0,0,0)
        CASE 2
            padded = cStr(number)
        CASE 3
            padded = number
    END SELECT
   

   while (len(padded) < width)
        IF typePAd=1 THEN
            padded = padSymbol & padded
            ELSE
                padded = padded&padSymbol
        END IF
   wend

   PadNumber = padded
end function

FUNCTION getSpecialData(custID,specialFileldName) 
    DIM retValue
    retValue=" "
    Set SpecialDataConn = Server.CreateObject("ADODB.Connection")
    SpecialDataConn.open (Session("ClientCnnString"))
    Set SpecialDataTable = Server.CreateObject("ADODB.Recordset")
    SpecialDataTable.CursorLocation = 3 
    SpecialDataSql = "SELECT * FROM AR_CustomerBillinfo WHERE CustID="&custID&" AND IncludeOnInvoices=1 AND BillInfoFieldTitle='"& specialFileldName &"'"
    Set SpecialDataTable = SpecialDataConn.Execute(SpecialDataSql)
    IF NOT SpecialDataTable.EOF THEN
        retValue=SpecialDataTable("BillInfoFieldData")
       
    END IF
    SpecialDataTable.Close
    SET SpecialDataTable=Nothing

    SpecialDataConn.Close
    SET SpecialDataConn=Nothing
    getSpecialData=retValue
END FUNCTION

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