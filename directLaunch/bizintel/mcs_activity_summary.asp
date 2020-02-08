<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear
Dim PageNum, RowCount, FontSizeVar

Server.ScriptTimeout = 25000

'These DIMs must be here

Dim TotalMCSClients, TotalMCSCommitment, TotalSalesAllMCSCustomers
Dim TotalCustomersUnderButRecovered, TotalCustomersOver, TotalOverDollars, TotalCustomersUnder, TotalUnderButRecoveredDeficitDollars 
Dim TotalUnderDollars, TotalLVFLastMonth, TotalPendingLVF, TotalCustomersZeroSales ,TotalZeroSalesCommitment
Dim ReportDate

Dim TotalCustsMCSAdded, TotalCustsMCSRemoved, TotalNetMCSChange, TotalNetLVFChange, TotalNoAction, TotalFollowup, TotalCustsMCSAddedDollars, TotalCustsMCSRemovedDollars
Dim TotalNumCustsInvoiced, TotalActedUpon, TotalNOTActedUpon, TotalNotActedUponMCSDollars
Dim TotalNUMMCSChange, TotalNUMLVFChange ,TotalMsgsSent, TotalLVFInvoicedAmount 

FontSizeVar = 9
PageNum = 0
NoBreak = False
'Adjust = -3
'MAdjust = -1
PageWidth = 1100

Response.Write("<style type='text/css'>")
Response.Write("mark {")
Response.Write("    background-color: yellow;")
Response.Write("    color: black;")
Response.Write("}")
Response.Write("</style>")

ReportDate = Month(Now()) & "/01/" & Year(Now())


Slsmn = Request.QueryString("sls") ' Gets passed if only being run for one salesman

' Will rebuild by default unless 0 is passed in
' When running from the luancher, should be set to 0 to prevent timeouts
' BUT the rebuild helper page MUST BE RUN FIRST to ensure data is good
Rbld = Request.QueryString("rbld")
If Rbld = "" Then Rbld = 1

%>
<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_BizIntel.asp"-->
<!--#include file="../../inc/InsightFuncs_Equipment.asp"-->
<%dummy=MUV_Write("ClientID","") 'Need this here

ClientKey = Request.QueryString("c")

SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open InsightCnnString
'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3
'Response.Write("SQL:" & SQL& "<br>")

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Inventory Need To Know Report<%
Else
	ClientCnnStringvar = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	ClientCnnStringvar = ClientCnnStringvar  & ";Database=" & Recordset.Fields("dbCatalog")
	ClientCnnStringvar = ClientCnnStringvar & ";Uid=" & Recordset.Fields("dbLogin")
	ClientCnnStringvar = ClientCnnStringvar & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	dummy = MUV_Write("ClientCnnString",ClientCnnStringvar)
	Recordset.close
	Connection.close	
End If	

Session("ClientCnnString") = MUV_READ("ClientCnnString")


'This is here so we only open it once for the whole page
Set cnn_Settings_BizIntel = Server.CreateObject("ADODB.Connection")
cnn_Settings_BizIntel.open (MUV_READ("ClientCnnString"))
Set rs_Settings_BizIntel = Server.CreateObject("ADODB.Recordset")
rs_Settings_BizIntel.CursorLocation = 3 
SQL_Settings_BizIntel = "SELECT * FROM Settings_BizIntel"
Set rs_Settings_BizIntel = cnn_Settings_BizIntel.Execute(SQL_Settings_BizIntel)
If not rs_Settings_BizIntel.EOF Then
	MCSActivitySummaryOnOff = rs_Settings_BizIntel("MCSActivitySummaryOnOff")
	MCSUseAlternateHeader = rs_Settings_BizIntel("MCSUseAlternateHeader")	
	If MCSUseAlternateHeader <> 1 Then MCSUseAlternateHeader = 0
Else
	MCSActivitySummaryOnOff = 0
End If
Set rs_Settings_BizIntel = Nothing
cnn_Settings_BizIntel.Close
Set cnn_Settings_BizIntel = Nothing

If MCSActivitySummaryOnOff <> 1 Then
	%>MDS Insight: The MCS Activity Summary is not turned on.
	<%
	Response.End
End IF


If Rbld <> 0 Then Call RebuildMCSData

Call CalcSummaryInformation

Call CalcSummaryInformation2

%>
<style>
	.negative{
		color:red;	
	}

	.positive{
		color:green;	
	}

	.neutral{
		color:black;
	}
</style>

<body bgcolor="#FFFFFF" text="#000000" link="#000080" vlink="#000080" alink="#000080" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<table border="0" width="<%=PageWidth%>" align="center">
	<tr>
		<td width="100%" align="center">

			<%
			'*******************************************************
			'*** This section is the first page which prints all the
			'*** MCS Activity summary info
			'*** First it does all the calculations
			'*** This code was copied from the actual MCS report page
			'********************************************************
		
			Call PageHeader

			LinesPerPage = 42
						
			FontSizeVar = 9
			
			%>
			
			<br><br><br>
			<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
				<tr>
					<td colspan ="8">
					<font face="Consolas">
					<hr>
					<center><h2>MCS Activity Summary&nbsp;<%=Month(Now()) & "/1/" & Right(Year(Now()),2)%>&nbsp;-&nbsp;<%=Month(Now()) & "/" & Day(Now()) & "/" & Right(Year(Now()),2)%></h2></center>
					<hr>
					</font>
					</td>
				</tr>
				<tr>
					<td colspan ="8"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
				</tr>
				<tr>
					<td colspan ="8"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
				</tr>
				<tr>
					<td colspan ="8"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
				</tr>
				<tr>
					<td colspan ="8"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
				</tr>
				<tr>
					<td colspan ="3">
					<font face="Consolas">
					<hr>
					<center><h4><%=MonthName(Month(DateAdd("m",-1,ReportDate)))%>&nbsp;<%= Year(DateAdd("m",-1,ReportDate))%>&nbsp;MCS Analysis Data</h4></center>
					<hr>
					</font>
					</td>
					<td colspan ="5"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
				</tr>
				<% If MCSUseAlternateHeader = 1 Then %>
					<!--#include file="./mcsheader_alternate.asp"-->
				<% Else %>
					<!--#include file="./mcsheader_standard.asp"-->
				<% End If %>
				<br/><br/>
			</table>
	

<%			'*************************************************
			'*** SECOND SUMMARY SECTION SECOND SUMMARY SECTION
			'************************************************* %>
			
			<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table33" align="center">	
				<tr>
					<td colspan ="3">
					<font face="Consolas">
					<hr>
					<center><h4>Activity Summary This MTD (<%= MonthName(Month(ReportDate)) %>)</h4></center>
					<hr>
					</font>
					</td>
					<td colspan ="5"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
				</tr>
				<% If MCSUseAlternateHeader = 1 Then %>
					<!--#include file="./mcssubheader_alternate.asp"-->
				<% Else %>
					<!--#include file="./mcssubheader_standard.asp"-->
				<% End If %>
				<br/><br/>
			</table>

<%			
			RowCount = 38
		
			'*******************************************************
			'*** END END END END END END END END END END END END END 
			'*** This section is the first page which prints all the
			'*** MCS summary info
			'*******************************************************
%>
		</td>
	</tr>
	<tr>
		<td> <%
			Call Footer
			
			
			
			
			'*************************************************
			'*** This section is the DETAIL section of all the
			'*** MCS actions taken this MTD
			'*************************************************
			TotalCustsReported = 0
			

			SQL = "SELECT * "
			SQL = SQL & " FROM AR_Customer INNER JOIN "
			SQL = SQL & " BI_MCSData ON BI_MCSData.CustID = AR_Customer.CustNum INNER JOIN "
			SQL = SQL & " (SELECT CustID, MAX(RecordCreationDateTime) AS RecordCreationDateTimeForSort "
			SQL = SQL & " FROM BI_MCSActions "
			SQL = SQL & " WHERE (MONTH(RecordCreationDateTime) = MONTH(GETDATE())) "
			SQL = SQL & " GROUP BY CustID) AS derivedtbl_1 ON AR_Customer.CustNum = derivedtbl_1.CustID "
			SQL = SQL & " WHERE (AR_Customer.MonthlyContractedSalesDollars IS NOT NULL) "
			SQL = SQL & " ORDER BY derivedtbl_1.RecordCreationDateTimeForSort DESC"

			
			Set cnnMCSActivity = Server.CreateObject("ADODB.Connection")
			cnnMCSActivity.open (MUV_READ("ClientCnnString"))
			Set rsMCSActivtySummary = Server.CreateObject("ADODB.Recordset")
			rsMCSActivtySummary.CursorLocation = 3 
			rsMCSActivtySummary.Open SQL, cnnMCSActivity
			
			If Not rsMCSActivtySummary.EOF Then
				
				Call PageHeader
				Call SubHeader
			
				Do While Not rsMCSActivtySummary.EOF

					FontSizeVar = 9
					LinesPerPage = 32
					
					%>
					<tr>
					<%
								
			ShowThisRecord = True
				
			PrimarySalesMan =  ""
			SecondarySalesMan =  ""
			SelectedCustomerID = rsMCSActivtySummary("CustNum")
			CustName = rsMCSActivtySummary("Name")
			CustMonthlyContractedSalesDollars = 0
			InstallDate = ""
			EnrollmentDate = ""
			
			PrimarySalesMan = rsMCSActivtySummary("Salesman")
			SecondarySalesMan = rsMCSActivtySummary("SecondarySalesman")
			CustMonthlyContractedSalesDollars = rsMCSActivtySummary("MonthlyContractedSalesDollars")
			InstallDate = rsMCSActivtySummary("InstallDate")
			MaxMCSCharge = rsMCSActivtySummary("MaxMCSCharge")
			EnrollmentDate =  rsMCSActivtySummary("MCSEnrollmentDate")

			Month3Sales_NoRent = rsMCSActivtySummary("Month3Sales_NoRent") - rsMCSActivtySummary("Month3Cat21Sales") 

			If Month3Sales_NoRent >= CustMonthlyContractedSalesDollars Then ShowThisRecord = False

			'If Month3Sales_NoRent > 0 Then ShowThisRecord = False

			VarianceHolder = Month3Sales_NoRent - CustMonthlyContractedSalesDollars 
			CurrentHolder = rsMCSActivtySummary("CurrentHolder")
			
		    ' Calc under by the current month recovered the deficit
			If VarianceHolder < 0 Then 'Meaning they have a variance
				If CurrentHolder >= CustMonthlyContractedSalesDollars + ABS(VarianceHolder)  Then
					ShowThisRecord = False
				End If
			End If

			 If ABS(VarianceHolder) < 100 Then
				If Month3Sales_NoRent <> 0 Then
					VariancePercentHolder = 100 - ((Month3Sales_NoRent/CustMonthlyContractedSalesDollars) * 100) 
				End If
				VariancePercentHolder  = VariancePercentHolder  * -1
				If ABS(VariancePercentHolder) < 10 Then
					ARCount = ARCount + 1
					ShowThisRecord = False
				End If
			End If

			If GetLastMCSActionDateByMonthByYearByCust(SelectedCustomerID , Month(Now()), Year(Now())) = "" Then ShowThisRecord = False

			If ShowThisRecord <> False Then
			
				Month1Sales_NoRent = rsMCSActivtySummary("Month1Sales_NoRent") - rsMCSActivtySummary("Month1Cat21Sales") 
				Month2Sales_NoRent = rsMCSActivtySummary("Month2Sales_NoRent") - rsMCSActivtySummary("Month2Cat21Sales") 
				
				ThreePPSales = Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent
				
				Month3Cost_NoRent = rsMCSActivtySummary("Month3Cost_NoRent") 
				
				Month3GP = Month3Sales_NoRent - Month3Cost_NoRent
				If Not IsNumeric(Month3GP) Then Month3GP  = 0
			
				ThreePPAvgSales = (Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent) / 3
				
				ShortageHolder = ThreePPSales - (CustMonthlyContractedSalesDollars * 3)
				
				LVFHolder = rsMCSActivtySummary("LVFHolder") 
				
				LVFHolderCurrent = rsMCSActivtySummary("LVFHolderCurrent") 
				
				TotalEquipmentValue = rsMCSActivtySummary("TotalEquipmentValue")
				
				TotalCustsReported = TotalCustsReported + 1

				'' Now handle the notes spannig multiple lines
				' Must do this here because it should break before showing any of the detail ine
				NumNoteLines = 0 
				NumNoteLines = int(Len(LastActNote ) / 75)
				If Len(LastActNote ) MOD 75 <> 0 Then NumNoteLines = NumNoteLines + 1
				' If it's going to span a page break, break it now
				If RowCount + NumNoteLines > LinesPerPage Then
					%>
					</tr>
					<%
					Call Footer
					Call PageHeader
					Call SubHeader
					%>
					<tr>
					<%
				End If
					
				%>
				<td width="7%">
					<%
						LastMCSActionDate = GetLastMCSActionDateByMonthByYearByCust(SelectedCustomerID , Month(Now()), Year(Now()))
						'LastMCSActionDate = GetLastMCSActionDateByMonthByYearByCust(SelectedCustomerID , 12, 2018)
						If LastMCSActionDate <> "" Then
							LastMCSActionDate = cDate(LastMCSActionDate ) 
							eYear = Year(LastMCSActionDate )
							If Month(LastMCSActionDate ) < 10 Then eMonth = "0" & Month(LastMCSActionDate ) else eMonth = Month(LastMCSActionDate )
							If Day(LastMCSActionDate ) < 10 Then eDay = "0" & Day(LastMCSActionDate ) else eDay = Day(LastMCSActionDate )
							LastMCSActionDispayableDate = eMonth & "/" & eDay  & "/" & Right(eYear,2)
							'LastMCSActionDispayableDate = cDate(LastMCSActionDispayableDate ) 
						Else
							LastMCSActionDispayableDate = "none"
						End If
					%>
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= LastMCSActionDispayableDate %></font>
				</td>
				<td width="5%">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= SelectedCustomerID %></font>
				</td>
				<td width="24%">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><% If Len(rsMCSActivtySummary("Name")) > 35 Then Response.Write(Left(rsMCSActivtySummary("Name"),35)) Else Response.Write(rsMCSActivtySummary("Name"))  %></font>
				</td>
				<td width="7%" align="right">
					<% If rsMCSActivtySummary("Month3Sales_NoRent") < 0 Then %>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt" class="negative"><%= FormatCurrency(rsMCSActivtySummary("Month3Sales_NoRent"),0)%>&nbsp;</font>
					<% ElseIf rsMCSActivtySummary("Month3Sales_NoRent") > 0 Then %>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= FormatCurrency(rsMCSActivtySummary("Month3Sales_NoRent"),0)%>&nbsp;</font>
					<% ElseIf rsMCSActivtySummary("Month3Sales_NoRent") = 0 Then
						If ClientKey = "1071" or Ucase(ClientKey)="1071D" Then %>
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><mark>ZERO</mark>&nbsp;</font>					
						<% Else %>
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= FormatCurrency(rsMCSActivtySummary("Month3Sales_NoRent"),0)%>&nbsp;</font>												
						<% End If %>	
					<% End If %>
				</td>
				<td width="5%" align="right">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= FormatCurrency(CustMonthlyContractedSalesDollars ,0)%>&nbsp;</font>
				</td>
				<td width="7%" align="right">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt" class="negative"><%= FormatCurrency(VarianceHolder,0)%>&nbsp;</font>
				</td>
				<%
				'EnrollmentDate Date
				EnrollmentDate = cDate(EnrollmentDate) 
				eYear = Year(EnrollmentDate)
				If Month(EnrollmentDate) < 10 Then eMonth = "0" & Month(EnrollmentDate) else eMonth = Month(EnrollmentDate)
				If Day(EnrollmentDate) < 10 Then eDay = "0" & Day(EnrollmentDate) else eDay = Day(EnrollmentDate)
				EnrollmentDispayableDate = eMonth & "/" & eDay  & "/" & Right(eYear,2)
				'EnrollmentDispayableDate  = cDate(EnrollmentDispayableDate) 
				%>
				<td width="7%">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= EnrollmentDispayableDate  %></font>
				</td>
				<%
				LastActNote = GetLastMCSActionNoteByMonthByYearByCust(SelectedCustomerID , Month(Now()), Year(Now())) 
				LastActNote = Replace(LastActNote,"Action selected: ","")
				
				If ClientKey = "1071" or Ucase(ClientKey)="1071D" Then
					If LastActNote = "No action necessary at this time" Then
						LastActNote ="No Action Necessary"
						'Now append the reason
						LastActNote = LastActNote & " - Reason: " & GetMCSReasonByReasonNum(GetLastMCSActionNoteReasonByMonthByYearByCust(SelectedCustomerID , Month(Now()), Year(Now())))
					End If
					If Left(LastActNote,12) = "Client added" Then
						LastActNote = Trim(LastActNote)
						LastActNoteHolder = ""
						
						LastActNoteHolder = "ADD - MCS: "
						
						LastActMCSHolder = ""
						x = Instr(LastActNote,"$")
						For z = x to Len(LastActNote)
							If Mid(LastActNote,z,1) = "." Then Exit For
							LastActMCSHolder = LastActMCSHolder & Mid(LastActNote,z,1)
						Next 
						
						LastActNoteHolder  = LastActNoteHolder  & LastActMCSHolder 

						LastActLVFHolder = ""
						LastActLVFHolder = Right(LastActNote ,Len(LastActNote)-InstrRev(LastActNote,"$"))
						
						LastActNoteHolder  = LastActNoteHolder  & "  LVF: $" & LastActLVFHolder 
						LastActNote = LastActNoteHolder  
					End If						
				
				End If
				
				
				
				'' Now handle the notes spannig multiple lines - the actual printing
				NumNoteLines = 0 
				NumNoteLines = int(Len(LastActNote ) / 70)
				If Len(LastActNote ) MOD 70 <> 0 Then NumNoteLines = NumNoteLines + 1

				ReDim DetailLinesArray(NumNoteLines)
				For x = 0 to NumNoteLines -1
					If x = 0 Then
						DetailLinesArray(x) = Mid(LastActNote ,1,70)
					ElseIf x = 1 Then
						DetailLinesArray(x) = Mid(LastActNote ,71,70)
					ElseIf x = NumNoteLines -1 Then 
						DetailLinesArray(x) = Mid(LastActNote ,(x*70)+1,Len(LastActNote)- ((x*70)+1))
					Else
						DetailLinesArray(x) = Mid(LastActNote ,(x*70)+1,70)
					End If
				Next

				For z = 0 to Ubound(DetailLinesArray) -1
					If z > 0 Then Response.Write("<tr><td colspan='7'>&nbsp;</td>")
					%>
					<td width="43%" style="white-space: nowrap;"> 
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= DetailLinesArray(z) %></font>
					</td> 
					<%
					If z > 0 Then Response.Write("</tr>")
					RowCount = RowCount + 1
				Next
				
				%>
				</tr>
				<% 
				'We don't really need to add blank lines for spacing if there are multiple lines of notes
'				Response.Write("<tr><td colspan='8'>&nbsp;</td></tr>")
				
				RowCount = RowCount + 1
			End If

			rsMCSActivtySummary.Movenext	

			If RowCount > LinesPerPage Then
				%></table><%
				Call Footer
				Call PageHeader
				Call SubHeader
			End If
				
			Loop


			If TotalCustsReported = 0 Then ' There have been no action this month
				%>
				<td colspan="7">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">There has been no activity this month-to-date</font>
				</td>
				<%
			End If
				
			'NoBreak = True

			Call Footer	
		End If


			'*************************************************
			'***END END END END END END END END END END END 
			'*** This section is the DETAIL section of all the
			'*** MCS actions taken this MTD
			'*************************************************


		'	Call Footer
			
			
			
			
			'****************************************************
			'*** This section is the NEW DETAIL section of all the
			'*** MCS customers which have NOT been reviewed
			'****************************************************
			TotalCustsReported = 0
			
			SQL = "SELECT * FROM AR_Customer INNER JOIN BI_MCSData ON BI_MCSData.CustID = AR_Customer.CustNum WHERE MonthlyContractedSalesDollars IS NOT Null " 
			SQL = SQL & " AND AR_Customer.CustNum NOT IN (SELECT CustID AS CustNum FROM BI_MCSActions WHERE MONTH(RecordCreationDateTime) = MONTH(GETDATE())) "
			SQL = SQL & " ORDER BY (Month3Sales_NoRent - MonthlyContractedSalesDollars )"

'Response.Write(SQL)
			
			Set cnnMCSActivity = Server.CreateObject("ADODB.Connection")
			cnnMCSActivity.open (MUV_READ("ClientCnnString"))
			Set rsMCSActivtySummary = Server.CreateObject("ADODB.Recordset")
			rsMCSActivtySummary.CursorLocation = 3 
			rsMCSActivtySummary.Open SQL, cnnMCSActivity
			
			If Not rsMCSActivtySummary.EOF Then
				
				Call PageHeader
				Call SubHeaderNoActivity
			
				Do While Not rsMCSActivtySummary.EOF

					FontSizeVar = 9
					LinesPerPage = 32
					
					%>
					<tr>
					<%
								
			ShowThisRecord = True
				
			PrimarySalesMan =  ""
			SecondarySalesMan =  ""
			SelectedCustomerID = rsMCSActivtySummary("CustNum")
			CustName = rsMCSActivtySummary("Name")
			CustMonthlyContractedSalesDollars = 0
			InstallDate = ""
			EnrollmentDate = ""
			
			CustMonthlyContractedSalesDollars = rsMCSActivtySummary("MonthlyContractedSalesDollars")
			InstallDate = rsMCSActivtySummary("InstallDate")
			MaxMCSCharge = rsMCSActivtySummary("MaxMCSCharge")
			EnrollmentDate =  rsMCSActivtySummary("MCSEnrollmentDate")

			Month3Sales_NoRent = rsMCSActivtySummary("Month3Sales_NoRent") - rsMCSActivtySummary("Month3Cat21Sales") 

			If Month3Sales_NoRent >= CustMonthlyContractedSalesDollars Then ShowThisRecord = False


			VarianceHolder = Month3Sales_NoRent - CustMonthlyContractedSalesDollars 
			CurrentHolder = rsMCSActivtySummary("CurrentHolder")
			
		    ' Calc under by the current month recovered the deficit
			If VarianceHolder < 0 Then 'Meaning they have a variance
				If CurrentHolder >= CustMonthlyContractedSalesDollars + ABS(VarianceHolder)  Then
					ShowThisRecord = False
				End If
			End If

			 If ABS(VarianceHolder) < 100 Then
				If Month3Sales_NoRent <> 0 Then
					VariancePercentHolder = 100 - ((Month3Sales_NoRent/CustMonthlyContractedSalesDollars) * 100) 
				End If
				VariancePercentHolder  = VariancePercentHolder  * -1
				If ABS(VariancePercentHolder) < 10 Then
					ARCount = ARCount + 1
					ShowThisRecord = False
				End If
			End If


			If ShowThisRecord <> False Then
			
				Month1Sales_NoRent = rsMCSActivtySummary("Month1Sales_NoRent") - rsMCSActivtySummary("Month1Cat21Sales") 
				Month2Sales_NoRent = rsMCSActivtySummary("Month2Sales_NoRent") - rsMCSActivtySummary("Month2Cat21Sales") 
				
				ThreePPSales = Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent
				
				Month3Cost_NoRent = rsMCSActivtySummary("Month3Cost_NoRent") 
				
				Month3GP = Month3Sales_NoRent - Month3Cost_NoRent
				If Not IsNumeric(Month3GP) Then Month3GP  = 0
			
				ThreePPAvgSales = (Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent) / 3
				
				ShortageHolder = ThreePPSales - (CustMonthlyContractedSalesDollars * 3)
				
				LVFHolder = rsMCSActivtySummary("LVFHolder") 
				
				LVFHolderCurrent = rsMCSActivtySummary("LVFHolderCurrent") 
				
				TotalEquipmentValue = rsMCSActivtySummary("TotalEquipmentValue")
				
				TotalCustsReported = TotalCustsReported + 1

				'' Now handle the notes spannig multiple lines
				' Must do this here because it should break before showing any of the detail ine
				NumNoteLines = 0 
				NumNoteLines = int(Len(LastActNote ) / 75)
				If Len(LastActNote ) MOD 75 <> 0 Then NumNoteLines = NumNoteLines + 1
				' If it's going to span a page break, break it now
				If RowCount + NumNoteLines > LinesPerPage Then
					%>
					</tr>
					<%
					Call Footer
					Call PageHeader
					Call SubHeaderNoActivity
					%>
					<tr>
					<%
				End If
					
				%>
				<td width="5%">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= SelectedCustomerID %></font>
				</td>
				<td width="31%">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><% If Len(rsMCSActivtySummary("Name")) > 35 Then Response.Write(Left(rsMCSActivtySummary("Name"),35)) Else Response.Write(rsMCSActivtySummary("Name"))  %></font>
				</td>
				<td width="7%" align="right">
					<% If rsMCSActivtySummary("Month3Sales_NoRent") < 0 Then %>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt" class="negative"><%= FormatCurrency(rsMCSActivtySummary("Month3Sales_NoRent"),0)%>&nbsp;</font>
					<% ElseIf rsMCSActivtySummary("Month3Sales_NoRent") > 0 Then %>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= FormatCurrency(rsMCSActivtySummary("Month3Sales_NoRent"),0)%>&nbsp;</font>
					<% ElseIf rsMCSActivtySummary("Month3Sales_NoRent") = 0 Then
						If ClientKey = "1071" or Ucase(ClientKey)="1071D" Then %>
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><mark>ZERO</mark>&nbsp;</font>					
						<% Else %>
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= FormatCurrency(rsMCSActivtySummary("Month3Sales_NoRent"),0)%>&nbsp;</font>												
						<% End If %>	
					<% End If %>
				</td>
				<td width="5%" align="right">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= FormatCurrency(CustMonthlyContractedSalesDollars ,0)%>&nbsp;</font>
				</td>
				<td width="7%" align="right">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt" class="negative"><%= FormatCurrency(VarianceHolder,0)%>&nbsp;</font>
				</td>
				<%
				'EnrollmentDate Date
				EnrollmentDate = cDate(EnrollmentDate) 
				eYear = Year(EnrollmentDate)
				If Month(EnrollmentDate) < 10 Then eMonth = "0" & Month(EnrollmentDate) else eMonth = Month(EnrollmentDate)
				If Day(EnrollmentDate) < 10 Then eDay = "0" & Day(EnrollmentDate) else eDay = Day(EnrollmentDate)
				EnrollmentDispayableDate = eMonth & "/" & eDay  & "/" & Right(eYear,2)
				'EnrollmentDispayableDate  = cDate(EnrollmentDispayableDate) 
				%>
				<td width="7%">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= EnrollmentDispayableDate  %></font>
				</td>
				<td width="43%" style="white-space: nowrap;"> 
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font>
				</td> 
				<%

				RowCount = RowCount + 1
				
				%>
				</tr>
				<% 
				

			End If

			rsMCSActivtySummary.Movenext	

			If RowCount > LinesPerPage Then
				%></table><%
				Call Footer
				Call PageHeader
				Call SubHeaderNoActivity
			End If
				
				
			Loop

				
			'NoBreak = True

			Call Footer	
		End If


			'*************************************************
			'***END END END END END END END END END END END 
			'*** This section is the NEW DETAIL section of all the
			'*** MCS customers which have NOT been reviewed
			'****************************************************


			'***************************************************
			'*** This section is the ENDING ADD / REMOVE Summary
			'***************************************************
			
			Call PageHeader

			Call SubHeaderAddDelSummary
			
			LinesPerPage = 42
						
			FontSizeVar = 9

			SQL = "SELECT * FROM AR_Customer INNER JOIN BI_MCSActions ON BI_MCSActions.CustID = AR_Customer.CustNum "
			SQL = SQl & " WHERE Month(BI_MCSActions.RecordCreationDateTime) = Month(getdate()) "
			SQL = SQl & " AND Year(BI_MCSActions.RecordCreationDateTime) = Year(getdate()) "
			SQL = SQl & " AND (BI_MCSActions.Action = 'MCS Client Added' OR BI_MCSActions.Action = 'MCS Client Removed' "
			SQL = SQl & " OR BI_MCSActions.Action = 'remove_client' OR BI_MCSActions.Action = 'add_client')"
		'	SQL = SQl & " ORDER BY BI_MCSActions.Action"
			
			'Response.Write(SQL)
			
			Set cnnMCSActivity = Server.CreateObject("ADODB.Connection")
			cnnMCSActivity.open (MUV_READ("ClientCnnString"))
			Set rsMCSActivtySummary = Server.CreateObject("ADODB.Recordset")
			rsMCSActivtySummary.CursorLocation = 3 
			rsMCSActivtySummary.Open SQL, cnnMCSActivity
			

			If Not rsMCSActivtySummary.EOF Then
				
				Do While Not rsMCSActivtySummary.EOF

					FontSizeVar = 9
					LinesPerPage = 32
					
					%>
					<tr>
					<%
				
					SelectedCustomerID = rsMCSActivtySummary("CustNum")
					CustName = rsMCSActivtySummary("Name")
					TotalCustsReported = TotalCustsReported + 1
		
				%>
					<td width="7%">
						<%
							LastMCSActionDate = GetLastMCSActionDateByMonthByYearByCust(SelectedCustomerID , Month(Now()), Year(Now()))
							'LastMCSActionDate = GetLastMCSActionDateByMonthByYearByCust(SelectedCustomerID , 12, 2018)
							If LastMCSActionDate <> "" Then
								LastMCSActionDate = cDate(LastMCSActionDate ) 
								eYear = Year(LastMCSActionDate )
								If Month(LastMCSActionDate ) < 10 Then eMonth = "0" & Month(LastMCSActionDate ) else eMonth = Month(LastMCSActionDate )
								If Day(LastMCSActionDate ) < 10 Then eDay = "0" & Day(LastMCSActionDate ) else eDay = Day(LastMCSActionDate )
								LastMCSActionDispayableDate = eMonth & "/" & eDay  & "/" & Right(eYear,2)
								'LastMCSActionDispayableDate = cDate(LastMCSActionDispayableDate ) 
							Else
								LastMCSActionDispayableDate = "none"
							End If
						%>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= LastMCSActionDispayableDate %></font>
					</td>
				
				<td width="5%">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= SelectedCustomerID %></font>
				</td>
				<td width="24%">
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><% If Len(rsMCSActivtySummary("Name")) > 35 Then Response.Write(Left(rsMCSActivtySummary("Name"),35)) Else Response.Write(rsMCSActivtySummary("Name"))  %></font>
				</td>
				<td width="64%" style="white-space: nowrap;">
				<% If rsMCSActivtySummary("Action") = "MCS Client Added" Then %>
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">ADDED</font>
				<% End If %>	
				<% If rsMCSActivtySummary("Action") = "MCS Client Removed" or rsMCSActivtySummary("Action") = "remove_client" Then %>
					<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">REMOVED</font>
				<% End If %>	
				</td> 
				</tr>
				<% 

				RowCount = RowCount + 1


				rsMCSActivtySummary.Movenext	
	
				If RowCount > LinesPerPage Then
					%></table><%
					Call Footer
					Call PageHeader
					Call SubHeader
				End If
				
			Loop
			
		Else
			
			%>
			<td colspan="7">
				<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">No <%=GetTerm("customers")%> added or deleted this month</font>
			</td>
			<%
		
		End If
		
		
		NoBreak = True
		
		Call Footer	

		'***************************************************
		'*** This section is the ENDING ADD / REMOVE Summary
		'***************************************************
			
%>

</td>
</tr>
</table>

</body>
</html>


<%
Sub PageHeader

	RowCount = 0
	%>

	<table border="0" width="100%">
		<tr>
			<td width="50%"><img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png"></td>
			<td width="50%">
				<p align="center"><b><font face="Consolas" size="4">MCS Activity Summary</font></b></p>
				<p align="center"><font face="Consolas" size="2">Report Generated: <%= WeekDayName(WeekDay(DateValue(Now()))) %>&nbsp;<%= Now() %><br></font></p>
			</td>
		</tr>
		<tr>
			<td width="20%" height="16">
				<p align="right"><font face="Consolas" size="1">&nbsp;</font></p>
			</td>
		</tr>
	</table>
	<%
	PageNum = PageNum + 1
End Sub

Sub SubHeader
	%> 
	<br><br><br>
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<tr>
				<td>
				<font face="Consolas">
				<hr>
				<center><h4>Reviewed <%=GetTerm("Customers") %></h4></center>
				</font>
				</td>
			</tr>
		<tr>
			<td>
			<font face="Consolas">
			<hr>
			</font>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
	
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">
	<tr>
		<td width="7%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Activty<br>Date</font></u></strong>
		</td>
		<td width="5%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=GetTerm("Customer")%>#</font></u></strong>
		</td>
		<td width="24%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Name</font></u></strong>
		</td>
		<td width="7%" align="center">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Sales<br><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %>&nbsp;<%= Year(DateAdd("m",-1,ReportDate))%></font></u></strong>
		</td>
		<td width="5%" align="center">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">MCS</font></u></strong>
		</td>
		<td width="7%" align="center">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">MCS<br>Variance</font></u></strong>
		</td>
		<td width="7%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Enrollment<br>Date</font></u></strong>
		</td>
		<td width="43%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Last Action</font></u></strong>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<%
End Sub

Sub SubHeaderNoActivity
	%> 
	<br><br><br>
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<tr>
				<td>
				<font face="Consolas">
				<hr>
				<center><h4>Not Reviewed</h4></center>
				</font>
				</td>
			</tr>
		<tr>
			<td>
			<font face="Consolas">
			<hr>
			</font>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
	
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">
	<tr>
		<td width="5%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=GetTerm("Customer")%>#</font></u></strong>
		</td>
		<td width="31%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Name</font></u></strong>
		</td>
		<td width="7%" align="center">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Sales<br><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %>&nbsp;<%= Year(DateAdd("m",-1,ReportDate))%></font></u></strong>
		</td>
		<td width="5%" align="center">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">MCS</font></u></strong>
		</td>
		<td width="7%" align="center">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">MCS<br>Variance</font></u></strong>
		</td>
		<td width="7%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Enrollment<br>Date</font></u></strong>
		</td>
		<td width="43%">
			<strong><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></strong>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<%
End Sub

Sub SubHeaderAddDelSummary
	%> 
	<br><br><br>
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
		<tr>
			<td>
			<font face="Consolas">
			<hr>
			</font>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>

	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">
	<tr>
	<td colspan ="3">
		<font face="Consolas">
			<hr>
			<center><h3>ADDS / REMOVES <%= MonthName(Month(ReportDate)) %></h3></center>
			<hr>
		</font>
		</td>
	</tr>

	<tr>
		<td width="7%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Activty<br>Date</font></u></strong>
		</td>
		<td width="5%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%=GetTerm("Customer")%>#</font></u></strong>
		</td>
		<td width="24%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Name</font></u></strong>
		</td>
		<td width="64%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">ADD / REMOVE</font></u></strong>
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>
	<%
End Sub

Sub Footer

	'Now get us to the next page
	Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'><table>")
	For x = 1 to LinesPerPage - RowCount
		Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'><tr><td border='1'>&nbsp;</td></tr>")
	Next
	Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'></table>")
	%>
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
		<tr>
			<td colspan="3">
				<hr>
			</td>
		</tr>
		<tr>
			<td width="33%">
				<font face="Consolas" style="font-size: 9pt">directlaunch/bizintel/mcs_activity_summary.asp</font>
			</td>
			<td width="33%" align="center">
				<font face="Consolas" style="font-size: 12pt">Page:&nbsp;<%=PageNum%></font>
			</td>
			<td width="33%">
				<font face="Consolas" style="font-size: 12pt">&nbsp;</font>
			</td>
		</tr>
	</table>
	<% If NoBreak <> True Then %>
		<BR style="page-break-after: always">
	<% End If

End Sub


Sub RebuildMCSData ()


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
			LVFHolder = LVFHolder + TotalUNPostedLVFByCustByMonthByYear(rsMCSData("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))

			
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

End Sub

Sub CalcSummaryInformation


	TotalSalesAllMCSCustomers = 0
	TotalPendingLVF = 0
	
	SQL = "SELECT * FROM AR_Customer INNER JOIN BI_MCSData ON BI_MCSData.CustID = AR_Customer.CustNum WHERE MonthlyContractedSalesDollars <> 0" 
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	Set rs = cnn8.Execute(SQL)

	If Not rs.Eof Then
	
		Do While Not rs.EOF

			ShowThisRecord = True
				
			If ShowThisRecord <> False Then			
			
				PrimarySalesMan =  ""
				SecondarySalesMan =  ""
				SelectedCustomerID = rs("CustNum")
				CustName = rs("Name")

				PrimarySalesMan = rs("Salesman")
				SecondarySalesMan = rs("SecondarySalesman")
				CustMonthlyContractedSalesDollars = rs("MonthlyContractedSalesDollars")
					
				'Decide if this record meets the filter criteria
				If FilterSlsmn1 <> "" And FilterSlsmn1 <> "All" Then
					If CInt(FilterSlsmn1) <> Cint(rs("Salesman")) Then ShowThisRecord = False
				End If
				If FilterSlsmn2 <> "" And FilterSlsmn2 <> "All" Then
					If CInt(FilterSlsmn2) <> Cint(rs("SecondarySalesman")) Then ShowThisRecord = False
				End If
		
			End If
			

			Month3Sales_NoRent = rs("Month3Sales_NoRent") - rs("Month3Cat21Sales") 
				

			TotalSalesAllMCSCustomers = TotalSalesAllMCSCustomers + Month3Sales_NoRent
			TotalMCSClients = TotalMCSClients + 1
			TotalMCSCommitment = TotalMCSCommitment + rs("MonthlyContractedSalesDollars")
			
			TotalLVFLastMonth = TotalLVFLastMonth + rs("LVFHolder")
			
			If Month3Sales_NoRent >= rs("MonthlyContractedSalesDollars") Then
				TotalCustomersOver = TotalCustomersOver + 1
				TotalOverDollars = TotalOverDollars + (Month3Sales_NoRent - rs("MonthlyContractedSalesDollars"))
			End If
			

			
			' Calc under by the current month recovered the deficit
			If Month3Sales_NoRent < rs("MonthlyContractedSalesDollars") Then 
				Month3LVF = TotalPostedLVFByCustByMonthByYear(rs("CustID"),Month(DateAdd("m",-1,ReportDate)),Year(DateAdd("m",-1,ReportDate)))
				If Not IsNumeric(Month3LVF ) Then Month3LVF = 0
				If Month3LVF < 1 Then Month3LVF = 0
				M3Stemp = rs("Month3Sales_NoRent") - (rs("Month3Cat21Sales") + Month3LVF)
				
				If rs("CurrentHolder") >= rs("MonthlyContractedSalesDollars") +  ABS((M3Stemp - rs("MonthlyContractedSalesDollars"))) Then
					TotalCustomersUnderButRecovered = TotalCustomersUnderButRecovered + 1
					TotalUnderButRecoveredDeficitDollars = TotalUnderButRecoveredDeficitDollars + (rs("MonthlyContractedSalesDollars") - Month3Sales_NoRent)
				Else
					 If ABS(rs("Month3Sales_NoRent") - rs("MonthlyContractedSalesDollars")) < 100 Then ' Variance
						If rs("Month3Sales_NoRent") <> 0 Then
							VariancePercentHolder = 100 - ((rs("Month3Sales_NoRent")/rs("MonthlyContractedSalesDollars")) * 100)
						Else
							VariancePercentHolder = 100 
						End If
						VariancePercentHolder  = VariancePercentHolder  * -1
				End If
			End If
			If Month3Sales_NoRent > 0 Then
				TotalCustomersUnder = TotalCustomersUnder + 1
				TotalUnderDollars = TotalUnderDollars + (rs("MonthlyContractedSalesDollars") - Month3Sales_NoRent)
			End If
		End If

			
			
			If Month3Sales_NoRent <= 0 Then
				TotalCustomersZeroSales = TotalCustomersZeroSales + 1
				TotalZeroSalesCommitment = TotalZeroSalesCommitment + rs("MonthlyContractedSalesDollars")
			End If

			
			If ShowThisRecord <> False Then
			
				TotalMonth3Sales = TotalMonth3Sales + Month3Sales_NoRent
				TotalVariance = TotalVariance + Month3Sales_NoRent - rs("MonthlyContractedSalesDollars")
				If Month3Sales_NoRent = 0 Then TotalClientWithZeroSales = TotalClientWithZeroSales + 1
				
				   
		    End If
		    
		    TotalPendingLVF = TotalPendingLVF + rs("PendingLVF")

			rs.movenext
				
		Loop
		
End If

End Sub


Sub CalcSummaryInformation2

	TotalCustsMCSAdded = 0 : TotalCustsMCSRemoved = 0 : TotalNetMCSChange = 0 : TotalNetLVFChange = 0 : TotalNoAction = 0 : TotalCustsMCSAddedDollars = 0 : TotalCustsMCSRemovedDollars = 0
	TotalFollowup = 0 : TotalNumCustsInvoiced = 0 : TotalActedUpon = 0 : TotalNOTActedUpon = 0
	TotalNUMMCSChange = 0 : TotalNUMLVFChange = 0 : TotalMsgsSent = 0 : TotalLVFInvoicedAmount = 0 : TotalNotActedUponMCSDollars = 0
	
	SQL = "SELECT * FROM BI_MCSActions WHERE "
	SQL = SQL & "(MONTH(RecordCreationDateTime) = MONTH(DATEADD(m, 0, GETDATE()))) AND "
	SQL = SQL & "(YEAR(RecordCreationDateTime) = YEAR(DATEADD(m, 0, GETDATE()))) "
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	Set rs = cnn8.Execute(SQL)

	If Not rs.Eof Then
	
		Do While Not rs.EOF

			Select Case Ucase(rs("Action"))

				Case "CHANGE_LVF"

					TotalNUMLVFChange = TotalNUMLVFChange + 1
					
					' Parse out the Action Notes & Figure out the chnage
					' to the LVF amount
					
					' FROM Amount
					ActionNotes = rs("ActionNotes")
					
					xpos = InStr(ActionNotes,Chr(36))
					ActionNotes = Mid(ActionNotes,xpos,len(ActionNotes)-xpos)
					ActionNotes = Trim(ActionNotes)
					
					xpos = InStr(ActionNotes," ")
					From_Amt = Mid(ActionNotes ,1 , xpos-1)
					From_Amt = cdbl(Trim(From_Amt))
					
					' TO AMOUNT
					ActionNotes = rs("ActionNotes")
					
					xpos = InStrRev(ActionNotes,Chr(36))
					To_Amt = Right(ActionNotes,len(ActionNotes)-xpos)
					To_Amt = cdbl(Trim(To_Amt))
					
					Net_Change = To_Amt - From_Amt
					
					TotalNetLVFChange = TotalNetLVFChange + Net_Change
					
				Case "CHANGE_MCS"

					TotalNUMMCSChange = TotalNUMMCSChange + 1
					
					' Parse out the Action Notes & Figure out the chnage
					' to the MCS amount
					
					' FROM Amount
					ActionNotes = rs("ActionNotes")
					
					xpos = InStr(ActionNotes,Chr(36))
					ActionNotes = Mid(ActionNotes,xpos,len(ActionNotes)-xpos)
					ActionNotes = Trim(ActionNotes)
					
					xpos = InStr(ActionNotes," ")
					From_Amt = Mid(ActionNotes ,1 , xpos-1)
					From_Amt = cdbl(Trim(From_Amt))
					
					' TO AMOUNT
					ActionNotes = rs("ActionNotes")
					
					xpos = InStrRev(ActionNotes,Chr(36))
					To_Amt = Right(ActionNotes,len(ActionNotes)-xpos)
					To_Amt = cdbl(Trim(To_Amt))
					
					Net_Change = To_Amt - From_Amt
					
					TotalNetMCSChange = TotalNetMCSChange + Net_Change
					
				Case "MCS CLIENT ADDED"
				
					TotalCustsMCSAdded = TotalCustsMCSAdded + 1
					
					'Get the MCS Dollars for the addition
					
					ActionNotes = rs("ActionNotes")
					
					xpos = InStr(ActionNotes,"$")
					ActionNotes = Right(ActionNotes,Len(ActionNotes)-xpos)
					
					xpos = InStr(ActionNotes,".")
					ActionNotes = Left(ActionNotes,xpos-1)

					NewDollars = cdbl(Trim(ActionNotes))
					
					TotalCustsMCSAddedDollars = TotalCustsMCSAddedDollars + NewDollars

				Case "MCS CLIENT REMOVED"

					TotalCustsMCSRemoved = TotalCustsMCSRemoved + 1
					
					'Get the MCS Dollars for the removal
					
					ActionNotes = rs("ActionNotes")
					
					xpos = InStr(ActionNotes,"$")
					
					If xpos <> 0 Then ' account for before we had the $ in the note
					
						ActionNotes = Right(ActionNotes,Len(ActionNotes)-xpos)

						NewDollars = cdbl(Trim(ActionNotes))
	
						TotalCustsMCSRemovedDollars = TotalCustsMCSRemovedDollars + NewDollars 
					End If
					
				Case "NO_ACTION_NECESSARY"

					TotalNoAction = TotalNoAction + 1				
					
				Case "REMOVE_CLIENT"
				
					TotalCustsMCSRemoved = TotalCustsMCSRemoved + 1
									
				Case "NOTIFY_SELECTED_SALES_PERSON"
	
					TotalFollowup = TotalFollowup + 1

				Case "SEND_INVOICE"			

					TotalNumCustsInvoiced = TotalNumCustsInvoiced + 1				

					' Parse action notes to accumulate the
					' total LVF invoced amount

					' INV AMOUNT
					ActionNotes = rs("ActionNotes")
					
					xpos = InStrRev(ActionNotes,Chr(36))
					Inv_Amt = Right(ActionNotes,len(ActionNotes)-xpos)
					Inv_Amt = cdbl(Trim(Inv_Amt))
					
					TotalLVFInvoicedAmount = TotalLVFInvoicedAmount + Inv_Amt

				Case "SEND_MESSAGE_TO_SOMEONE"
				
					TotalMsgsSent = TotalMsgsSent + 1
					
			End Select

			rs.movenext
				
		Loop
		
	End If

	'Now a quick query to get the total number of
	'customers acted uopn
	SQL = "SELECT DISTINCT CustID FROM BI_MCSActions WHERE "
	SQL = SQL & "(MONTH(RecordCreationDateTime) = MONTH(DATEADD(m, 0, GETDATE()))) AND "
	SQL = SQL & "(YEAR(RecordCreationDateTime) = YEAR(DATEADD(m, 0, GETDATE()))) "
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	Set rs = cnn8.Execute(SQL)

	If Not rs.EOF Then
		Do While Not rs.EOF
			TotalActedUpon = TotalActedUpon + 1
			rs.MoveNext
		Loop
	End IF
	
	
	'Now get the total mcs $ for those accounts that have not been acted on
	SQL = "SELECT SUM(MonthlyContractedSalesDollars) AS MCSSum FROM AR_Customer  WHERE "
	SQL = SQL &  "MonthlyContractedSalesDollars IS NOT NULL AND CustNum NOT IN ("
	SQL = SQL &  "SELECT DISTINCT CustID FROM BI_MCSActions WHERE "
	SQL = SQL & "(MONTH(RecordCreationDateTime) = MONTH(DATEADD(m, 0, GETDATE()))) AND "
	SQL = SQL & "(YEAR(RecordCreationDateTime) = YEAR(DATEADD(m, 0, GETDATE()))) "
	SQL = SQL & ")"
	
	Set rs = cnn8.Execute(SQL)

	If Not rs.EOF Then
		Do While Not rs.EOF
			TotalNotActedUponMCSDollars = rs("MCSSum")
			rs.MoveNext
		Loop
	End IF
	
	Set rs = Nothing
	cnn8.Close
	Set cnn8 = Nothing

	
End Sub

%>