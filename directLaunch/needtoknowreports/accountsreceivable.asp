<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear
Dim PageNum, RowCount, FontSizeVar
FontSizeVar = 10
PageNum = 0
NoBreak = False

%>
<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<%

dummy=MUV_Write("ClientID","") 'Need this here

ClientKey = Request.QueryString("c")
Session("ClientCnnString") = ""
dummy = MUV_Write("ClientCnnString","")

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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Need to know accounts receivable report<%
	Response.End
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
Set cnn_Settings_NeedToKnow = Server.CreateObject("ADODB.Connection")
cnn_Settings_NeedToKnow.open (MUV_READ("ClientCnnString"))
Set rs_Settings_NeedToKnow = Server.CreateObject("ADODB.Recordset")
rs_Settings_NeedToKnow.CursorLocation = 3 
SQL_Settings_NeedToKnow = "SELECT * FROM Settings_NeedToKnow"
Set rs_Settings_NeedToKnow = cnn_Settings_NeedToKnow.Execute(SQL_Settings_NeedToKnow)
If not rs_Settings_NeedToKnow.EOF Then
	N2KARReportONOFF = rs_Settings_NeedToKnow("N2KARReportONOFF")
Else
	N2KARReportONOFF = 0
End If


Set rs_Settings_NeedToKnow = Nothing
cnn_Settings_NeedToKnow.Close
Set cnn_Settings_NeedToKnow = Nothing

If N2KARReportONOFF <> 1 Then
	%>MDS Insight: The accounts receivable need to know report is not turned on.
	<%
	Response.End
End IF

%>



<body bgcolor="#FFFFFF" text="#000000" link="#000080" vlink="#000080" alink="#000080" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">

<table border="0" width="800" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">
	<tr>
		<td width="100%" align="center">

		<% Call PageHeader
					
		'**************************************************************************'
		'*** First determine which accounts receivable N2K reports to include
		'**************************************************************************'
		
		LinesPerPage = 40

		SQLNeedToKnowInclude = "SELECT * FROM Settings_NeedToKnow"

		Set cnnNeedToKnowInclude = Server.CreateObject("ADODB.Connection")
		cnnNeedToKnowInclude.open (MUV_READ("ClientCnnString"))
		Set rsNeedToKnowInclude  = Server.CreateObject("ADODB.Recordset")
		rsNeedToKnowInclude.CursorLocation = 3 
		rsNeedToKnowInclude.Open SQLNeedToKnowInclude, cnnNeedToKnowInclude 
		
		If Not rsNeedToKnowInclude.EOF Then
		
			N2KARIncludeEmptyCustomerName = rsNeedToKnowInclude("N2KARIncludeEmptyCustomerName")
			N2KARIncludeEmptyAddress2 = rsNeedToKnowInclude("N2KARIncludeEmptyAddress2")
			N2KARIncludeEmptyPhoneNumber = rsNeedToKnowInclude("N2KARIncludeEmptyPhoneNumber")
			N2KARIncludeEmptyCity = rsNeedToKnowInclude("N2KARIncludeEmptyCity")
			N2KARIncludeEmptyState = rsNeedToKnowInclude("N2KARIncludeEmptyState")
			N2KARIncludeEmptyZip = rsNeedToKnowInclude("N2KARIncludeEmptyZip")
			N2KARIncludeEmptyCityStateZip = rsNeedToKnowInclude("N2KARIncludeEmptyCityStateZip")
			N2KARIncludeInvalidCityStateZip = rsNeedToKnowInclude("N2KARIncludeInvalidCityStateZip")
			N2KARIncludeInvalidPhoneNumber = rsNeedToKnowInclude("N2KARIncludeInvalidPhoneNumber")
			N2KARIncludeInvalidState = rsNeedToKnowInclude("N2KARIncludeInvalidState")
			N2KARIncludeInvalidZipCode = rsNeedToKnowInclude("N2KARIncludeInvalidZipCode")
			N2KARIncludeMissingcustomertype = rsNeedToKnowInclude("N2KARIncludeMissingcustomertype")
			N2KARIncludeMissingprimarysalesman = rsNeedToKnowInclude("N2KARIncludeMissingprimarysalesman")
			N2KARIncludeMissingsecondarysalesman = rsNeedToKnowInclude("N2KARIncludeMissingsecondarysalesman")
			N2KARIncludeNotAssignedToRegion	= rsNeedToKnowInclude("N2KARIncludeNotAssignedToRegion")
			
		Else
		
			N2KARIncludeEmptyCustomerName = 0
			N2KARIncludeEmptyAddress2 = 0
			N2KARIncludeEmptyPhoneNumber = 0
			N2KARIncludeEmptyCity = 0
			N2KARIncludeEmptyState = 0
			N2KARIncludeEmptyZip = 0
			N2KARIncludeEmptyCityStateZip = 0
			N2KARIncludeInvalidCityStateZip = 0
			N2KARIncludeInvalidPhoneNumber = 0
			N2KARIncludeInvalidState = 0
			N2KARIncludeInvalidZipCode = 0
			N2KARIncludeMissingcustomertype = 0
			N2KARIncludeMissingprimarysalesman = 0
			N2KARIncludeMissingsecondarysalesman = 0
			N2KARIncludeNotAssignedToRegion	= 0
		
		End If
		
		'**************************************************************************'
		'*** Outer loop to get the Summary Description Records
		'**************************************************************************'
		
		SQL = "SELECT SummaryDescription, Count(SummaryDescription) as ProblemCount FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND InsightStaffOnly <> 1"
		SQL = SQL & "Group By SummaryDescription ORDER BY SummaryDescription"

		Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
		cnnNeed2Know.open (MUV_READ("ClientCnnString"))
		Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
		rsN2KOuter.CursorLocation = 3 
		rsN2KOuter.Open SQL, cnnNeed2Know 
		
		
		'**************************************************************************'
		'*** Get the current active customer count
		'**************************************************************************'
		
		ActiveCustomerCountNotXStatus = NumberOfARCustAccountsNotX()
		
		If Not rsN2KOuter.EOF Then
			%>
			<br><br><br>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
				<tr>
					<td>
					<font face="Consolas">
					<hr>
					<center><h2>Summary</h2></center>
					<hr>
					</font>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td><font face="Consolas" style="font-size: 12pt">&nbsp;</font></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>

					
			<%
				If Not rsN2KOuter.EOF Then
				
					EmptyCustomerNameCount = 0
					EmptyAddress2Count = 0		
					EmptyCityCount = 0
					EmptyStateCount = 0
					EmptyZipCount = 0
					EmptyCityStateZipCount = 0
					EmptyPhoneNumberCount = 0
					InvalidStateCount = 0
					InvalidZipCodeCount = 0
					InvalidCityStateZipCount = 0
					InvalidPhoneNumberCount = 0
					MissingCustomerTypeCount = 0
					MissingPrimarySalesmanCount = 0
					MissingSecondarySalesmanCount = 0
					NoRegionForCustomerCount = 0
				
					Do While Not (rsN2KOuter.EOF)
			
						 SummaryDescription = rsN2KOuter("SummaryDescription")
						 
						 Select Case UCASE(SummaryDescription)
						 
						 	Case "EMPTY CUSTOMER NAME"
						 		EmptyCustomerNameCount = rsN2KOuter("ProblemCount")
						 	
						 	Case "EMPTY ADDRESS 2"
						 		EmptyAddress2Count = rsN2KOuter("ProblemCount")

						 	Case "EMPTY CITY"
						 		EmptyCityCount = rsN2KOuter("ProblemCount")

						 	Case "EMPTY STATE"
						 		EmptyStateCount = rsN2KOuter("ProblemCount")

						 	Case "EMPTY ZIP"
						 		EmptyZipCount = rsN2KOuter("ProblemCount")

						 	Case "EMPTY CITYSTATEZIP"
						 		EmptyCityStateZipCount = rsN2KOuter("ProblemCount")

						 	Case "EMPTY PHONE NUMBER"
						 		EmptyPhoneNumberCount = rsN2KOuter("ProblemCount")

						 	Case "INVALID STATE"
						 		InvalidStateCount = rsN2KOuter("ProblemCount")

						 	Case "INVALID ZIP CODE"
						 		InvalidZipCodeCount = rsN2KOuter("ProblemCount")

						 	Case "INVALID CITYSTATEZIP"
						 		InvalidCityStateZipCount = rsN2KOuter("ProblemCount")

						 	Case "INVALID PHONE NUMBER"
						 		InvalidPhoneNumberCount = rsN2KOuter("ProblemCount")

							Case "MISSING CUSTOMER TYPE"
								MissingCustomerTypeCount = rsN2KOuter("ProblemCount")
								
							Case "MISSING PRIMARY SALESMAN"
								MissingPrimarySalesmanCount = rsN2KOuter("ProblemCount")
								
							Case "MISSING SECONDARY SALESMAN"
								MissingSecondarySalesmanCount = rsN2KOuter("ProblemCount")
								
							Case "NO REGION FOR THIS CUSTOMER"
								NoRegionForCustomerCount = rsN2KOuter("ProblemCount")
								
						End Select
																								
										
					rsN2KOuter.MoveNext
					Loop
				
						
					If N2KARIncludeEmptyCustomerName = 1 Then
						IF EmptyCustomerNameCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= EmptyCustomerNameCount %>&nbsp;customers have an Empty Customer Name</font></td></tr><%
						End If
					End If
		
					If N2KARIncludeEmptyAddress2 = 1 Then
						IF EmptyAddress2Count > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= EmptyAddress2Count %>&nbsp;customers have an Empty Address 2</font></td></tr><%
						End If
					End If
				
					If N2KARIncludeEmptyCity = 1 Then
						If EmptyCityCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= EmptyCityCount %>&nbsp;customers have an Empty City</font></td></tr><%
						End If
					End If

					If N2KARIncludeEmptyState = 1 Then
						If EmptyStateCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= EmptyStateCount %>&nbsp;customers have an Empty State</font></td></tr><%
						End If
					End If

					If N2KARIncludeEmptyZip = 1 Then
						If EmptyZipCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= EmptyZipCount %>&nbsp;customers have an Empty Zip Code</font></td></tr><%
						End If
					End If

					If N2KARIncludeEmptyCityStateZip = 1 Then
						If EmptyCityStateZipCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= EmptyCityStateZipCount %>&nbsp;customers have an Empty CityStateZip</font></td></tr><%
						End If
					End If

					If N2KARIncludeEmptyPhoneNumber = 1 Then
						If EmptyPhoneNumberCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= EmptyPhoneNumberCount %>&nbsp;customers have an Empty Phone Number</font></td></tr><%
						End If
					End If

					If N2KARIncludeInvalidState = 1 Then
						If InvalidStateCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= InvalidStateCount %>&nbsp;customers have an Invalid State</font></td></tr><%
						End If
					End If

					If N2KARIncludeInvalidZipCode = 1 Then
						If InvalidZipCodeCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= InvalidZipCodeCount %>&nbsp;customers have an Invalid Zip Code</font></td></tr><%
						End If
					End If

					If N2KARIncludeInvalidCityStateZip = 1 Then
						If InvalidCityStateZipCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= InvalidCityStateZipCount %>&nbsp;customers have an Invalid CityStateZip</font></td></tr><%
						End If
					End If

					If N2KARIncludeInvalidPhoneNumber = 1 Then
						If InvalidPhoneNumberCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= InvalidPhoneNumberCount %>&nbsp;customers have an Invalid Phone Number</font></td></tr><%
						End If
					End If

					If N2KARIncludeMissingcustomertype = 1 Then
						If MissingCustomerTypeCount > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= MissingCustomerTypeCount %>&nbsp;customers have a Missing Customer Type</font></td></tr><%
						End If
					End If

					If N2KARIncludeMissingprimarysalesman = 1 Then
						If MissingPrimarySalesmanCount  > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= MissingPrimarySalesmanCount %>&nbsp;customers have a Missing Primary Salesman</font></td></tr><%
						End If
					End If

					If N2KARIncludeMissingsecondarysalesman = 1 Then
						If MissingSecondarySalesmanCount  > 0 Then
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= MissingSecondarySalesmanCount %>&nbsp;customers have a Missing Secondary Salesman</font></td></tr><%
						End If
					End If

					If N2KARIncludeNotAssignedToRegion = 1 Then
						If NoRegionForCustomerCount > 0 Then 
							%><tr><td><font face="Consolas" style="font-size: 14pt"><%= NoRegionForCustomerCount %>&nbsp;customers are Not Assigned to a Region</font></td></tr><%
						End If
					End If
							
					%>
					
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					

			<%	Else %>
			
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<hr>
							<center>Great news! There are no issues to report.</center></font>
							<hr>
							<% NoBreak = True %>
						</td>
					</tr>

			<% End If %>


		<%	Else %>
		
				<tr><td>&nbsp;</td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td>
						<font face="Consolas" style="font-size: 14pt">
						<hr>
						<center>Great news! There are no issues to report.</center></font>
						<hr>
						<% NoBreak = True %>
					</td>
				</tr>

		<% End If %>

	</table>
	</td>
</tr>
<tr>
<td>
	
<% 

Call Footer(LinesPerPage)

			'Now we start doing all the individual detail sections
			
			FontSizeVar = 10
			LinesPerPage = 45
			

			'*************************************************************************************************************
			' E M P T Y     C U S T O M E R     N A M E
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Empty Customer Name' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeEmptyCustomerName = 1 AND EmptyCustomerNameCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(EmptyCustomerNameCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Empty Customer Name")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an empty customer name.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((EmptyCustomerNameCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Empty Customer Name")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an empty customer name. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Empty Customer Name")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Empty Customer Name")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    E M P T Y     C U S T O M E R     N A M E
			'*************************************************************************************************************


			'*************************************************************************************************************
			' E M P T Y   A D D R E S S    2
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Empty Address 2' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeEmptyAddress2 = 1 AND EmptyAddress2Count > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(EmptyAddress2Count) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Empty Address 2")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an empty address 2.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((EmptyAddress2Count/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Empty Address 2")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an empty address 2. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Empty Address 2")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Empty Address 2")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D    E M P T Y   A D D R E S S    2
			'*************************************************************************************************************



			'*************************************************************************************************************
			' E M P T Y   C I T Y
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Empty City' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeEmptyCity = 1 AND EmptyCityCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(EmptyCityCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Empty City")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an empty city.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((EmptyCityCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Empty City")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an empty city. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Empty City")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Empty City")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D    E M P T Y   C I T Y
			'*************************************************************************************************************
			
			


			'*************************************************************************************************************
			' E M P T Y   S T A T E 
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Empty State' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeEmptyState = 1 AND EmptyStateCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(EmptyStateCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Empty State")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an empty state.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((EmptyStateCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Empty State")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an empty state. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Empty State")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Empty State")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D    E M P T Y    S T A T E 
			'*************************************************************************************************************
			



			'*************************************************************************************************************
			' E M P T Y   Z I P
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Empty Zip' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeEmptyZip = 1 AND EmptyZipCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(EmptyZipCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Empty State")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an empty zip code.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((EmptyZipCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Empty Zip")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an empty zip code. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Empty Zip")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Empty Zip")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D    E M P T Y    Z I P
			'*************************************************************************************************************
			
			

			'*************************************************************************************************************
			' E M P T Y   C I T Y S T A T E Z I P
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Empty CityStateZip' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeEmptyCityStateZip = 1 AND EmptyCityStateZipCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(EmptyCityStateZipCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Empty CityStateZip")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an empty zip code.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((EmptyCityStateZipCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Empty CityStateZip")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an empty CityStateZip. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Empty CityStateZip")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Empty CityStateZip")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D    E M P T Y     C I T Y S T A T E Z I P
			'*************************************************************************************************************
			

			

			'*************************************************************************************************************
			' E M P T Y   P H O N E    N U M B E R 
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Empty Phone Number' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeEmptyPhoneNumber = 1 AND EmptyPhoneNumberCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(EmptyPhoneNumberCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Empty Phone Number")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an empty phone number.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((EmptyPhoneNumberCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Empty Phone Number")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an empty phone number. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Empty Phone Number")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Empty Phone Number")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D    E M P T Y     P H O N E    N U M B E R 
			'*************************************************************************************************************



			

			'*************************************************************************************************************
			' I N V A L I D     S T A T E
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Invalid State' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeInvalidState = 1 AND InvalidStateCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(InvalidStateCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Invalid State")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an invalid state.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((InvalidStateCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Invalid State")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an invalid state. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Invalid State")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<%
									'If the accounts receivable description contains "invalid" data, parse out the invalid info from DetailedDescription1
									DetailedDescription1 = rsN2KOuter("DetailedDescription1")
									Set r = New RegExp
									r.Global = True
									r.Pattern = "\(([^)]+)\)"
									For Each m In r.Execute(DetailedDescription1)
										InvalidDesc = m.SubMatches(0)
									Next
									
									%>
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt">Invalid State <strong>(<%= InvalidDesc %>)</strong> for <%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Invalid State")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D    I N V A L I D     S T A T E
			'*************************************************************************************************************
			
			

			

			'*************************************************************************************************************
			' I N V A L I D     C I T Y 
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Invalid City' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeInvalidCity = 1 AND InvalidCityCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(InvalidCityCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Invalid City")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an invalid city.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((InvalidCityCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Invalid City")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an invalid city. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Invalid City")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<%
									'If the accounts receivable description contains "invalid" data, parse out the invalid info from DetailedDescription1
									DetailedDescription1 = rsN2KOuter("DetailedDescription1")
									Set r = New RegExp
									r.Global = True
									r.Pattern = "\(([^)]+)\)"
									For Each m In r.Execute(DetailedDescription1)
										InvalidDesc = m.SubMatches(0)
									Next
									
									%>
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt">Invalid City <strong>(<%= InvalidDesc %>)</strong> for <%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Invalid City")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D    I N V A L I D    C I T Y
			'*************************************************************************************************************
			
			
			
			

			'*************************************************************************************************************
			' I N V A L I D     C I T Y S T A T E Z I P
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Invalid CityStateZip' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeInvalidCityStateZip = 1 AND InvalidCityStateZipCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(InvalidCityStateZipCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Invalid CityStateZip")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an invalid CityStateZip.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((InvalidCityStateZipCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Invalid CityStateZip")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an invalid CityStateZip. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Invalid CityStateZip")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<%
									'If the accounts receivable description contains "invalid" data, parse out the invalid info from DetailedDescription1
									DetailedDescription1 = rsN2KOuter("DetailedDescription1")
									Set r = New RegExp
									r.Global = True
									r.Pattern = "\(([^)]+)\)"
									For Each m In r.Execute(DetailedDescription1)
										InvalidDesc = m.SubMatches(0)
									Next
									
									%>
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt">Invalid CityStateZip <strong>(<%= InvalidDesc %>)</strong> for <%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Invalid CityStateZip")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D    I N V A L I D     C I T Y S T A T E Z I P
			'*************************************************************************************************************


			

			'*************************************************************************************************************
			' I N V A L I D    P H O N E   N U M B E R
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Invalid Phone Number' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeInvalidPhoneNumber = 1 AND InvalidPhoneNumberCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(InvalidPhoneNumberCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Invalid Phone Number")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have an invalid phone number.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((InvalidPhoneNumberCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Invalid Phone Number")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have an invalid phone number. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Invalid Phone Number")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<%
									'If the accounts receivable description contains "invalid" data, parse out the invalid info from DetailedDescription1
									DetailedDescription1 = rsN2KOuter("DetailedDescription1")
									Set r = New RegExp
									r.Global = True
									r.Pattern = "\(([^)]+)\)"
									For Each m In r.Execute(DetailedDescription1)
										InvalidDesc = m.SubMatches(0)
									Next
									
									%>
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt">Invalid Phone Number <strong>(<%= InvalidDesc %>)</strong> for <%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Invalid Phone Number")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D    I N V A L I D    P H O N E   N U M B E R
			'*************************************************************************************************************



		

			'*************************************************************************************************************
			' M I S S I N G    C U S T O M E R   T Y P E
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Missing Customer Type' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeMissingCustomerType = 1 AND MissingCustomerTypeCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(MissingCustomerTypeCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Missing Customer Type")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have a missing Customer Type.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((MissingCustomerTypeCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Missing Customer Type")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have a missing Customer Type. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Missing Customer Type")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Missing Customer Type")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D   M I S S I N G  C U S T O M E R   T Y P E
			'*************************************************************************************************************
			
			

		

			'*************************************************************************************************************
			' M I S S I N G   P R I M A R Y    S A L E S M A N
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Missing Primary Salesman' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeMissingPrimarySalesman = 1 AND MissingPrimarySalesmanCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(MissingPrimarySalesmanCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Missing Primary Salesman")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have a missing Primary Salesman.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((MissingPrimarySalesmanCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Missing Primary Salesman")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have a missing Primary Salesman. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Missing Primary Salesman")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Missing Primary Salesman")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D   M I S S I N G    P R I M A R Y    S A L E S M A N
			'*************************************************************************************************************
			


			'*************************************************************************************************************
			' M I S S I N G    S E C O N D A R Y    S A L E S M A N
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'Missing Secondary Salesman' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeMissingSecondarySalesman = 1 AND MissingSecondarySalesmanCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(MissingSecondarySalesmanCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Missing Secondary Salesman")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers have a missing Secondary Salesman.</font>
						</td>
					</tr>
					
					</table>
					<%	 
					Call Footer(RowCount+5)
				
				'*******************************75% Threshold************************************	
				ElseIf (((MissingSecondarySalesmanCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Missing Secondary Salesman")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers have a missing Secondary Salesman. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Missing Secondary Salesman")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Missing Secondary Salesman")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D   M I S S I N G    S E C O N D A R Y    S A L E S M A N
			'*************************************************************************************************************




			'*************************************************************************************************************
			' N O T    A S S I G N E D   T O    R E G I O N
			'*************************************************************************************************************
			
			RowCount = 0
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Accounts Receivable' AND SummaryDescription = 'No region for this customer' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KARIncludeNotAssignedToRegion = 1 AND NoRegionForCustomerCount > 0 Then
								
				'*******************************ALL CUSTOMERS************************************
			 	If cInt(NoRegionForCustomerCount) = cInt(ActiveCustomerCountNotXStatus) Then
			 	'*******************************ALL CUSTOMERS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Not Assigned To Region")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All active customers are not assigned to a region.</font>
						</td>
					</tr>
					
					</table>
					<%	 
					Call Footer(RowCount)
				
				'*******************************75% Threshold************************************	
				ElseIf (((NoRegionForCustomerCount/ActiveCustomerCountNotXStatus)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Not Assigned To Region")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>							
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of active customers are not assigned to a region. Too many to list individually.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Not Assigned To Region")
				
					Do While Not rsN2KOuter.EOF
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("CustIDIfApplicable") %></font>
								</td>
								<td width="80%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetCustNameByCustNum(rsN2KOuter("CustIDIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount)
							Call PageHeader
							Call SubHeader("Not Assigned To Region")
							RowCount = 0
						End If
						
					Loop
					
					NoBreak = True
					Call Footer(RowCount)
	
				End If
	
			End If

			
			'*************************************************************************************************************
			' E N D   N O T    A S S I G N E D   T O    R E G I O N
			'*************************************************************************************************************
				
				


			 %>
			</table>


</td>
</tr>
</table>

</body>
</html>


<%
Sub PageHeader

	RowCount = 0
	%>

	<table border="0" width="800" align="center">
		<tr>
			<td width="50%"><img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png"></td>
			<td width="50%">
				<p align="center"><b><font face="Consolas" size="4">MDS Insight Accounts Receivable<br>Need To Know Report</font></b></p>
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

Sub SubHeader(HeaderText)
	%> 
	<br><br><br>
	<table border="0" width="100%" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
		<tr>
			<td>
			<font face="Consolas">
			<hr>
			<center><h2><%= HeaderText %></h2></center>
			<hr>
			</font>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
	
	<table border="0" width="100%" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">
	<tr>
		<td width="15%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Cust</font></u></strong>
		</td>

		<td width="80%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Details</font></u></strong>
		</td>

	</tr>
	<%
End Sub

Sub Footer(passedRowCount)

	'Now get us to the next page
	Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'><table>")
	For x = 1 to LinesPerPage - passedRowCount
		Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'><tr><td border='1'>&nbsp;</td></tr>")
	Next
	Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'></table>")
	%>
	<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
		<tr>
			<td colspan="3">
				<hr>
			</td>
		</tr>
		<tr>
			<td width="33%">
				<font face="Consolas" style="font-size: 9pt">directlaunch/needtoknowreports/accountsreceivable.asp</font>
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

%>