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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Global Settings Need To Know Report<%
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
	N2KGlobalSettingsReportONOFF = rs_Settings_NeedToKnow("N2KGlobalSettingsReportONOFF")
Else
	N2KGlobalSettingsReportONOFF = 0
End If
Set rs_Settings_NeedToKnow = Nothing
cnn_Settings_NeedToKnow.Close
Set cnn_Settings_NeedToKnow = Nothing

If N2KGlobalSettingsReportONOFF <> 1 Then
	%>MDS Insight: The global settings need to know report is not turned on.
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
		'*** First determine which Global Settings N2K reports to include
		'**************************************************************************'
		
		LinesPerPage = 40

		SQLNeedToKnowInclude = "SELECT * FROM Settings_NeedToKnow"

		Set cnnNeedToKnowInclude = Server.CreateObject("ADODB.Connection")
		cnnNeedToKnowInclude.open (MUV_READ("ClientCnnString"))
		Set rsNeedToKnowInclude  = Server.CreateObject("ADODB.Recordset")
		rsNeedToKnowInclude.CursorLocation = 3 
		rsNeedToKnowInclude.Open SQLNeedToKnowInclude, cnnNeedToKnowInclude 
		
		If Not rsNeedToKnowInclude.EOF Then
			N2KGlobalIncludeMissingClientLogoFile = rsNeedToKnowInclude("N2KGlobalIncludeMissingClientLogoFile")
			N2KGlobalIncludeMissingHolidayinCompanyCalendar = rsNeedToKnowInclude("N2KGlobalIncludeMissingHolidayinCompanyCalendar")
		Else
			N2KGlobalIncludeMissingClientLogoFile = 0
			N2KGlobalIncludeMissingHolidayinCompanyCalendar = 0
		End If
		
		'**************************************************************************'
		'*** Outer loop to get the Summary Description Records
		'**************************************************************************'
		
		SQL = "SELECT SummaryDescription, Count(SummaryDescription) as ProblemCount FROM SC_NeedToKnow WHERE Module = 'Global Settings' AND InsightStaffOnly <> 1"
		SQL = SQL & "Group By SummaryDescription ORDER BY SummaryDescription"

		Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
		cnnNeed2Know.open (MUV_READ("ClientCnnString"))
		Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
		rsN2KOuter.CursorLocation = 3 
		rsN2KOuter.Open SQL, cnnNeed2Know 
		 
		
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
				
					CountN2KGlobalIncludeMissingClientLogoFile = 0
					CountN2KGlobalIncludeMissingHolidayinCompanyCalendar = 0
				
					Do While Not (rsN2KOuter.EOF)
			
						 SummaryDescription = rsN2KOuter("SummaryDescription")
						 
						 Select Case UCASE(SummaryDescription)

						 	Case "MISSING CLIENT LOGO FILE"
						 		CountN2KGlobalIncludeMissingClientLogoFile = rsN2KOuter("ProblemCount")
						 		
						 	Case "MISSING HOLIDAY IN COMPANY CALENDAR"
						 		CountN2KGlobalIncludeMissingHolidayinCompanyCalendar = rsN2KOuter("ProblemCount")
								
						End Select
																								
										
					rsN2KOuter.MoveNext
					Loop
					
						
					If N2KGlobalIncludeMissingClientLogoFile = 1 AND CountN2KGlobalIncludeMissingClientLogoFile > 0 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt">*YOU HAVE NOT UPLOADED A CLIENT LOGO FILE. PLEASE UPLOAD ONE TODAY.*</font></td></tr><%
					End If
		
					If N2KGlobalIncludeMissingHolidayinCompanyCalendar = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KGlobalIncludeMissingHolidayinCompanyCalendar %>&nbsp;Holidays are Missing from the Company Calendar.</font></td></tr><%
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

Call Footer(5,LinesPerPage)

			'Now we start doing all the individual detail sections
			

			'*************************************************************************************************************
			' M I S S I N G    H O L I D A Y    I N     C O M P A N Y    C A L E N D A R
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 49
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Global Settings' AND SummaryDescription = 'Missing Holiday in Company Calendar' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KGlobalIncludeMissingHolidayinCompanyCalendar = 1 AND CountN2KGlobalIncludeMissingHolidayinCompanyCalendar > 0 Then
			
				Call PageHeader
				Call SubHeader("Missing Holiday in Company Calendar")
			
				Do While Not rsN2KOuter.EOF	
					%>
						<tr>						
							<td width="100%"><!-- Asset Tag 1-->
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
							<% RowCount = RowCount + 1 %>
						</tr>
						<tr><td>&nbsp;</td></tr>
					<%
					
					RowCount = RowCount + 2
					rsN2KOuter.Movenext	
	
					If RowCount > LinesPerPage Then
						%></table><%
						Call Footer(RowCount,LinesPerPage)
						Call PageHeader
						Call SubHeader("Missing Holiday in Company Calendar")
						RowCount = 0
					End If
					
				Loop
				Call Footer(RowCount,LinesPerPage)
	
			End If
			
			'*************************************************************************************************************
			' E N D    M I S S I N G    H O L I D A Y    I N     C O M P A N Y    C A L E N D A R
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
	<table border="0" width="100%">
		<tr>
			<td width="50%"><img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png"></td>
			<td width="50%">
				<p align="center"><b><font face="Consolas" size="4">MDS Insight Global Settings<br>Need To Know Report</font></b></p>
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
	<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
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
	
	<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">
		<tr>
			<td width="100%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Message Details</font></u></strong>
			</td>
		</tr>
		<tr><td width="100%">&nbsp;</td></tr>	
	<%
End Sub

Sub Footer(passedRowCount,passedLinesPerPage)

	'Now get us to the next page
	Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'><table>")
	For x = 1 to passedLinesPerPage - passedRowCount
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
				<font face="Consolas" style="font-size: 9pt">directlaunch/needtoknowreports/globalsettings.asp</font>
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