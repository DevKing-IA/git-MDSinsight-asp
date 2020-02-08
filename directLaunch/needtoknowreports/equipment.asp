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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Equipment Need To Know Report<%
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
	N2KEquipmentReportONOFF = rs_Settings_NeedToKnow("N2KEquipmentReportONOFF")
Else
	N2KEquipmentReportONOFF = 0
End If
Set rs_Settings_NeedToKnow = Nothing
cnn_Settings_NeedToKnow.Close
Set cnn_Settings_NeedToKnow = Nothing

If N2KEquipmentReportONOFF <> 1 Then
	%>MDS Insight: The equipment need to know report is not turned on.
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
		'*** First determine which Equipment N2K reports to include
		'**************************************************************************'
		
		LinesPerPage = 40

		SQLNeedToKnowInclude = "SELECT * FROM Settings_NeedToKnow"

		Set cnnNeedToKnowInclude = Server.CreateObject("ADODB.Connection")
		cnnNeedToKnowInclude.open (MUV_READ("ClientCnnString"))
		Set rsNeedToKnowInclude  = Server.CreateObject("ADODB.Recordset")
		rsNeedToKnowInclude.CursorLocation = 3 
		rsNeedToKnowInclude.Open SQLNeedToKnowInclude, cnnNeedToKnowInclude 
		
		If Not rsNeedToKnowInclude.EOF Then
			N2KEqpIncludeBlankInsightAssetTagBrandPrefix = rsNeedToKnowInclude("N2KEqpIncludeBlankInsightAssetTagBrandPrefix")
			N2KEqpIncludeBlankInsightAssetTagClassPrefix = rsNeedToKnowInclude("N2KEqpIncludeBlankInsightAssetTagClassPrefix")
			N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix = rsNeedToKnowInclude("N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix")
			N2KEqpIncludeBlankInsightAssetTagModelPrefix = rsNeedToKnowInclude("N2KEqpIncludeBlankInsightAssetTagModelPrefix")
			N2KEqpIncludeUndefinedBrandExistsforEqp = rsNeedToKnowInclude("N2KEqpIncludeUndefinedBrandExistsforEqp")
			N2KEqpIncludeUndefinedClassExistsforEqp = rsNeedToKnowInclude("N2KEqpIncludeUndefinedClassExistsforEqp")
			N2KEqpIncludeUndefinedConditionCodeExistsforEqp = rsNeedToKnowInclude("N2KEqpIncludeUndefinedConditionCodeExistsforEqp")
			N2KEqpIncludeUndefinedGroupExistsforEqp = rsNeedToKnowInclude("N2KEqpIncludeUndefinedGroupExistsforEqp")
			N2KEqpIncludeUndefinedManufacturerExistsforEqp = rsNeedToKnowInclude("N2KEqpIncludeUndefinedManufacturerExistsforEqp")
			N2KEqpIncludeUndefinedModelExistsforEqp = rsNeedToKnowInclude("N2KEqpIncludeUndefinedModelExistsforEqp")
			N2KEqpIncludeUndefinedStatusCodeExistsforEqp = rsNeedToKnowInclude("N2KEqpIncludeUndefinedStatusCodeExistsforEqp")
			N2KEqpIncludeZeroDollarRentalsExistforEqp = rsNeedToKnowInclude("N2KEqpIncludeZeroDollarRentalsExistforEqp")
		Else
			N2KEqpIncludeBlankInsightAssetTagBrandPrefix = 0
			N2KEqpIncludeBlankInsightAssetTagClassPrefix = 0
			N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix = 0
			N2KEqpIncludeBlankInsightAssetTagModelPrefix = 0
			N2KEqpIncludeUndefinedBrandExistsforEqp = 0
			N2KEqpIncludeUndefinedClassExistsforEqp = 0
			N2KEqpIncludeUndefinedConditionCodeExistsforEqp = 0
			N2KEqpIncludeUndefinedGroupExistsforEqp = 0
			N2KEqpIncludeUndefinedManufacturerExistsforEqp = 0
			N2KEqpIncludeUndefinedModelExistsforEqp = 0
			N2KEqpIncludeUndefinedStatusCodeExistsforEqp = 0
			N2KEqpIncludeZeroDollarRentalsExistforEqp = 0
		End If
		
		'**************************************************************************'
		'*** Outer loop to get the Summary Description Records
		'**************************************************************************'
		
		SQL = "SELECT SummaryDescription, Count(SummaryDescription) as ProblemCount FROM SC_NeedToKnow WHERE Module = 'Equipment' AND InsightStaffOnly <> 1"
		SQL = SQL & "Group By SummaryDescription ORDER BY SummaryDescription"

		Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
		cnnNeed2Know.open (MUV_READ("ClientCnnString"))
		Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
		rsN2KOuter.CursorLocation = 3 
		rsN2KOuter.Open SQL, cnnNeed2Know 
		 
		
		'**************************************************************************'
		'*** Get the current EQ_Equipment Record count
		'**************************************************************************'
		
		EquipmentPieceCount = GetTotalNumberOfCustomerEquipmentRecords()
		
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
					<td><font face="Consolas" style="font-size: 12pt">*Note:* This report includes all equipment records.</font></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>

					
			<%
				If Not rsN2KOuter.EOF Then
				
					CountN2KEqpIncludeBlankInsightAssetTagBrandPrefix = 0
					CountN2KEqpIncludeBlankInsightAssetTagClassPrefix = 0
					CountN2KEqpIncludeBlankInsightAssetTagManufacturerPrefix = 0
					CountN2KEqpIncludeBlankInsightAssetTagModelPrefix = 0
					CountN2KEqpIncludeUndefinedBrandExistsforEqp = 0
					CountN2KEqpIncludeUndefinedClassExistsforEqp = 0
					CountN2KEqpIncludeUndefinedConditionCodeExistsforEqp = 0
					CountN2KEqpIncludeUndefinedGroupExistsforEqp = 0
					CountN2KEqpIncludeUndefinedManufacturerExistsforEqp = 0
					CountN2KEqpIncludeUndefinedModelExistsforEqp = 0
					CountN2KEqpIncludeUndefinedStatusCodeExistsforEqp = 0
					CountN2KEqpIncludeZeroDollarRentalsExistforEqp = 0
				
					Do While Not (rsN2KOuter.EOF)
			
						 SummaryDescription = rsN2KOuter("SummaryDescription")
						 
						 Select Case UCASE(SummaryDescription)

						 	Case "BLANK INSIGHT ASSET TAG BRAND PREFIX"
						 		CountN2KEqpIncludeBlankInsightAssetTagBrandPrefix = rsN2KOuter("ProblemCount")
						 	
						 	Case "BLANK INSIGHT ASSET TAG CLASS PREFIX"
						 		CountN2KEqpIncludeBlankInsightAssetTagClassPrefix = rsN2KOuter("ProblemCount")
						 	
						 	Case "BLANK INSIGHT ASSET TAG MANUFACTURER PREFIX"
						 		CountN2KEqpIncludeBlankInsightAssetTagManufacturerPrefix = rsN2KOuter("ProblemCount")
						 	
						 	Case "BLANK INSIGHT ASSET TAG MODEL PREFIX"
						 		CountN2KEqpIncludeBlankInsightAssetTagModelPrefix = rsN2KOuter("ProblemCount")
						 	
						 	Case "UNDEFINED BRAND EXISTS FOR EQUIPMENT"
						 		CountN2KEqpIncludeUndefinedBrandExistsforEqp = rsN2KOuter("ProblemCount")
						 	
						 	Case "UNDEFINED CLASS EXISTS FOR EQUIPMENT"
						 		CountN2KEqpIncludeUndefinedClassExistsforEqp = rsN2KOuter("ProblemCount")
						 	
						 	Case "UNDEFINED CONDITION CODE EXISTS FOR EQUIPMENT"
						 		CountN2KEqpIncludeUndefinedConditionCodeExistsforEqp = rsN2KOuter("ProblemCount")
						 	
						 	Case "UNDEFINED GROUP EXISTS FOR EQUIPMENT"
						 		CountN2KEqpIncludeUndefinedGroupExistsforEqp = rsN2KOuter("ProblemCount")
						 	
						 	Case "UNDEFINED MANUFACTURER EXISTS FOR EQUIPMENT"
						 		CountN2KEqpIncludeUndefinedManufacturerExistsforEqp = rsN2KOuter("ProblemCount")
						 	
						 	Case "UNDEFINED MODEL EXISTS FOR EQUIPMENT"
						 		CountN2KEqpIncludeUndefinedModelExistsforEqp = rsN2KOuter("ProblemCount")
						 	
						 	Case "UNDEFINED STATUS CODE EXISTS FOR EQUIPMENT"
						 		CountN2KEqpIncludeUndefinedStatusCodeExistsforEqp = rsN2KOuter("ProblemCount")
						 	
						 	Case "ZERO DOLLAR RENTALS EXIST FOR EQUIPMENT"
						 		CountN2KEqpIncludeZeroDollarRentalsExistforEqp = rsN2KOuter("ProblemCount")
										
						End Select
																								
										
					rsN2KOuter.MoveNext
					Loop
					
						
					If N2KEqpIncludeBlankInsightAssetTagBrandPrefix = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeBlankInsightAssetTagBrandPrefix %>&nbsp;Insight Asset Tag Brand Prefixes are Blank</font></td></tr><%
					End If
		
					If N2KEqpIncludeBlankInsightAssetTagClassPrefix = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeBlankInsightAssetTagClassPrefix %>&nbsp;Insight Asset Tag Class Prefixes are Blank</font></td></tr><%
					End If
				
					If N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeBlankInsightAssetTagManufacturerPrefix %>&nbsp;Insight Asset Tag Manufacturer Prefixes are Blank</font></td></tr><%
					End If
							
					If N2KEqpIncludeBlankInsightAssetTagModelPrefix = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeBlankInsightAssetTagModelPrefix %>&nbsp;Insight Asset Tag Model Prefixes are Blank</font></td></tr><%
					End If

					If N2KEqpIncludeUndefinedBrandExistsforEqp = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeUndefinedBrandExistsforEqp %>&nbsp;Brands are Undefined</font></td></tr><%
					End If

					If N2KEqpIncludeUndefinedClassExistsforEqp = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeUndefinedClassExistsforEqp %>&nbsp;Classes are Undefined</font></td></tr><%
					End If

					If N2KEqpIncludeUndefinedConditionCodeExistsforEqp = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeUndefinedConditionCodeExistsforEqp %>&nbsp;Condition Codes are Undefined</font></td></tr><%
					End If

					If N2KEqpIncludeUndefinedGroupExistsforEqp = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeUndefinedGroupExistsforEqp %>&nbsp;Groups are Undefined </font></td></tr><%
					End If

					If N2KEqpIncludeUndefinedManufacturerExistsforEqp = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeUndefinedManufacturerExistsforEqp %>&nbsp;Manufacturers are Undefined</font></td></tr><%
					End If

					If N2KEqpIncludeUndefinedModelExistsforEqp = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeUndefinedModelExistsforEqp %>&nbsp;Models are Undefined </font></td></tr><%
					End If

					If N2KEqpIncludeUndefinedStatusCodeExistsforEqp = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeUndefinedStatusCodeExistsforEqp %>&nbsp;Statuses are Undefined</font></td></tr><%
					End If

					If N2KEqpIncludeZeroDollarRentalsExistforEqp = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KEqpIncludeZeroDollarRentalsExistforEqp %>&nbsp;pieces of customer equipment have a Zero Dollar Rentals</font></td></tr><%
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

Call Footer(15,LinesPerPage)

			'Now we start doing all the individual detail sections
			

			'*************************************************************************************************************
			' B L A N K   I N S I G H T    A S S E T    T A G    B R A N D    P R E F I X
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Blank Insight Asset Tag Brand Prefix' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeBlankInsightAssetTagBrandPrefix = 1 AND CountN2KEqpIncludeBlankInsightAssetTagBrandPrefix > 0 Then
			
				BrandCount = GetTotalNumberOfBrands()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagBrandPrefix) = cInt(BrandCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Brand Prefix")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Insight Asset Tag Brand Prefixes are Blank.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagBrandPrefix/BrandCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Brand Prefix")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Insight Asset Tag Brand Prefixes are Blank.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Brand Prefix")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Blank Insight Asset Tag Brand Prefix")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    B L A N K   I N S I G H T    A S S E T    T A G    B R A N D    P R E F I X
			'*************************************************************************************************************




			'*************************************************************************************************************
			' B L A N K   I N S I G H T    A S S E T    T A G    C L A S S    P R E F I X
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Blank Insight Asset Tag Class Prefix' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeBlankInsightAssetTagClassPrefix = 1 AND CountN2KEqpIncludeBlankInsightAssetTagClassPrefix > 0 Then
			
				ClassCount = GetTotalNumberOfClasses()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagClassPrefix) = cInt(ClassCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Class Prefix")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Insight Asset Tag Class Prefixes are Blank.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagClassPrefix/ClassCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Class Prefix")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Insight Asset Tag Class Prefixes are Blank.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Class Prefix")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Blank Insight Asset Tag Class Prefix")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    B L A N K   I N S I G H T    A S S E T    T A G    C L A S S    P R E F I X
			'*************************************************************************************************************




			'*************************************************************************************************************
			' B L A N K   I N S I G H T    A S S E T    T A G    M A N U F A C T U R E R    P R E F I X
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Blank Insight Asset Tag Manufacturer Prefix' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeBlankInsightAssetTagManufacturerPrefix = 1 AND CountN2KEqpIncludeBlankInsightAssetTagManufacturerPrefix > 0 Then
			
				ManufacturerCount = GetTotalNumberOfManufacturers()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagManufacturerPrefix) = cInt(ManufacturerCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Manufacturer Prefix")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Insight Asset Tag Manufacturer Prefixes are Blank.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagManufacturerPrefix/ManufacturerCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Manufacturer Prefix")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Insight Asset Tag Manufacturer Prefixes are Blank.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Manufacturer Prefix")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Blank Insight Asset Tag Manufacturer Prefix")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    B L A N K   I N S I G H T    A S S E T    T A G    M A N U F A C T U R E R     P R E F I X
			'*************************************************************************************************************





			'*************************************************************************************************************
			' B L A N K   I N S I G H T    A S S E T    T A G    M O D E L    P R E F I X
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Blank Insight Asset Tag Model Prefix' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeBlankInsightAssetTagModelPrefix = 1 AND CountN2KEqpIncludeBlankInsightAssetTagModelPrefix > 0 Then
			
				ModelCount = GetTotalNumberOfModels()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagModelPrefix) = cInt(ModelCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Model Prefix")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Insight Asset Tag Model Prefixes are Blank.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagModelPrefix/ModelCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Model Prefix")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Insight Asset Tag Model Prefixes are Blank.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Blank Insight Asset Tag Model Prefix")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Blank Insight Asset Tag Model Prefix")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    B L A N K   I N S I G H T    A S S E T    T A G    M O D E L     P R E F I X
			'*************************************************************************************************************



			'*************************************************************************************************************
			' U N D E F I N E D    B R A N D
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Undefined Brand Exists for Equipment' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeUndefinedBrandExistsforEqp = 1 AND CountN2KEqpIncludeBlankInsightAssetTagBrandPrefix > 0 Then
			
				BrandCount = GetTotalNumberOfBrands()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagBrandPrefix) = cInt(BrandCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Undefined Brand Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Brands are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagBrandPrefix/BrandCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Undefined Brand Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Brands are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Undefined Brand Exists for Equipment")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Undefined Brand Exists for Equipment")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    U N D E F I N E D    B R A N D
			'*************************************************************************************************************


			'*************************************************************************************************************
			' U N D E F I N E D    C L A S S
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Undefined Class Exists for Equipment' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeUndefinedClassExistsforEqp = 1 AND CountN2KEqpIncludeBlankInsightAssetTagClassPrefix > 0 Then
			
				ClassCount = GetTotalNumberOfClasses()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagClassPrefix) = cInt(ClassCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Undefined Class Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Classes are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagClassPrefix/ClassCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Undefined Class Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Classes are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Undefined Class Exists for Equipment")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Undefined Class Exists for Equipment")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    U N D E F I N E D    C L A S S
			'*************************************************************************************************************



			'*************************************************************************************************************
			' U N D E F I N E D    C O N D I T I O N    C O D E 
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Undefined Condition Code Exists for Equipment' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeUndefinedConditionCodeExistsforEqp = 1 AND CountN2KEqpIncludeBlankInsightAssetTagConditionCodePrefix > 0 Then
			
				ConditionCodeCount = GetTotalNumberOfConditionCodes()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagConditionCodePrefix) = cInt(ConditionCodeCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Undefined Condition Code Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Condition Codes are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagConditionCodePrefix/ConditionCodeCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Undefined Condition Code Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Condition Codes are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Undefined Condition Code Exists for Equipment")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Undefined Condition Code Exists for Equipment")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    C O N D I T I O N    C O D E 
			'*************************************************************************************************************


			'*************************************************************************************************************
			' U N D E F I N E D    G R O U P
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Undefined Group Exists for Equipment' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeUndefinedGroupExistsforEqp = 1 AND CountN2KEqpIncludeBlankInsightAssetTagGroupPrefix > 0 Then
			
				GroupCount = GetTotalNumberOfGroups()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagGroupPrefix) = cInt(GroupCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Undefined Group Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Groups are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagGroupPrefix/GroupCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Undefined Group Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Groups are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Undefined Group Exists for Equipment")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Undefined Group Exists for Equipment")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    U N D E F I N E D    G R O U P
			'*************************************************************************************************************



			'*************************************************************************************************************
			' U N D E F I N E D    M A N U F A C T U R E R
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Undefined Manufacturer Exists for Equipment' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeUndefinedManufacturerExistsforEqp = 1 AND CountN2KEqpIncludeBlankInsightAssetTagManufacturerPrefix > 0 Then
			
				ManufacturerCount = GetTotalNumberOfManufacturers()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagManufacturerPrefix) = cInt(ManufacturerCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Undefined Manufacturer Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Manufacturers are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagManufacturerPrefix/ManufacturerCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Undefined Manufacturer Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Manufacturers are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Undefined Manufacturer Exists for Equipment")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Undefined Manufacturer Exists for Equipment")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    U N D E F I N E D    M A N U F A C T U R E R
			'*************************************************************************************************************


			'*************************************************************************************************************
			' U N D E F I N E D    M O D E L
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Undefined Model Exists for Equipment' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeUndefinedModelExistsforEqp = 1 AND CountN2KEqpIncludeBlankInsightAssetTagModelPrefix > 0 Then
			
				ModelCount = GetTotalNumberOfModels()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagModelPrefix) = cInt(ModelCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Undefined Model Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Models are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagModelPrefix/ModelCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Undefined Model Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Models are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Undefined Model Exists for Equipment")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Undefined Model Exists for Equipment")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    U N D E F I N E D     M O D E L
			'*************************************************************************************************************


			'*************************************************************************************************************
			' U N D E F I N E D    M O D E L
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Undefined Status Code Exists for Equipment' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeUndefinedStatusCodeExistsforEqp = 1 AND CountN2KEqpIncludeBlankInsightAssetTagStatusCodePrefix > 0 Then
			
				StatusCodeCount = GetTotalNumberOfStatusCodes()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeBlankInsightAssetTagStatusCodePrefix) = cInt(StatusCodeCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Undefined Status Code Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Status Codes are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagStatusCodePrefix/StatusCodeCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Undefined Status Code Exists for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Status Codes are Undefined.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Undefined Status Code Exists for Equipment")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Undefined Status Code Exists for Equipment")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    U N D E F I N E D     M O D E L
			'*************************************************************************************************************



			'*************************************************************************************************************
			' Z E R O    D O L L A R    R E N T A L S    E X I S T    F O R    E Q U I P M E N T
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 45
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Equipment' AND SummaryDescription = 'Zero Dollar Rentals Exist for Equipment' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KEqpIncludeZeroDollarRentalsExistforEqp = 1 AND CountN2KEqpIncludeZeroDollarRentalsExistforEqp > 0 Then
			
				CustomerEquipmentCount = GetTotalNumberOfCustomerEquipmentRecords()
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KEqpIncludeZeroDollarRentalsExistforEqp) = cInt(CustomerEquipmentCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Zero Dollar Rentals Exist for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All Customer Equipment Has Zero Dollar Rentals.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KEqpIncludeBlankInsightAssetTagStatusCodePrefix/CustomerEquipmentCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Zero Dollar Rentals Exist for Equipment")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater Than 75% of Customer Equipment Has Zero Dollar Rentals.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Zero Dollar Rentals Exist for Equipment")
				
					Do While Not rsN2KOuter.EOF	
						%>
							<tr>						
								<td width="100%"><!-- Asset Tag 1-->
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("DetailedDescription1") %></font>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Zero Dollar Rentals Exist for Equipment")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    Z E R O    D O L L A R    R E N T A L S    E X I S T    F O R    E Q U I P M E N T
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
				<p align="center"><b><font face="Consolas" size="4">MDS Insight Equipment<br>Need To Know Report</font></b></p>
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
	
	<% If HeaderText = "Zero Dollar Rentals Exist For Equipment" Then %>
		<tr>
			<td width="100%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Message Details</font></u></strong>
			</td>
		</tr>
	<% Else %>
		<tr>
			<td width="100%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Message Details</font></u></strong>
			</td>
		</tr>
	<% End If %>
	
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
				<font face="Consolas" style="font-size: 9pt">directlaunch/needtoknowreports/equipment.asp</font>
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