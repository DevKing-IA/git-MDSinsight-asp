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
<!--#include file="../../inc/InsightFuncs_InventoryControl.asp"-->

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
	N2KInventoryReportONOFF = rs_Settings_NeedToKnow("N2KInventoryReportONOFF")
Else
	N2KInventoryReportONOFF = 0
End If
Set rs_Settings_NeedToKnow = Nothing
cnn_Settings_NeedToKnow.Close
Set cnn_Settings_NeedToKnow = Nothing

If N2KInventoryReportONOFF <> 1 Then
	%>MDS Insight: The inventory need to know report is not turned on.
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
		'*** First determine which inventory N2K reports to include
		'**************************************************************************'
		
		LinesPerPage = 40

		SQLNeedToKnowInclude = "SELECT * FROM Settings_NeedToKnow"

		Set cnnNeedToKnowInclude = Server.CreateObject("ADODB.Connection")
		cnnNeedToKnowInclude.open (MUV_READ("ClientCnnString"))
		Set rsNeedToKnowInclude  = Server.CreateObject("ADODB.Recordset")
		rsNeedToKnowInclude.CursorLocation = 3 
		rsNeedToKnowInclude.Open SQLNeedToKnowInclude, cnnNeedToKnowInclude 
		
		If Not rsNeedToKnowInclude.EOF Then
			N2KInventoryIncludeBlankCaseBin = rsNeedToKnowInclude("N2KInventoryIncludeBlankCaseBin")
			N2KInventoryIncludeBlankCaseUPCCode = rsNeedToKnowInclude("N2KInventoryIncludeBlankCaseUPCCode")
			N2KInventoryIncludeBlankUnitandCaseUPCCode = rsNeedToKnowInclude("N2KInventoryIncludeBlankUnitandCaseUPCCode")
			N2KInventoryIncludeBlankUnitBin = rsNeedToKnowInclude("N2KInventoryIncludeBlankUnitBin")
			N2KInventoryIncludeBlankUnitUPCCode = rsNeedToKnowInclude("N2KInventoryIncludeBlankUnitUPCCode")
			N2KInventoryIncludeDuplicateUnitorCaseBin = rsNeedToKnowInclude("N2KInventoryIncludeDuplicateUnitorCaseBin")
			N2KInventoryIncludeDuplicateUPCCode = rsNeedToKnowInclude("N2KInventoryIncludeDuplicateUPCCode")		
		Else
			N2KInventoryIncludeBlankCaseBin = 0
			N2KInventoryIncludeBlankCaseUPCCode = 0
			N2KInventoryIncludeBlankUnitandCaseUPCCode = 0
			N2KInventoryIncludeBlankUnitBin = 0
			N2KInventoryIncludeBlankUnitUPCCode = 0
			N2KInventoryIncludeDuplicateUnitorCaseBin = 0
			N2KInventoryIncludeDuplicateUPCCode = 0			
		End If
		
		'**************************************************************************'
		'*** Outer loop to get the Summary Description Records
		'**************************************************************************'
		
		SQL = "SELECT SummaryDescription, Count(SummaryDescription) as ProblemCount FROM SC_NeedToKnow WHERE Module = 'Inventory Control' AND InsightStaffOnly <> 1"
		SQL = SQL & "Group By SummaryDescription ORDER BY SummaryDescription"

		Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
		cnnNeed2Know.open (MUV_READ("ClientCnnString"))
		Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
		rsN2KOuter.CursorLocation = 3 
		rsN2KOuter.Open SQL, cnnNeed2Know 
		 
		
		'**************************************************************************'
		'*** Get the current IC_Product SKU count
		'**************************************************************************'
		
		ProductSKUCount = NumberICProductsInventoriedOrPickable()
		
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
					<td><font face="Consolas" style="font-size: 12pt">*Note:* This report only includes products which are set to Inventoried = Y or Pickable = Y</font></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
				</tr>

					
			<%
				If Not rsN2KOuter.EOF Then
				
				
					CountN2KInventoryIncludeBlankCaseBin = 0
					CountN2KInventoryIncludeBlankCaseUPCCode = 0
					CountN2KInventoryIncludeBlankUnitandCaseUPCCode = 0
					CountN2KInventoryIncludeBlankUnitBin = 0
					CountN2KInventoryIncludeBlankUnitUPCCode = 0
					CountN2KInventoryIncludeDuplicateUnitorCaseBin = 0
					CountN2KInventoryIncludeDuplicateUPCCode = 0						
				
				
					Do While Not (rsN2KOuter.EOF)
			
						 SummaryDescription = rsN2KOuter("SummaryDescription")
						 
						 Select Case UCASE(SummaryDescription)
						 
						 	Case "BLANK CASE BIN"
						 		CountN2KInventoryIncludeBlankCaseBin = rsN2KOuter("ProblemCount")
						 	
						 	Case "BLANK CASE UPC CODE"
						 		CountN2KInventoryIncludeBlankCaseUPCCode = rsN2KOuter("ProblemCount")

						 	Case "BLANK UNIT AND CASE UPC CODE"
						 		CountN2KInventoryIncludeBlankUnitandCaseUPCCode = rsN2KOuter("ProblemCount")

						 	Case "BLANK UNIT BIN"
						 		CountN2KInventoryIncludeBlankUnitBin = rsN2KOuter("ProblemCount")

						 	Case "BLANK UNIT UPC CODE"
						 		CountN2KInventoryIncludeBlankUnitUPCCode = rsN2KOuter("ProblemCount")

						 	Case "DUPLICATE UNIT OR CASE BIN"
						 		CountN2KInventoryIncludeDuplicateUnitorCaseBin = rsN2KOuter("ProblemCount")

						 	Case "DUPLICATE UPC CODE"
						 		CountN2KInventoryIncludeDuplicateUPCCode = rsN2KOuter("ProblemCount")
								
						End Select
																								
										
					rsN2KOuter.MoveNext
					Loop
				
						
					If N2KInventoryIncludeBlankCaseBin = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KInventoryIncludeBlankCaseBin %>&nbsp;products have a Blank Case Bin</font></td></tr><%
					End If
		
					If N2KInventoryIncludeBlankCaseUPCCode = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KInventoryIncludeBlankCaseUPCCode %>&nbsp;products have a Blank Case UPC Code</font></td></tr><%
					End If
				
					If N2KInventoryIncludeBlankUnitandCaseUPCCode = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KInventoryIncludeBlankUnitandCaseUPCCode %>&nbsp;products have a Blank Unit and Case UPC Code</font></td></tr><%
					End If

					If N2KInventoryIncludeBlankUnitBin = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KInventoryIncludeBlankUnitBin %>&nbsp;products have a Blank Unit Bin</font></td></tr><%
					End If

					If N2KInventoryIncludeBlankUnitUPCCode = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KInventoryIncludeBlankUnitUPCCode %>&nbsp;products have a Blank Unit UPC Code</font></td></tr><%
					End If

					If N2KInventoryIncludeDuplicateUnitorCaseBin = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KInventoryIncludeDuplicateUnitorCaseBin %>&nbsp;products have a Duplicate Unit or Case Bin</font></td></tr><%
					End If

					If N2KInventoryIncludeDuplicateUPCCode = 1 Then
						%><tr><td><font face="Consolas" style="font-size: 14pt"><%= CountN2KInventoryIncludeDuplicateUPCCode %>&nbsp;products have a Duplicate UPC Code</font></td></tr><%
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
			' B L A N K   C A S E    B I N
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 49
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Inventory Control' AND SummaryDescription = 'Blank Case Bin' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KInventoryIncludeBlankCaseBin = 1 AND CountN2KInventoryIncludeBlankCaseBin > 0 Then
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KInventoryIncludeBlankCaseBin) = cInt(ProductSKUCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Blank Case Bin")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All products have a Blank Case Bin.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KInventoryIncludeBlankCaseBin/ProductSKUCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Blank Case Bin")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of products have a Blank Case Bin.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Blank Case Bin")
				
					Do While Not rsN2KOuter.EOF
					
						CategoryID = GetCategoryIDByProdSKU(rsN2KOuter("prodSKUIfApplicable"))
						If CategoryID <> "" Then
							CategoryName = GetCategoryByID(CategoryID)
						Else
							CategoryName = ""
						End If
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("prodSKUIfApplicable") %></font>
								</td>
								<td width="30%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetProdDescByprodSKU(rsN2KOuter("prodSKUIfApplicable")) %></font>
								</td>
								<td width="25%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= CategoryName %></font>
								</td>
								<td width="30%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetProdUnitBinByprodSKU(rsN2KOuter("prodSKUIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Blank Case Bin")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    B L A N K   C A S E    B I N
			'*************************************************************************************************************



			'*************************************************************************************************************
			' B L A N K   C A S E    U P C   C O D E
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 49
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Inventory Control' AND SummaryDescription = 'Blank Case UPC Code' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KInventoryIncludeBlankCaseUPCCode = 1 AND CountN2KInventoryIncludeBlankCaseUPCCode > 0 Then
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KInventoryIncludeBlankCaseUPCCode) = cInt(ProductSKUCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Blank Case UPC Code")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All products have a Blank Case UPC Code.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KInventoryIncludeBlankCaseUPCCode/ProductSKUCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Blank Case UPC Code")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of products have a Blank Case UPC Code.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Blank Case UPC Code")
				
					Do While Not rsN2KOuter.EOF
					
					
						CategoryID = GetCategoryIDByProdSKU(rsN2KOuter("prodSKUIfApplicable"))
						If CategoryID <> "" Then
							CategoryName = GetCategoryByID(CategoryID)
						Else
							CategoryName = ""
						End If
					
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("prodSKUIfApplicable") %></font>
								</td>
								<td width="50%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetProdDescByprodSKU(rsN2KOuter("prodSKUIfApplicable")) %></font>
								</td>
								<td width="35%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= CategoryName %></font>
								</td>								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Blank Case UPC Code")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    B L A N K     C A S E    U P C   C O D E
			'*************************************************************************************************************






			'*************************************************************************************************************
			' B L A N K   U N I T    A N D    C A S E    U P C   C O D E
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 49
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Inventory Control' AND SummaryDescription = 'Blank Unit and Case UPC Code' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KInventoryIncludeBlankUnitandCaseUPCCode = 1 AND CountN2KInventoryIncludeBlankUnitandCaseUPCCode > 0 Then
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KInventoryIncludeBlankUnitandCaseUPCCode) = cInt(ProductSKUCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Blank Unit and Case UPC Code")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All products have a Blank Unit and Case UPC Code.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KInventoryIncludeBlankUnitandCaseUPCCode/ProductSKUCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Blank Unit and Case UPC Code")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of products have a Blank Unit and Case UPC Code.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Blank Unit and Case UPC Code")
				
					Do While Not rsN2KOuter.EOF
					
					
						CategoryID = GetCategoryIDByProdSKU(rsN2KOuter("prodSKUIfApplicable"))
						If CategoryID <> "" Then
							CategoryName = GetCategoryByID(CategoryID)
						Else
							CategoryName = ""
						End If
					
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("prodSKUIfApplicable") %></font>
								</td>
								<td width="50%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetProdDescByprodSKU(rsN2KOuter("prodSKUIfApplicable")) %></font>
								</td>
								<td width="35%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= CategoryName %></font>
								</td>	
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Blank Unit and Case UPC Code")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    B L A N K   U N I T    A N D    C A S E    U P C   C O D E
			'*************************************************************************************************************




			'*************************************************************************************************************
			' B L A N K   U N I T   B I N
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 49
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Inventory Control' AND SummaryDescription = 'Blank Unit Bin' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KInventoryIncludeBlankUnitBin = 1 AND CountN2KInventoryIncludeBlankUnitBin > 0 Then
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KInventoryIncludeBlankUnitBin) = cInt(ProductSKUCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Blank Unit Bin")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All products have a Blank Unit Bin.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KInventoryIncludeBlankUnitBin/ProductSKUCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Blank Unit Bin")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of products have a Blank Unit Bin.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Blank Unit Bin")
				
					Do While Not rsN2KOuter.EOF
					
						CategoryID = GetCategoryIDByProdSKU(rsN2KOuter("prodSKUIfApplicable"))
						If CategoryID <> "" Then
							CategoryName = GetCategoryByID(CategoryID)
						Else
							CategoryName = ""
						End If
					
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("prodSKUIfApplicable") %></font>
								</td>
								<td width="30%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetProdDescByprodSKU(rsN2KOuter("prodSKUIfApplicable")) %></font>
								</td>
								<td width="25%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= CategoryName %></font>
								</td>
								<td width="30%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetProdCaseBinByprodSKU(rsN2KOuter("prodSKUIfApplicable")) %></font>
								</td>
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Blank Unit Bin")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    B L A N K   U N I T   B I N
			'*************************************************************************************************************







			'*************************************************************************************************************
			' B L A N K    U N I T    U P C    C O D E
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 10
			LinesPerPage = 49
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Inventory Control' AND SummaryDescription = 'Blank Unit UPC Code' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KInventoryIncludeBlankUnitUPCCode = 1 AND CountN2KInventoryIncludeBlankUnitUPCCode > 0 Then
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KInventoryIncludeBlankUnitUPCCode) = cInt(ProductSKUCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Blank Unit UPC Code")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All products have a Blank Unit UPC Code.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KInventoryIncludeBlankUnitUPCCode/ProductSKUCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Blank Unit UPC Code")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of products have a Blank Unit UPC Code.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Blank Unit UPC Code")
				
					Do While Not rsN2KOuter.EOF
					
						CategoryID = GetCategoryIDByProdSKU(rsN2KOuter("prodSKUIfApplicable"))
						If CategoryID <> "" Then
							CategoryName = GetCategoryByID(CategoryID)
						Else
							CategoryName = ""
						End If
					
						%>
							<tr>						
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("prodSKUIfApplicable") %></font>
								</td>
								<td width="50%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= GetProdDescByprodSKU(rsN2KOuter("prodSKUIfApplicable")) %></font>
								</td>
								<td width="35%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= CategoryName %></font>
								</td>
								
								<% RowCount = RowCount + 1 %>
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Blank Unit UPC Code")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    B L A N K    U N I T    U P C    C O D E
			'*************************************************************************************************************




			'*************************************************************************************************************
			' D U P L I C A T E    U N I T    O R    C A S E    B I N
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 8
			LinesPerPage = 42
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Inventory Control' AND SummaryDescription = 'Duplicate Unit or Case Bin' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KInventoryIncludeDuplicateUnitorCaseBin = 1 AND CountN2KInventoryIncludeDuplicateUnitorCaseBin > 0 Then
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KInventoryIncludeDuplicateUnitorCaseBin) = cInt(ProductSKUCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Duplicate Unit or Case Bin")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All products have a Duplicate Unit or Case Bin.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KInventoryIncludeDuplicateUnitorCaseBin/ProductSKUCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Duplicate Unit or Case Bin")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of products have a Duplicate Unit or Case Bin.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Duplicate Unit or Case Bin")
				
					Do While Not rsN2KOuter.EOF
					
						
						CategoryID = GetCategoryIDByProdSKU(rsN2KOuter("prodSKUIfApplicable"))
						If CategoryID <> "" Then
							CategoryName = GetCategoryByID(CategoryID)
						Else
							CategoryName = ""
						End If
					
						%>
							<tr>	
								<%					
								NumDetailLines = 0 
								NumDetailLines = int(Len(rsN2KOuter("DetailedDescription1")) / 120)
								If Len(rsN2KOuter("DetailedDescription1")) MOD 120 <> 0 Then NumDetailLines = NumDetailLines  + 1
								' If it's going to span a page break, break it now
								If RowCount + NumDetailLines > LinesPerPage Then
									Call Footer(RowCount,LinesPerPage)
									Call PageHeader
									Call SubHeader("Duplicate Unit or Case Bin")
									%><tr><%
								End If
								ReDim DetailLinesArray(NumDetailLines)
								For x = 0 to NumDetailLines-1
									If x = 0 Then
										DetailLinesArray(x) = Mid(rsN2KOuter("DetailedDescription1"),1,120)
									ElseIf x = 1 Then
										DetailLinesArray(x) = Mid(rsN2KOuter("DetailedDescription1"),121,120)
									ElseIf x = NumDetailLines-1 Then 
										DetailLinesArray(x) = Mid(rsN2KOuter("DetailedDescription1"),(x+1)*(120*x),Len(rsN2KOuter("DetailedDescription1"))-(120*x))
									Else
										DetailLinesArray(x) = Mid(rsN2KOuter("DetailedDescription1"),(x+1)*120,120)
									End If
								Next
								
								%>
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("prodSKUIfApplicable") %></font>
								</td>
								<%
								For z = 0 to Ubound(DetailLinesArray) -1
									If z > 0 Then Response.Write("<tr><td>&nbsp;</td>")
									%>
									<td width="50%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= DetailLinesArray(z)%></font>
									</td>
									<%
									If z > 0 Then Response.Write("</tr>")
									RowCount = RowCount + 1
								 Next
								 %>
								<td width="35%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= CategoryName %></font>
								</td>
								 
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Duplicate Unit or Case Bin")
							RowCount = 0
						End If
						
					Loop
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    D U P L I C A T E    U N I T    O R    C A S E    B I N
			'*************************************************************************************************************






			'*************************************************************************************************************
			' D U P L I C A T E    U P C   C O D E
			'*************************************************************************************************************
			
			%>
			<table border="0" width="855" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<%
			
			RowCount = 0
			FontSizeVar = 8
			LinesPerPage = 42
			
			
			SQL = "SELECT * FROM SC_NeedToKnow WHERE Module = 'Inventory Control' AND SummaryDescription = 'Duplicate UPC Code' AND InsightStaffOnly <> 1"
			SQL = SQL & " ORDER BY SummaryDescription, custIDIfApplicable"

			Set cnnNeed2Know = Server.CreateObject("ADODB.Connection")
			cnnNeed2Know.open (MUV_READ("ClientCnnString"))
			Set rsN2KOuter  = Server.CreateObject("ADODB.Recordset")
			rsN2KOuter.CursorLocation = 3 
			rsN2KOuter.Open SQL, cnnNeed2Know 
			
			
			'*****MAKE SURE THERE ARE RECORDS AND THAT THIS SECTION IS TURNED ON**********
			If Not rsN2KOuter.EOF AND N2KInventoryIncludeDuplicateUPCCode = 1 AND CountN2KInventoryIncludeDuplicateUPCCode > 0 Then
								
				'*******************************ALL PRODUCTS************************************
			 	If cInt(CountN2KInventoryIncludeDuplicateUPCCode) = cInt(ProductSKUCount) Then
			 	'*******************************ALL PRODUCTS************************************
			 	
			 		Call PageHeader
					Call SubHeader("Duplicate UPC Code")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">All products have a Duplicate UPC Code.</font>
						</td>
					</tr>
					
					</table>
					<%	
					Call Footer(RowCount+5,LinesPerPage)
				
				'*******************************75% Threshold************************************	
				ElseIf (((CountN2KInventoryIncludeDuplicateUPCCode/ProductSKUCount)*100) >= 75) Then
				'*******************************75% Threshold************************************
				
			 		Call PageHeader
					Call SubHeader("Duplicate UPC Code")
					%>
					<tr><td>&nbsp;</td></tr>
					<tr><td>&nbsp;</td></tr>					
					<tr>						
						<td width="100%" style="white-space: nowrap;">
							<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Greater than 75% of products have a Duplicate UPC Code.</font>
						</td>
					</tr>
					
					</table>
					<%						
					Call Footer(RowCount+5,LinesPerPage)

				'*******************************We have records to write************************************	
				Else
				
					Call PageHeader
					Call SubHeader("Duplicate UPC Code")
				
					Do While Not rsN2KOuter.EOF
					
					
						CategoryID = GetCategoryIDByProdSKU(rsN2KOuter("prodSKUIfApplicable"))
						If CategoryID <> "" Then
							CategoryName = GetCategoryByID(CategoryID)
						Else
							CategoryName = ""
						End If
					
						%>
							<tr>	
								<%					
								NumDetailLines = 0 
								NumDetailLines = int(Len(rsN2KOuter("DetailedDescription1")) / 120)
								If Len(rsN2KOuter("DetailedDescription1")) MOD 120 <> 0 Then NumDetailLines = NumDetailLines  + 1
								' If it's going to span a page break, break it now
								If RowCount + NumDetailLines > LinesPerPage Then
									Call Footer(RowCount,LinesPerPage)
									Call PageHeader
									Call SubHeader("Duplicate UPC Code")
									%><tr><%
								End If
								ReDim DetailLinesArray(NumDetailLines)
								For x = 0 to NumDetailLines-1
									If x = 0 Then
										DetailLinesArray(x) = Mid(rsN2KOuter("DetailedDescription1"),1,120)
									ElseIf x = 1 Then
										DetailLinesArray(x) = Mid(rsN2KOuter("DetailedDescription1"),121,120)
									ElseIf x = NumDetailLines-1 Then 
										DetailLinesArray(x) = Mid(rsN2KOuter("DetailedDescription1"),(x+1)*(120*x),Len(rsN2KOuter("DetailedDescription1"))-(120*x))
									Else
										DetailLinesArray(x) = Mid(rsN2KOuter("DetailedDescription1"),(x+1)*120,120)
									End If
								Next
								%>
								<td width="15%">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsN2KOuter("prodSKUIfApplicable") %></font>
								</td>
								<%
								For z = 0 to Ubound(DetailLinesArray) -1
									If z > 0 Then Response.Write("<tr><td>&nbsp;</td>")
									%>
									<td width="50%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= DetailLinesArray(z)%></font>
									</td>
									<%
									If z > 0 Then Response.Write("</tr>")
									RowCount = RowCount + 1
								 Next
								 %>
								<td width="35%" style="white-space: nowrap;">
									<font face="Consolas"  style="font-size: <%= FontSizeVar %>pt"><%= CategoryName %></font>
								</td>
								 
							</tr>
						<%
						
						rsN2KOuter.Movenext	
		
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer(RowCount,LinesPerPage)
							Call PageHeader
							Call SubHeader("Duplicate UPC Code")
							RowCount = 0
						End If
						
					Loop
					
					NoBreak = True
					Call Footer(RowCount,LinesPerPage)
					
				End If
	
			End If
			
			'*************************************************************************************************************
			' E N D    D U P L I C A T E    U P C   C O D E
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
				<p align="center"><b><font face="Consolas" size="4">MDS Insight Inventory<br>Need To Know Report</font></b></p>
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
		<td width="15%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Product ID</font></u></strong>
		</td>
		<% If HeaderText = "Duplicate UPC Code" or HeaderText = "Duplicate Unit or Case Bin" Then%>
			<td width="45%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Details</font></u></strong>
			</td>
		<% ElseIf HeaderText = "Blank Case Bin" Then%>
			<td width="35%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Product Description</font></u></strong>
			</td>
			<td width="15%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Product Category</font></u></strong>
			</td>
			<td width="10%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Unit Bin</font></u></strong>
			</td>
		<% ElseIf HeaderText = "Blank Unit Bin" Then%>
			<td width="35%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Product Description</font></u></strong>
			</td>
			<td width="15%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Product Category</font></u></strong>
			</td>			
			<td width="10%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Case Bin</font></u></strong>
			</td>			
		<% Else %>
			<td width="35%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Product Description</font></u></strong>
			</td>
			<td width="15%">
				<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Product Category</font></u></strong>
			</td>			
		<% End If %>
	</tr>
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
				<font face="Consolas" style="font-size: 9pt">directlaunch/needtoknowreports/inventory.asp</font>
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