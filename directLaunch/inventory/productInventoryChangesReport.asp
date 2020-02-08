<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear
Dim PageNum, RowCount, FontSizeVar
FontSizeVar = 9
PageNum = 0
NoBreak = False
'Adjust = -3
'MAdjust = -1
PageWidth = 1200

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
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Inventory Product Changes Report<%
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
Set cnn_Settings_InventoryControl = Server.CreateObject("ADODB.Connection")
cnn_Settings_InventoryControl.open (MUV_READ("ClientCnnString"))
Set rs_Settings_InventoryControl = Server.CreateObject("ADODB.Recordset")
rs_Settings_InventoryControl.CursorLocation = 3 

SQL_Settings_InventoryControl = "SELECT * FROM Settings_InventoryControl"

Set rs_Settings_InventoryControl = cnn_Settings_InventoryControl.Execute(SQL_Settings_InventoryControl)
If not rs_Settings_InventoryControl.EOF Then
	InventoryProductChangesReportOnOff = rs_Settings_InventoryControl("InventoryProductChangesReportOnOff")
Else
	InventoryProductChangesReportOnOff = 0
End If

Set rs_Settings_InventoryControl = Nothing
cnn_Settings_InventoryControl.Close
Set cnn_Settings_InventoryControl = Nothing

If InventoryProductChangesReportOnOff <> 1 Then
	%>MDS Insight: The inventory product changes report is not turned on.
	<%
	Response.End
End IF
%>

<body bgcolor="#FFFFFF" text="#000000" link="#000080" vlink="#000080" alink="#000080" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<table border="0" width="<%=PageWidth%>" align="center">
	<tr>
		<td width="100%" align="center">

		<%
		
		If PageNum = 0 Then

			'*******************************************************
			'*** This section is the first page which prints all the
			'*** product changes summary info
			'*******************************************************
		
			Call PageHeader

			LinesPerPage = 2
			
			
			%><!--#include file="productInventoryChangesReportChangeCounts.asp"--><%
						
			If NoInventoryControlBackups = 1 Then
				%>MDS Insight: The inventory product changes report is turned on but there are no backups to compare to.
				<%
				Response.End
			End IF
			
			%>
			<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
				<tr>
					<td>
					<font face="Consolas">
					<hr>
					<center><h2>Summary for <%= Date() %></h2></center>
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
				
				<% If ICProductRowsAdded  = 0 AND ICProductRowsDeleted = 0 AND ICProductUnitDescriptionRowChanged = 0 AND ICProductUnitCostRowsChanged = 0 AND ICProductUnitPriceRowsChanged = 0_
					  AND ICProductBinNoRowsChanged = 0_
					  AND ICProductPerpetualFlagRowsChanged = 0 AND ICProductCasePricingRowsChanged = 0 AND ICProductBinCaseRowsChanged = 0_
					  AND ICProductCaseDescriptionRowsChanged = 0 AND ICProductInventoriedRowsChanged = 0 AND ICProductPickableRowsChanged = 0 Then
					  
					'*********************************
					'ORDER OF FIELDS TO DISPLAY:
					'*********************************
					'prodSKU**
					'prodDescription**
					'prodCaseDescription**
					'prodUnitCost**
					'prodPriceLvl1
					'prodCasePricing**
					'prodUnitBin**
					'prodCaseBin**
					'DisplayOnWeb**
					'prodInventoriedItem**
					'prodPickableItem**
					'*********************************						  
				%>
					
						<tr><td>&nbsp;</td></tr>
						<tr><td>&nbsp;</td></tr>
						<tr><td>&nbsp;</td></tr>
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td>
								<font face="Consolas" style="font-size: 14pt">
								<hr>
								<center>Great news! There are no product changes to report.</center></font>
								<hr>
								<% NoBreak = True %>
							</td>
						</tr>
						<% RowCount = 0

				Else %>
				
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductRowsAdded %>&nbsp;products were added</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductRowsDeleted %>&nbsp;products were deleted</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductUnitDescriptionRowChanged %>&nbsp;product unit descriptions have changed</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductCaseDescriptionRowsChanged %>&nbsp;product case descriptions have changed</font>
						</td>
					</tr>	
					<tr><td>&nbsp;</td></tr>					
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductUnitCostRowsChanged %>&nbsp;product unit costs have changed</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductUnitPriceRowsChanged %>&nbsp;product unit prices have changed</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductCasePricingRowsChanged %>&nbsp;product cases prices have changed</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductBinNoRowsChanged %>&nbsp;product unit bins have changed</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductBinCaseRowsChanged %>&nbsp;product case bins have changed</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductPerpetualFlagRowsChanged %>&nbsp;products have changed perpetual flag statuses</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>						
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductInventoriedRowsChanged %>&nbsp;products have changed inventory statuses</font>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td>
							<font face="Consolas" style="font-size: 14pt">
							<%= ICProductPickableRowsChanged %>&nbsp;products have changed pickable statuses</font>
						</td>
					</tr>
					<%
				

				End If %>
				
		<% End If %>
		</table>
		</td>
	</tr>
	<tr>
	
	<td>
	
<% 

			Call Footer
		

			'******************************************************************************************
			'*** This section is the detail section of all the product changes that have been found
			'******************************************************************************************
			
			'Now we start doing all the individual product change detail sections
			
			
			'************************************
			'ORDER OF FIELD CHANGES TO DISPLAY:
			'************************************
			'products added to IC_Product (prodSKU)
			'products deleted from IC_Product (prodSKU)
			'prodUnitDescription**
			'prodCaseDescription**
			'prodUnitCost**
			'prodPriceLvl1**
			'prodCasePricing**
			'prodUnitBin**
			'prodCaseBin**
			'DisplayOnWeb**
			'prodInventoriedItem**
			'prodPickableItem**
			'*********************************					
			
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			If ICProductRowsAdded > 0 Then
			
				HeldPageDescription = "PRODUCTS ADDED"
				
				'*********************************************************************************************************************
				'GET DETAILS OF TOTAL ROWS ADDED TO THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
			
				SQL = "SELECT prodSKU, prodDescription, prodCategory FROM IC_Product WHERE prodSKU IN "
				SQL = SQL & " (SELECT  prodSKU "
				SQL = SQL & " FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU "
				SQL = SQL & " FROM  " & ICProductBackupTableName & ") "
			
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
												
						FontSizeVar = 9
						LinesPerPage = 36
						
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<td width="55%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">This product has been added.</font>
							</td>
						 			 
						</tr>
					<%
					
						RowCount = RowCount + 1
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			If ICProductRowsDeleted > 0 or 1 = 1 Then
			
				HeldPageDescription = "PRODUCTS DELETED"
				
				'*********************************************************************************************************************
				'GET DETAILS OF TOTAL ROWS DELETED THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
			
				SQL = "SELECT prodSKU, prodDescription, prodCategory FROM " & ICProductBackupTableName & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT  prodSKU "
				SQL = SQL & " FROM " & ICProductBackupTableName & " "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU "
				SQL = SQL & " FROM  IC_Product) "

				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 36
						
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<td width="55%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">This product has been deleted.</font>
							</td>
						 			 
						</tr>
					<%
					
						RowCount = RowCount + 1
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			

			If ICProductUnitDescriptionRowChanged > 0 Then
			
				HeldPageDescription = "PRODUCTS WITH CHANGED UNIT DESCRIPTIONS"
				
				'*********************************************************************************************************************
				'GET DETAILS OF PRODUCT WITH CHANGED UNIT DESCRIPTIONS IN THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
			
				SQL = "SELECT prodSKU, prodDescription, prodCategory FROM IC_Product "
				SQL = SQL & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT prodSKU from "
				SQL = SQL & " (SELECT prodSKU,prodDescription FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU,prodDescription FROM " & ICProductBackupTableName & ") as t) "			
				SQL = SQL & " AND prodSKU NOT IN (" & addedProductsString & ")"
				
				
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 19
						
						'**********GET DETAILS OF CHANGE*****************************************
						
						Set cnnProductChangesDetail = Server.CreateObject("ADODB.Connection")
						cnnProductChangesDetail.open(MUV_READ("ClientCnnString"))
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodDescription FROM IC_Product WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodDescriptionToday = rsProductChangesDetail("prodDescription")
						
						Set rsProductChangesDetail = Nothing
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodDescription FROM " & ICProductBackupTableName & " WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodDescriptionYesterday = rsProductChangesDetail("prodDescription")
						
						Set rsProductChangesDetail = Nothing
						cnnProductChangesDetail.Close
						Set cnnProductChangesDetail = Nothing	
						
						lengthOfUnitDescription	= Len("The unit description has changed from " & prodDescriptionYesterday & " to " & prodDescriptionToday & ".")				
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<% If lengthOfUnitDescription >= 225 Then %>
								<td width="55%" height="35">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The unit description has changed from <strong><%= prodDescriptionYesterday %></strong> to <strong><%= prodDescriptionToday %></strong>.</font>
								</td>
						 	<% Else %>
								<td width="55%" height="25">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The unit description has changed from <strong><%= prodDescriptionYesterday %></strong> to <strong><%= prodDescriptionToday %></strong>.</font>
								</td>						 	
							<% End If %>		 
						</tr>
					<%
					
						If lengthOfUnitDescription >= 225 Then
							RowCount = RowCount + 2
						Else
							RowCount = RowCount + 1
						End If
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
			
			If ICProductCaseDescriptionRowsChanged > 0 Then
			
				HeldPageDescription = "PRODUCTS WITH CHANGED CASE DESCRIPTIONS"
				
				'*********************************************************************************************************************
				'GET DETAILS OF PRODUCT WITH CHANGED UNIT DESCRIPTIONS IN THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
			
				SQL = "SELECT prodSKU, prodDescription, prodCaseDescription, prodCategory FROM IC_Product "
				SQL = SQL & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT prodSKU from "
				SQL = SQL & " (SELECT prodSKU,prodCaseDescription FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU,prodCaseDescription FROM " & ICProductBackupTableName & ") as t) "			
				SQL = SQL & " AND prodSKU NOT IN (" & addedProductsString & ")"
				
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 19
						
						'**********GET DETAILS OF CHANGE*****************************************
						
						Set cnnProductChangesDetail = Server.CreateObject("ADODB.Connection")
						cnnProductChangesDetail.open(MUV_READ("ClientCnnString"))
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodCaseDescription FROM IC_Product WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodCaseDescriptionToday = rsProductChangesDetail("prodCaseDescription")
						
						Set rsProductChangesDetail = Nothing
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodCaseDescription FROM " & ICProductBackupTableName & " WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodCaseDescriptionYesterday = rsProductChangesDetail("prodCaseDescription")
						
						Set rsProductChangesDetail = Nothing
						cnnProductChangesDetail.Close
						Set cnnProductChangesDetail = Nothing	
						
						lengthOfCaseDescription	= Len("The case description has changed from " & prodCaseDescriptionYesterday & " to " & prodCaseDescriptionToday & ".")					
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>
							<% If lengthOfCaseDescription >= 225 Then %>
								<td width="55%" height="35">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The case description has changed from <strong><%= prodCaseDescriptionYesterday %></strong> to <strong><%= prodCaseDescriptionToday %></strong>.</font>
								</td>
						 	<% Else %>
								<td width="55%" height="25">
									<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The case description has changed from <strong><%= prodCaseDescriptionYesterday %></strong> to <strong><%= prodCaseDescriptionToday %></strong>.</font>
								</td>						 	
							<% End If %>								
						 			 
						</tr>
					<%
					
						If lengthOfCaseDescription >= 225 Then
							RowCount = RowCount + 2
						Else
							RowCount = RowCount + 1
						End If
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			


			If ICProductUnitCostRowsChanged > 0 Then
			
				HeldPageDescription = "PRODUCTS WITH CHANGED UNIT COSTS"
				
				'*********************************************************************************************************************
				'GET DETAILS OF PRODUCT WITH A CHANGED UNIT COST IN THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
			
				SQL = "SELECT prodSKU, prodDescription, prodUnitCost, prodCategory FROM IC_Product "
				SQL = SQL & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT prodSKU from "
				SQL = SQL & " (SELECT prodSKU,prodUnitCost FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU,prodUnitCost FROM " & ICProductBackupTableName & ") as t) "			
				SQL = SQL & " AND prodSKU NOT IN (" & addedProductsString & ")"
			
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 36

						'**********GET DETAILS OF CHANGE*****************************************
						
						Set cnnProductChangesDetail = Server.CreateObject("ADODB.Connection")
						cnnProductChangesDetail.open(MUV_READ("ClientCnnString"))
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodUnitCost FROM IC_Product WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodUnitCostToday = rsProductChangesDetail("prodUnitCost")
						
						Set rsProductChangesDetail = Nothing
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodUnitCost FROM " & ICProductBackupTableName & " WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodUnitCostYesterday = rsProductChangesDetail("prodUnitCost")
						
						Set rsProductChangesDetail = Nothing
						cnnProductChangesDetail.Close
						Set cnnProductChangesDetail = Nothing						
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<td width="55%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The unit cost has changed from <strong><%= prodUnitCostYesterday %></strong> to <strong><%= prodUnitCostToday %></strong>.</font>
							</td>
						 			 
						</tr>
					<%
					
						RowCount = RowCount + 1
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
	

			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			


			If ICProductUnitPriceRowsChanged > 0 Then
			
				HeldPageDescription = "PRODUCTS WITH CHANGED UNIT PRICES"
				
				'*********************************************************************************************************************
				'GET DETAILS OF PRODUCT WITH A CHANGED UNIT PRICE IN THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
				
				SQL = "SELECT prodSKU, prodDescription, prodPriceLvl1, prodCategory FROM IC_Product "
				SQL = SQL & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT prodSKU from "
				SQL = SQL & " (SELECT prodSKU,prodPriceLvl1 FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU,prodPriceLvl1 FROM " & ICProductBackupTableName & ") as t) "			
				SQL = SQL & " AND prodSKU NOT IN (" & addedProductsString & ")"
				
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 36
						
						'**********GET DETAILS OF CHANGE*****************************************
						
						Set cnnProductChangesDetail = Server.CreateObject("ADODB.Connection")
						cnnProductChangesDetail.open(MUV_READ("ClientCnnString"))
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodPriceLvl1 FROM IC_Product WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodPriceLvl1Today = rsProductChangesDetail("prodPriceLvl1")
						
						Set rsProductChangesDetail = Nothing
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodPriceLvl1 FROM " & ICProductBackupTableName & " WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodPriceLvl1Yesterday = rsProductChangesDetail("prodPriceLvl1")
						
						Set rsProductChangesDetail = Nothing
						cnnProductChangesDetail.Close
						Set cnnProductChangesDetail = Nothing						
						
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<td width="55%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The unit price has changed from <strong><%= prodPriceLvl1Yesterday %></strong> to <strong><%= prodPriceLvl1Today %></strong>.</font>
							</td>
						 			 
						</tr>
					<%
					
						RowCount = RowCount + 1
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



			If ICProductCasePricingRowsChanged > 0 Then
			
				HeldPageDescription = "PRODUCTS WITH CHANGED CASE PRICES"
				
				'*********************************************************************************************************************
				'GET DETAILS OF PRODUCT WITH A CHANGED CASE PRICE IN THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
			
				SQL = "SELECT prodSKU, prodDescription, prodCasePricing, prodCategory FROM IC_Product "
				SQL = SQL & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT prodSKU from "
				SQL = SQL & " (SELECT prodSKU,prodCasePricing FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU,prodCasePricing FROM " & ICProductBackupTableName & ") as t) "			
				SQL = SQL & " AND prodSKU NOT IN (" & addedProductsString & ")"
				
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 36
						
						'**********GET DETAILS OF CHANGE*****************************************
						
						Set cnnProductChangesDetail = Server.CreateObject("ADODB.Connection")
						cnnProductChangesDetail.open(MUV_READ("ClientCnnString"))
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodCasePricing FROM IC_Product WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodCasePricingToday = rsProductChangesDetail("prodCasePricing")
						
						Set rsProductChangesDetail = Nothing
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodCasePricing FROM " & ICProductBackupTableName & " WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodCasePricingYesterday = rsProductChangesDetail("prodCasePricing")
						
						Set rsProductChangesDetail = Nothing
						cnnProductChangesDetail.Close
						Set cnnProductChangesDetail = Nothing								
						
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<td width="55%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The case price has changed from <strong><%= prodCasePricingYesterday %></strong> to <strong><%= prodCasePricingToday %></strong>.</font>
							</td>
						 			 
						</tr>
					<%
					
						RowCount = RowCount + 1
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop

					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	


			If ICProductBinNoRowsChanged > 0 Then
			
				HeldPageDescription = "PRODUCTS WITH CHANGED UNIT BINS"
				
				'*********************************************************************************************************************
				'GET DETAILS OF PRODUCT WITH A CHANGED UNIT BINS IN THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
			
				SQL = "SELECT prodSKU, prodDescription, prodUnitBin, prodCategory FROM IC_Product "
				SQL = SQL & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT prodSKU from "
				SQL = SQL & " (SELECT prodSKU,prodUnitBin FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU,prodUnitBin FROM " & ICProductBackupTableName & ") as t) "			
				SQL = SQL & " AND prodSKU NOT IN (" & addedProductsString & ")"
				
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 36
						
						'**********GET DETAILS OF CHANGE*****************************************
						
						Set cnnProductChangesDetail = Server.CreateObject("ADODB.Connection")
						cnnProductChangesDetail.open(MUV_READ("ClientCnnString"))
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodUnitBin FROM IC_Product WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodUnitBinToday = rsProductChangesDetail("prodUnitBin")
						
						Set rsProductChangesDetail = Nothing
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodUnitBin FROM " & ICProductBackupTableName & " WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodUnitBinYesterday = rsProductChangesDetail("prodUnitBin")
						
						Set rsProductChangesDetail = Nothing
						cnnProductChangesDetail.Close
						Set cnnProductChangesDetail = Nothing	
													
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<td width="55%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The unit bin has changed from <strong><%= prodUnitBinYesterday %></strong> to <strong><%= prodUnitBinToday %></strong>.</font>
							</td>
						 			 
						</tr>
					<%
					
						RowCount = RowCount + 1
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			




			If ICProductBinCaseRowsChanged > 0 Then
			
				HeldPageDescription = "PRODUCTS WITH CHANGED CASE BINS"
				
				'*********************************************************************************************************************
				'GET DETAILS OF PRODUCT WITH A CHANGED CASE BINS IN THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
			
				SQL = "SELECT prodSKU, prodDescription, prodCaseBin, prodCategory FROM IC_Product "
				SQL = SQL & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT prodSKU from "
				SQL = SQL & " (SELECT prodSKU,prodCaseBin FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU,prodCaseBin FROM " & ICProductBackupTableName & ") as t) "			
				SQL = SQL & " AND prodSKU NOT IN (" & addedProductsString & ")"
				
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 36
						
						'**********GET DETAILS OF CHANGE*****************************************
						
						Set cnnProductChangesDetail = Server.CreateObject("ADODB.Connection")
						cnnProductChangesDetail.open(MUV_READ("ClientCnnString"))
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodCaseBin FROM IC_Product WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodCaseBinToday = rsProductChangesDetail("prodCaseBin")
						
						Set rsProductChangesDetail = Nothing
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodCaseBin FROM " & ICProductBackupTableName & " WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodCaseBinYesterday = rsProductChangesDetail("prodCaseBin")
						
						Set rsProductChangesDetail = Nothing
						cnnProductChangesDetail.Close
						Set cnnProductChangesDetail = Nothing	
						
						
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<td width="55%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The case bin has changed from <strong><%= prodCaseBinYesterday %></strong> to <strong><%= prodCaseBinToday %></strong>.</font>
							</td>
						 			 
						</tr>
					<%
					
						RowCount = RowCount + 1
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




			If ICProductPerpetualFlagRowsChanged > 0 Then
			
				HeldPageDescription = "PRODUCTS WITH CHANGED PERPETUAL FLAGS"
				
				'*********************************************************************************************************************
				'GET DETAILS OF PRODUCT WITH A CHANGED DisplayOnWeb/PERPETUAL FLAGS IN THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
							
				SQL = "SELECT prodSKU, prodDescription, DisplayOnWeb, prodCategory FROM IC_Product "
				SQL = SQL & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT prodSKU from "
				SQL = SQL & " (SELECT prodSKU,DisplayOnWeb FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU,DisplayOnWeb FROM " & ICProductBackupTableName & ") as t) "			
				SQL = SQL & " AND prodSKU NOT IN (" & addedProductsString & ")"
				
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 36
						
						'**********GET DETAILS OF CHANGE*****************************************
						
						Set cnnProductChangesDetail = Server.CreateObject("ADODB.Connection")
						cnnProductChangesDetail.open(MUV_READ("ClientCnnString"))
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT DisplayOnWeb FROM IC_Product WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						DisplayOnWebToday = rsProductChangesDetail("DisplayOnWeb")
						
						Set rsProductChangesDetail = Nothing
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT DisplayOnWeb FROM " & ICProductBackupTableName & " WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						DisplayOnWebYesterday = rsProductChangesDetail("DisplayOnWeb")
						
						Set rsProductChangesDetail = Nothing
						cnnProductChangesDetail.Close
						Set cnnProductChangesDetail = Nothing	
						
						If UCASE(DisplayOnWebYesterday) = "FALSE" Then	
							DisplayOnWebYesterday = "N"
						End If
						If UCASE(DisplayOnWebToday) = "FALSE" Then	
							DisplayOnWebToday = "N"
						End If
						If UCASE(DisplayOnWebYesterday) = "TRUE" Then	
							DisplayOnWebYesterday = "Y"
						End If
						If UCASE(DisplayOnWebToday) = "TRUE" Then	
							DisplayOnWebToday = "Y"
						End If
						
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<td width="55%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The perpetual flag has changed from <strong><%= DisplayOnWebYesterday %></strong> to <strong><%= DisplayOnWebToday %></strong>.</font>
							</td>
						 			 
						</tr>
					<%
					
						RowCount = RowCount + 1
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			




			If ICProductInventoriedRowsChanged > 0 Then
			
				HeldPageDescription = "PRODUCTS WITH CHANGED INVENTORY FLAGS"
				
				'*********************************************************************************************************************
				'GET DETAILS OF PRODUCT WITH A CHANGED INVENTORY FLAGS IN THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************
							
				SQL = "SELECT prodSKU, prodDescription, prodInventoriedItem, prodCategory FROM IC_Product "
				SQL = SQL & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT prodSKU from "
				SQL = SQL & " (SELECT prodSKU,prodInventoriedItem FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU,prodInventoriedItem FROM " & ICProductBackupTableName & ") as t) "			
				SQL = SQL & " AND prodSKU NOT IN (" & addedProductsString & ")"
				
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 36
						
						'**********GET DETAILS OF CHANGE*****************************************
						
						Set cnnProductChangesDetail = Server.CreateObject("ADODB.Connection")
						cnnProductChangesDetail.open(MUV_READ("ClientCnnString"))
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodInventoriedItem FROM IC_Product WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodInventoriedItemToday = rsProductChangesDetail("prodInventoriedItem")
						
						Set rsProductChangesDetail = Nothing
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodInventoriedItem FROM " & ICProductBackupTableName & " WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodInventoriedItemYesterday = rsProductChangesDetail("prodInventoriedItem")
						
						Set rsProductChangesDetail = Nothing
						cnnProductChangesDetail.Close
						Set cnnProductChangesDetail = Nothing	
							
						If UCASE(prodInventoriedItemYesterday) = "FALSE" Then	
							prodInventoriedItemYesterday = "N"
						End If
						If UCASE(prodInventoriedItemToday) = "FALSE" Then	
							prodInventoriedItemToday = "N"
						End If
						If UCASE(prodInventoriedItemYesterday) = "TRUE" Then	
							prodInventoriedItemYesterday = "Y"
						End If
						If UCASE(prodInventoriedItemToday) = "TRUE" Then	
							prodInventoriedItemToday = "Y"
						End If
					
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<td width="55%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The inventories status has changed from <strong><%= prodInventoriedItemYesterday %></strong> to <strong><%= prodInventoriedItemToday %></strong>.</font>
							</td>
						 			 
						</tr>
					<%
					
						RowCount = RowCount + 1
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

			

			If ICProductPickableRowsChanged > 0 Then
			
				HeldPageDescription = "PRODUCTS WITH CHANGED PICKABLE STATUS FLAGS"
				
				'*********************************************************************************************************************
				'GET DETAILS OF PRODUCT WITH A CHANGED PICKABLE STATUS FLAG IN THE PRODUCTS TABLE TODAY
				'*********************************************************************************************************************

				SQL = "SELECT prodSKU, prodDescription, prodPickableItem, prodCategory FROM IC_Product "
				SQL = SQL & " WHERE prodSKU IN "
				SQL = SQL & " (SELECT prodSKU from "
				SQL = SQL & " (SELECT prodSKU,prodPickableItem FROM IC_Product "
				SQL = SQL & " EXCEPT "
				SQL = SQL & " SELECT prodSKU,prodPickableItem FROM " & ICProductBackupTableName & ") as t) "			
				SQL = SQL & " AND prodSKU NOT IN (" & addedProductsString & ")"
				
				Set cnnProductChanges = Server.CreateObject("ADODB.Connection")
				cnnProductChanges.open (MUV_READ("ClientCnnString"))
				Set rsProductChanges = Server.CreateObject("ADODB.Recordset")
				rsProductChanges.CursorLocation = 3 
				rsProductChanges.Open SQL, cnnProductChanges
				
				If Not rsProductChanges.EOF Then
					
					Call PageHeader
					Call SubHeader
				
					Do While Not rsProductChanges.EOF
	
						FontSizeVar = 9
						LinesPerPage = 36
						
						'**********GET DETAILS OF CHANGE*****************************************
						
						Set cnnProductChangesDetail = Server.CreateObject("ADODB.Connection")
						cnnProductChangesDetail.open(MUV_READ("ClientCnnString"))
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodPickableItem FROM IC_Product WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodPickableItemToday = rsProductChangesDetail("prodPickableItem")
						
						Set rsProductChangesDetail = Nothing
						Set rsProductChangesDetail = Server.CreateObject("ADODB.Recordset")
						
						SQLChangesDetail = "SELECT prodPickableItem FROM " & ICProductBackupTableName & " WHERE prodSKU = '" & rsProductChanges("prodSKU") & "'"
						rsProductChangesDetail.Open SQLChangesDetail, cnnProductChangesDetail
						prodPickableItemYesterday = rsProductChangesDetail("prodPickableItem")
						
						Set rsProductChangesDetail = Nothing
						cnnProductChangesDetail.Close
						Set cnnProductChangesDetail = Nothing	
						
						If UCASE(prodPickableItemYesterday) = "FALSE" Then	
							prodPickableItemYesterday = "N"
						End If
						If UCASE(prodPickableItemToday) = "FALSE" Then	
							prodPickableItemToday = "N"
						End If
						If UCASE(prodPickableItemYesterday) = "TRUE" Then	
							prodPickableItemYesterday = "Y"
						End If
						If UCASE(prodPickableItemToday) = "TRUE" Then	
							prodPickableItemToday = "Y"
						End If
						
						%>
						<tr>
						
							<td width="10%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= rsProductChanges("prodSKU") %></font>
							</td>
							<td width="25%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%If Len(rsProductChanges("prodDescription")) > 100 Then Response.Write(Left(rsProductChanges("prodDescription"),100)) Else Response.Write(rsProductChanges("prodDescription"))  %></font>
							</td>
							<td width="15%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetCategoryByID(rsProductChanges("prodCategory")) %></font>
							</td>							
							<td width="55%">
								<font face="Consolas" style="font-size: <%= FontSizeVar %>pt">The pickable status has changed from <strong><%= prodPickableItemYesterday %></strong> to <strong><%= prodPickableItemToday %></strong>.</font>
							</td>
						 			 
						</tr>
					<%
					
						RowCount = RowCount + 1
	
						rsProductChanges.Movenext	
	
						If RowCount > LinesPerPage Then
							%></table><%
							Call Footer
							Call PageHeader
							Call SubHeader
						End If
						
					
					Loop
						
					NoBreak = True
					Call Footer	
				
				End If
				
			End If
			
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
			
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
				<p align="center"><b><font face="Consolas" size="4">MDS Insight Inventory Product Changes Report</font></b></p>
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
	
	<table border="0" width="<%= PageWidth %>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
		<tr>
			<td>
			<font face="Consolas">
			<hr>
			<center><h2><%= HeldPageDescription %></h2></center>
			<hr>
			</font>
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
	
	<table border="0" width="<%= PageWidth %>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">
	<tr>
		<td width="10%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Product ID</font></u></strong>
		</td>
		<td width="25%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Product Description</font></u></strong>
		</td>
		<td width="15%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Product Category</font></u></strong>
		</td>
		<td width="55%">
			<strong><u><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Noted Change</font></u></strong>
		</td>
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
				<font face="Consolas" style="font-size: 9pt">directlaunch/inventory/productInventoryChangesReport.asp</font>
			</td>
			<td width="33%" align="center"> 
				<font face="Consolas" style="font-size: 12pt">Page:&nbsp;<%=PageNum%></font>
			</td>
			<td width="33%" align="left">
				<font face="Consolas" style="font-size: 9pt">Comparing IC_Product to <%= ICProductBackupTableName %></font>
			</td>
		</tr>
	</table>
	<% If NoBreak <> True Then %>
		<BR style="page-break-after: always">
	<% End If

End Sub




function sortArray(arrShort)

    for i = UBound(arrShort) - 1 To 0 Step -1
        for j= 0 to i
            if arrShort(j)>arrShort(j+1) then
                temp=arrShort(j+1)
                arrShort(j+1)=arrShort(j)
                arrShort(j)=temp
            end if
        next
    next
    sortArray = arrShort

end function

	


%>