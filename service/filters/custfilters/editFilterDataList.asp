<!--#include file="../../../inc/InSightFuncs.asp"--> 
<!--#include file="../../../inc/InSightFuncs_Users.asp"--> 
<!--#include file="../../../inc/InsightFuncs_InventoryControl.asp"--> 
<!--#include file="../../../inc/InsightFuncs_BizIntel.asp"-->
<!--#include file="../../../inc/InsightFuncs_Equipment.asp"-->
<!--#include file="../../../inc/InsightFuncs_Service.asp"-->
<%
	DIM CustName
	DIM customerID
	
	customerID=REQUEST("CustomerID")
	
	Set cnnCustFiltersList = Server.CreateObject("ADODB.Connection")
	cnnCustFiltersList.open (Session("ClientCnnString"))
	
	CustName = GetCustNameByCustNum(REQUEST("CustomerID"))
	SQL = "SELECT * FROM FS_CustomerFilters WHERE CustID=" & REQUEST("CustomerID")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	Set rs = cnnCustFiltersList.Execute(SQL)
	%>
	
			
	
		<input type="hidden" name="EditCustomerFilterCustIDToPass" id="EditCustomerFilterCustIDToPass" value="<%=REQUEST("CustomerID")%>">
		<input type="hidden"  id="EditCustomerFilterCustName" value="<%=CustName%>">
		<div class="filter-program">
		<table class="table table-hover filterList">
			<thead>
				<tr>
					<th style="width:250px;">Filter ID</th>
					<th>Notes</th>
					<th style="width:100px;">Qty</th>
					<th style="width:100px;">Frequency Type</th>
					<th style="width:100px;">Frequency Time</th>
					<th style="width:100px;">Price</th>
					<th style="width:150px;">Last Changed</th>	
				</tr> 
			</thead>
			<tbody>
				
					<%DO WHILE NOT rs.EOF%>
						<tr class="for-edit" onclick="javascript: selectRow(this);">
							<td data-id="FilterData" data-value="<%=rs("FilterIntRecID")%>"><%=GetFilterIDByIntRecID(rs("FilterIntRecID"))%></td>
							<td data-id="location" data-value="<%=rs("notes")%>"><%=rs("notes")%></td>
							<td data-id="qty" data-value="<%=rs("qty")%>"><%=rs("qty")%></td>
							<td data-id="FrequencyType" data-value="<%=rs("FrequencyType")%>"><%=rs("FrequencyType")%></td>
							<td data-id="FrequencyTime" data-value="<%=rs("FrequencyTime")%>"><%=rs("FrequencyTime")%></td>
							<td data-id="Price" data-value="<%=rs("Price")%>"><%=rs("Price")%></td>
							<td data-id="LastChangedData" data-value="<%=rs("LastChangeDateTime")%>"><%=rs("LastChangeDateTime")%></td>
			
						</tr>
			
						<%rs.MoveNext%>
					<%LOOP%>
					
			</tbody>
		</table>
		</div>
		<div class="sub-title"> <h4>&nbsp;</h5><hr /></div>		
		<div class="equipment-list">
			
			<%
	CustIDPassed = customerID
	CustName = GetCustNameByCustNum(CustIDPassed)
	
	'TotalEquipmentValue = GetTotalValueOfEquipmentForCustomer(CustIDPassed)
	TotalEquipmentValue = 0
	
	
	Set rsCustomerEquipmentByClass = Server.CreateObject("ADODB.Recordset")
	rsCustomerEquipmentByClass.CursorLocation = 3 

	Set rsCustomerEquipment = Server.CreateObject("ADODB.Recordset")
	rsCustomerEquipment.CursorLocation = 3 
	
	
	Set rsEquipStatusCode = Server.CreateObject("ADODB.Recordset")
	rsEquipStatusCode.CursorLocation = 3 
	
		
	SQLCustomerEquipmentByClass = "SELECT EQ_Classes.Class, EQ_Classes.InternalRecordIdentifier, SUM(EQ_Equipment.PurchaseCost) AS Expr1 "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " FROM EQ_CustomerEquipment INNER JOIN "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " EQ_Equipment ON EQ_CustomerEquipment.EquipIntRecID = EQ_Equipment.InternalRecordIdentifier INNER JOIN "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier INNER JOIN "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " EQ_Classes ON EQ_Models.ClassIntRecID = EQ_Classes.InternalRecordIdentifier "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " WHERE        (EQ_CustomerEquipment.CustID = " & CustIDPassed & ") "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass  & " AND EQ_Models.MightUseAFilter = 1 "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " GROUP BY EQ_Classes.Class, EQ_Classes.InternalRecordIdentifier "
	SQLCustomerEquipmentByClass  = SQLCustomerEquipmentByClass & " ORDER BY Expr1 DESC"
		
	Set cnnCustomerEquipmentByClass = Server.CreateObject("ADODB.Connection")
	cnnCustomerEquipmentByClass.open (Session("ClientCnnString"))
	Set rsCustomerEquipmentByClass = cnnCustomerEquipmentByClass.Execute(SQLCustomerEquipmentByClass)
	
	If NOT rsCustomerEquipmentByClass.EOF Then
	
		Do While NOT rsCustomerEquipmentByClass.EOF
		
			ClassName = rsCustomerEquipmentByClass("Class")
			ClassIntRecID = rsCustomerEquipmentByClass("InternalRecordIdentifier")
			ClassTotalEquipValue = rsCustomerEquipmentByClass("Expr1")
	
	
			%>	
			<h3><%= ClassName %>&nbsp;<span style="color:green;"><%= FormatCurrency(ClassTotalEquipValue,2) %></span></h3>
			<table class="table table-condensed table-hover large-table">			
				<thead>
				  <tr style="background-color: #EEE;">
				  	<th style="width: 3%;">+</th>
				  	<th style="width: 25%;">Description/Type</th>
				  	<th>Status</th>
				  	<th>Frequency</th>
				  	<th>Rent $</th>
				  	<th style="text-align: center;">Install Date</th>
				  	<th style="text-align: center;">Equip. Value</th>
				  	<th style="text-align: center;">Serial #</th>
				  	<th style="text-align: center;">Asset #</th>
				  </tr>
				</thead>
				<tbody>
				
				<%	
				TotalPurchaseCost = 0 
				
				SQLCustomerEquipment = " SELECT        EQ_Equipment.ModelIntRecID, MAX(EQ_Equipment.PurchaseCost) AS purchsum "
				SQLCustomerEquipment = SQLCustomerEquipment & " FROM            EQ_CustomerEquipment INNER JOIN "
				SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID INNER JOIN "
				SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
				SQLCustomerEquipment = SQLCustomerEquipment & " WHERE        (EQ_CustomerEquipment.CustID = " & CustIDPassed & ") AND (EQ_Models.ClassIntRecID = " & ClassIntRecID & ") "
				SQLCustomerEquipment = SQLCustomerEquipment & " AND EQ_Models.MightUseAFilter = 1 "
				SQLCustomerEquipment = SQLCustomerEquipment & " GROUP BY EQ_Equipment.ModelIntRecID "
				SQLCustomerEquipment = SQLCustomerEquipment & " ORDER BY purchsum DESC "		
										
				Set cnnCustomerEquipment = Server.CreateObject("ADODB.Connection")
				cnnCustomerEquipment.open (Session("ClientCnnString"))
				Set rsCustomerEquipment = cnnCustomerEquipment.Execute(SQLCustomerEquipment)
				
				'***************************************************************************************
				'BUILD THE MASTER ORDER BY CLAUSE HERE
				'****************************************************************************************
				If Not rsCustomerEquipment.EOF Then
				
					EqpOrderByClauseCustom = " ORDER BY CASE ModelIntRecID "
					SortCount = 0
				
					Do While NOT rsCustomerEquipment.EOF
				
						EqpOrderByClauseCustom = EqpOrderByClauseCustom & " WHEN " & rsCustomerEquipment("ModelIntRecID") & " THEN " & Trim(SortCount) & " "
						SortCount = SortCount + 1
				
						rsCustomerEquipment.MoveNext
					Loop
					
					EqpOrderByClauseCustom = EqpOrderByClauseCustom & " END "
				
				End If
				
				'Response.write(EqpOrderByClauseCustom & "<br>")

				SQLCustomerEquipment = "SELECT * FROM EQ_CustomerEquipment INNER JOIN "
				SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Equipment ON EQ_Equipment.InternalRecordIdentifier = EQ_CustomerEquipment.EquipIntRecID "
				SQLCustomerEquipment = SQLCustomerEquipment & " INNER JOIN "
				SQLCustomerEquipment = SQLCustomerEquipment & " EQ_Models ON EQ_Equipment.ModelIntRecID = EQ_Models.InternalRecordIdentifier "
				SQLCustomerEquipment = SQLCustomerEquipment & " WHERE "
				SQLCustomerEquipment = SQLCustomerEquipment & " (EQ_CustomerEquipment.CustID = " & CustIDPassed & ") AND (EQ_Models.ClassIntRecID = " & ClassIntRecID & ") "
				SQLCustomerEquipment = SQLCustomerEquipment & " AND EQ_Models.MightUseAFilter = 1 "
				SQLCustomerEquipment = SQLCustomerEquipment & EqpOrderByClauseCustom 
				
				'Response.write(SQLCustomerEquipment & "<br>")

				Set rsCustomerEquipment = cnnCustomerEquipment.Execute(SQLCustomerEquipment)

				If NOT rsCustomerEquipment.EOF Then
				
					FirstPassOnModel = True
					ModelLoopCounter = 1
					TotalRentalAmount = 0
					TotalPurchaseCost = 0
				
					Do While NOT rsCustomerEquipment.EOF
					
						InstallDate = rsCustomerEquipment("InstallDate")
						StatusCodeIntRecID = rsCustomerEquipment("StatusCodeIntRecID")
						
						SQLEquipStatusCode = "SELECT * FROM EQ_StatusCodes WHERE InternalRecordIdentifier = " & StatusCodeIntRecID
							
						Set cnnEquipStatusCode = Server.CreateObject("ADODB.Connection")
						cnnEquipStatusCode.open (Session("ClientCnnString"))
						Set rsEquipStatusCode = cnnEquipStatusCode.Execute(SQLEquipStatusCode)
						
						If NOT rsEquipStatusCode.EOF Then
							InstallType = rsEquipStatusCode("statusBackendSystemCode")
							InstallTypeFullName = rsEquipStatusCode("statusDesc")
						Else
							InstallType = ""
							InstallTypeFullName = ""
						End If
												
						
						If InstallType = "R" then
						
							RentalFrequencyType = rsCustomerEquipment("RentalFrequencyType")
							
							Select Case RentalFrequencyType
							Case "D"
								RentalFrequencyFullName = "DAYS"
							Case "M"
								RentalFrequencyFullName = "MONTH(S)"
							Case "Y"
								RentalFrequencyFullName = "YEAR(S)"
							End Select
							
							RentalFrequencyNumber = rsCustomerEquipment("RentalFrequencyNumber")
							RentAmt = rsCustomerEquipment("RentAmt")
							
							If RentAmt <> "" Then
								TotalRentalAmount = TotalRentalAmount + RentAmt
								RentAmt = FormatCurrency(RentAmt,0)
							Else
								RentAmt = 0
								RentAmt = FormatCurrency(RentAmt,0)
							End If
							
						Else
							RentalFrequencyFullName = ""
							RentalFrequencyType = ""
							RentalFrequencyNumber = ""
							RentAmt = 0
							RentAmt = FormatCurrency(RentAmt,0)
						End If
												
						SerialNumber = rsCustomerEquipment("SerialNumber")
						PurchaseCost = rsCustomerEquipment("PurchaseCost")
						
						If PurchaseCost <> "" then
							TotalPurchaseCost = TotalPurchaseCost + PurchaseCost
							PurchaseCost = FormatCurrency(PurchaseCost,2)
						End If
						
						ModelIntRecID = rsCustomerEquipment("ModelIntRecID")
						
						If ModelIntRecID <> 0 Then
							BrandName = ModelIntRecID & "-" & GetBrandNameByModelIntRecID(ModelIntRecID)
						Else
							BrandName = ""
						End If
						
						AssetTag1 = rsCustomerEquipment("AssetTag1")
						Description = "DESC NEEDED"
						Description  = GetModelNameByIntRecID(rsCustomerEquipment("ModelIntRecID"))
						
						ModelCount = GetTotalNumberOfModelsForCustomer(CustIDPassed,ModelIntRecID)
						

						%>
					
						<% If cInt(ModelCount) = 1 Then %>
						
							<tr>
								<td>&nbsp;</td>
								<% If BrandName <> "" Then %>
									<td><%= UCASE(BrandName) %>&nbsp;<%= Description %></td>
								<% Else %>
									<td><%= Description %></td>
								<% End If %>
								<td><%= InstallTypeFullName %></td>
								<% If InstallType = "R" Then %>
									<td><%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></td>
									<td><%= RentAmt %></td>
								<% Else %>
									<td>&nbsp;</td>
									<td align="center"><%= RentAmt %></td>
								<% End If %>
								<td align="right"><%= InstallDate %></td>
								<td align="right"><%= FormatCurrency(PurchaseCost,0) %></td>
								<td align="center"><%= SerialNumber %></td>
								<td align="center"><%= AssetTag1 %></td>
							</tr>
							
						<% ElseIf (cInt(ModelCount) > 1) AND (cInt(ModelLoopCounter) <= cInt(ModelCount)) Then %>
						
							<% If FirstPassOnModel = True Then %>
							
								<% ModelLoopCounter = 1 %>
															
								<tr class="accordion-toggle">
									<% If BrandName <> "" Then %>
										<td data-toggle="collapse" data-target=".equip<%= ModelIntRecID %>"><i class="fa fa-plus-circle fa-lg" aria-hidden="true" style="color:#009800"></i></td>
										<td colspan="3"><%= UCASE(BrandName) %>&nbsp;<%= Description %>&nbsp;<span class="equip_qty">(<%= ModelCount %>)</span></td>
										<td align="center"><%= FormatCurrency(GetTotalValueOfRentalModelsForCustomer(CustIDPassed,ModelIntRecID),0) %></td>
										<td>&nbsp;</td>
										<td align="right"><%= FormatCurrency(GetTotalValueOfModelsForCustomer(CustIDPassed,ModelIntRecID),0) %></td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>
									<% Else %>
										<td data-toggle="collapse" data-target=".equip<%= ModelIntRecID %>" colspan="3"><%= Description %>&nbsp;(<%= ModelCount %>)&nbsp;<i class="fa fa-plus-circle" aria-hidden="true" style="color:#009800"></i></td>
										<td align="center"><%= FormatCurrency(GetTotalValueOfRentalModelsForCustomer(CustIDPassed,ModelIntRecID),0) %></td>
										<td>&nbsp;</td>
										<td align="right"><%= FormatCurrency(GetTotalValueOfModelsForCustomer(CustIDPassed,ModelIntRecID),0) %></td>
										<td>&nbsp;</td>
										<td>&nbsp;</td>										
									<% End If %>
								</tr>		  
								<tr class="collapse equip<%= ModelIntRecID %>" style="background-color:#e5ffe5">
									<td>&nbsp;</td>
									<% If BrandName <> "" Then %>
										<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<%= UCASE(BrandName) %>&nbsp;<%= Description %></td>
									<% Else %>
										<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<%= Description %></td>
									<% End If %>
									<td><%= InstallTypeFullName %></td>
									<% If InstallType = "R" Then %>
										<td><%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></td>
										<td align="center"><%= RentAmt %></td>
									<% Else %>
										<td>&nbsp;</td>
										<td align="center"><%= RentAmt %></td>
									<% End If %>
									<td align="right"><%= InstallDate %></td>
									<td align="right"><%= FormatCurrency(PurchaseCost,0) %></td>
									<td align="center"><%= SerialNumber %></td>
									<td align="center"><%= AssetTag1 %></td>			  	
								</tr>
								
								<% FirstPassOnModel = False %>
								<% ModelLoopCounter = ModelLoopCounter + 1 %>
								
							<% Else %>
							
								<tr class="collapse equip<%= ModelIntRecID %>" style="background-color:#e5ffe5">
									<td>&nbsp;</td>
									<% If BrandName <> "" Then %>
										<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<%= UCASE(BrandName) %>&nbsp;<%= Description %></td>
									<% Else %>
										<td style="padding-left:20px;"><i class="fa fa-folder-open-o" aria-hidden="true"></i>&nbsp;<%= Description %></td>
									<% End If %>
									<td><%= InstallTypeFullName %></td>
									<% If InstallType = "R" Then %>
										<td><%= RentalFrequencyNumber %>&nbsp;<%= RentalFrequencyFullName %></td>
										<td align="center"><%= RentAmt %></td>
									<% Else %>
										<td>&nbsp;</td>
										<td align="center"><%= RentAmt %></td>
									<% End If %>
									<td align="right"><%= InstallDate %></td>
									<td align="right"><%= FormatCurrency(PurchaseCost,0) %></td>
									<td align="center"><%= SerialNumber %></td>
									<td align="center"><%= AssetTag1 %></td>
								</tr>
								
								<% 
									ModelLoopCounter = ModelLoopCounter + 1
								
									If cInt(ModelLoopCounter) > cInt(ModelCount) Then
										FirstPassOnModel = True
										ModelLoopCounter = 1
									End If
								 %>
								
							<% End If %>
							
						<% End If %>
						<%
						cnnEquipStatusCode.Close
						rsCustomerEquipment.MoveNext
					
					Loop	
					
				End If
				
				%>
						  	
			</tbody>
			
			<tfoot>
			  <tr>
			  	<td colspan="2">TOTAL</td>
			  	<td>---</td>
			  	<td>---</td>
			  	<td align="center"><%= FormatCurrency(TotalRentalAmount,0) %></td>
			  	<td align="right">---</td>
			  	<td align="right"><%= FormatCurrency(TotalPurchaseCost,0) %></td>
			  	<td align="center">---</td>
			  	<td align="center">---</td>			  	
			  </tr>
			</tfoot>
		</table>
	<%
	
		rsCustomerEquipmentByClass.MoveNext
		Loop
		
		Set rsCustomerEquipment = Nothing
		cnnCustomerEquipment.Close
		Set cnnCustomerEquipment = Nothing
		
	End If

	Set rsEquipStatusCode = Nothing
	'cnnEquipStatusCode.Close
	Set cnnEquipStatusCode = Nothing

	Set rsCustomerEquipmentByClass = Nothing
	cnnCustomerEquipmentByClass.Close
	Set cnnCustomerEquipmentByClass = Nothing
			
			
			%>
		</div>
	
	<%
	rs.close
	cnnCustFiltersList.Close

%>