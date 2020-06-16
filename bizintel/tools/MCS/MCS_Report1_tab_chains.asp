<!-- row !-->
<div class="row" style="width:98%; margin-left:10px; margin-top:20px;">
	<div class="container-fluid">
		<div class="row">
			   <table id="tableSuperSum" class="display compact fold-table" style="width:100%;">
				  <thead>
					  <tr>	
							<th class="td-align1 gen-info-header" colspan="6" style="border-right: 2px solid #555 !important;">General</th>
							<th class="td-align1 vpc-3pavg-header" colspan="7" style="border-right: 2px solid #555 !important;">Sales<small>&nbsp;(Excluding Rent, XSFs, LVF & Category 21)</small></th>
							<th class="td-align1 vpc-lcp-header" colspan="5" style="border-right: 2px solid #555 !important;">MCS</th>
							<th class="td-align1 vpc-misc-header" colspan="3" style="border-right: 2px solid #555 !important;">MISC</th>
							<th class="td-align1 vpc-current-header" colspan="1" style="border-right: 2px solid #555 !important;">Equipment</th>
							<th class="td-align1 activities-header" colspan="2" style="border-right: 2px solid #555 !important;">Activities</th>
					</tr>
					<tr>
						<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important;  border-top: 2px solid #555 !important; width:40px!importnat;" id="salesColumn"><br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>ChainID</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Chain Name</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Primary<br> Slsmn</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Secondary<br> Slsmn</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Install<br> Date</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-3,ReportDate)),1) %><br><%= Year(DateAdd("m",-3,ReportDate))%></th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-2,ReportDate)),1) %><br><%= Year(DateAdd("m",-2,ReportDate))%></th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %><br><%= Year(DateAdd("m",-1,ReportDate))%></th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">3 Prior<br>mos avg $</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Shortage<br>Last 3 mos</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(ReportDate),1) %><br>Sales $</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">GP$<br><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %>&nbsp;<%= Year(DateAdd("m",-1,ReportDate))%></th>
						<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>MCS</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %>&nbsp;MCS<br>Variance</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(ReportDate),1) %>&nbsp;MTD<br>Variance</th>					
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Enrollment<br>Date</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important; border-right: 2px solid #555 !important;" id="salesColumn" >Pending<br>LVF</th>	
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%= MonthName(Month(DateAdd("m",-1,ReportDate))) %><br>Rental $</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">LVF<br><%=MonthName(Month(DateAdd("m",-2,ReportDate)),1) %>&nbsp;<%= Year(DateAdd("m",-3,ReportDate))%></th>									
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Max LVF<br>&nbsp;</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important; border-left: 2px solid #555 !important;" id="salesColumn">Eqp Value</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Action</th>
						<th class="td-align sorttable_numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">Additional<br>Info</th>
					</tr>
				  </thead>		
	
	<%		
			Response.Write("<tbody>")
		ChainSQL1 = "SELECT ChainID FROM AR_Customer INNER JOIN BI_MCSData ON BI_MCSData.CustID = AR_Customer.CustNum WHERE MonthlyContractedSalesDollars <> 0 and ChainID <> 0 group by ChainID"
		'SQL = "SELECT * FROM AR_Customer INNER JOIN BI_MCSData ON BI_MCSData.CustID = AR_Customer.CustNum WHERE MonthlyContractedSalesDollars <> 0 and ChainID <> 0" 	
		
		SQL = "SELECT * FROM AR_Customer INNER JOIN BI_MCSData ON BI_MCSData.CustID = AR_Customer.CustNum WHERE ChainID <> 0 ANd AR_Customer.Custnum In " 
		SQL = SQL & " (SELECT CustID FROM BI_MCSData WHERE ChainID <> 0 )" 
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		
		Set rs1 = Server.CreateObject("ADODB.Recordset")
		rs1.CursorLocation = 3
		Set rs1 = cnn8.Execute(ChainSQL1)
	
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorLocation = 3
		Set rs = cnn8.Execute(SQL)
		child_id = 0
		
		If Not rs1.Eof Then
			Do While Not rs1.EOF
				record_num = 0
				total_Month1Sales_NoRent = 0
				total_Month2Sales_NoRent = 0
				total_Month3Sales_NoRent = 0
				total_ThreePPAvgSales = 0
				total_ShortageHolder = 0
				total_CurrentHolder = 0
				total_Month3GP = 0
				total_CustMonthlyContractedSalesDollars = 0
				total_VarianceHolder = 0
				total_CurrentMonthVarianceHolder = 0
				total_PendingLVFHolder = 0.0
				total_RentalHolder = 0
				total_LVFHolder = 0
				total_TotalEquipmentValue = 0
				total_ChainName = ""
				total_ChainID = rs1("ChainID")
				If Not rs.Eof Then
					Do While Not rs.EOF

						If rs("ChainID") <> total_ChainID Then
						Else
							ShowThisRecord = True
							total_ChainName = rs("ChainName")

							If ShowThisRecord <> False Then			
								SelectedCustomerID = rs("CustNum")
								CustName = rs("Name")
								CustMonthlyContractedSalesDollars = 0							
								CustMonthlyContractedSalesDollars = rs("MonthlyContractedSalesDollars")
								
								If CustMonthlyContractedSalesDollars > 0  Then 
								Else
									CustMonthlyContractedSalesDollars = 0.00
								End If
								MaxMCSCharge = rs("MaxMCSCharge")

								'Decide if this record meets the filter criteria
								If FilterSlsmn1 <> "" And FilterSlsmn1 <> "All" Then
									If CInt(FilterSlsmn1) <> Cint(rs("Salesman")) Then ShowThisRecord = False
								End If

								If FilterSlsmn2 <> "" And FilterSlsmn2 <> "All" Then
									If CInt(FilterSlsmn2) <> Cint(rs("SecondarySalesman")) Then ShowThisRecord = False
								End If
							End If

							Month1Sales_NoRent = rs("Month1Sales_NoRent") - rs("Month1Cat21Sales")
							Month2Sales_NoRent = rs("Month2Sales_NoRent") - rs("Month2Cat21Sales")
							Month3Sales_NoRent = rs("Month3Sales_NoRent") - rs("Month3Cat21Sales")								
							VarianceHolder = Month3Sales_NoRent - CustMonthlyContractedSalesDollars 
							CurrentHolder = rs("CurrentHolder")
							CurrentMonthVarianceHolder = CurrentHolder - CustMonthlyContractedSalesDollars
								
							If ShowThisRecord <> False Then
								record_num = record_num + 1
								ThreePPSales = Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent		
								Month3Cost_NoRent = rs("Month3Cost_NoRent") 	
								Month3GP = Month3Sales_NoRent - Month3Cost_NoRent

								If Not IsNumeric(Month3GP) Then Month3GP  = 0

								ThreePPAvgSales = (Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent) / 3	
								ShortageHolder = ThreePPSales - (CustMonthlyContractedSalesDollars * 3)															
								LVFHolder = rs("LVFHolder") 				
								LVFHolderCurrent = rs("LVFHolderCurrent") 					
								TotalEquipmentValue = rs("TotalEquipmentValue")			
								TotalCustsReported = TotalCustsReported + 1										
								total_Month1Sales_NoRent = total_Month1Sales_NoRent + Month1Sales_NoRent'first month
								total_Month2Sales_NoRent = total_Month2Sales_NoRent + Month2Sales_NoRent'second month
								total_Month3Sales_NoRent = total_Month3Sales_NoRent + Month3Sales_NoRent'third month			
								total_ThreePPAvgSales = total_ThreePPAvgSales + ThreePPAvgSales'average
								total_ShortageHolder = total_ShortageHolder + ShortageHolder
								total_CurrentHolder = total_CurrentHolder + CurrentHolder
								total_Month3GP = total_Month3GP + Month3GP
								total_CustMonthlyContractedSalesDollars = total_CustMonthlyContractedSalesDollars + CustMonthlyContractedSalesDollars
								total_VarianceHolder = total_VarianceHolder + VarianceHolder
								total_CurrentMonthVarianceHolder = total_CurrentMonthVarianceHolder + CurrentMonthVarianceHolder
								PendingLVFHolder = rs("PendingLVF")
								total_PendingLVFHolder = total_PendingLVFHolder + PendingLVFHolder		
								RentalHolder = rs("RentalHolder")

								IF rs("Month3XSF") > 0 Then RentalHolder = RentalHolder  + rs("Month3XSF")
								
								total_RentalHolder = total_RentalHolder + RentalHolder
								total_LVFHolder = total_LVFHolder + LVFHolder	
																				
								If TotalEquipmentValue > 0 Then	 total_TotalEquipmentValue = total_TotalEquipmentValue + TotalEquipmentValue
							End If	
						End If			
						
						rs.movenext
							
					Loop
	
					rs.movefirst
					
				Else
					Response.Write("Nothing To Report")
				End If
	
				If record_num > 0 Then
					If ShowAllCusts <> 1 Then
						If total_Month3Sales_NoRent >= total_CustMonthlyContractedSalesDollars Then ShowThisRecord = False
					End If
	
					If ShowZeroSalesCusts = 1 Then
						If total_Month3Sales_NoRent > 0 Then ShowThisRecord = False
					End If

					' Calc under by the current month recovered the deficit
					If total_VarianceHolder < 0 Then 'Meaning they have a variance
						If total_CurrentHolder >= total_CustMonthlyContractedSalesDollars + ABS(total_VarianceHolder)  Then
							If IncludeDeficitCovered <> 1 Then ShowThisRecord = False
						End If
					End If

					If ABS(total_VarianceHolder) < 100 Then
						If total_Month3Sales_NoRent <> 0 Then
							total_VariancePercentHolder = 100 - ((total_Month3Sales_NoRent/total_CustMonthlyContractedSalesDollars) * 100) 
						End If
						total_VariancePercentHolder  = total_VariancePercentHolder  * -1
						If ApplyRule = 1 Then
							If ABS(total_VariancePercentHolder) < 10 Then
								ShowThisRecord = False
							End If
						End If
					End If
	
					If ShowThisRecord <> False Then					
						total_ThreePPSales = total_Month1Sales_NoRent + total_Month2Sales_NoRent + total_Month3Sales_NoRent		
						If Not IsNumeric(total_Month3GP) Then total_Month3GP  = 0
						total_ThreePPAvgSales = (total_Month1Sales_NoRent + total_Month2Sales_NoRent + total_Month3Sales_NoRent) / 3	
						If total_ThreePPAvgSales >= total_CustMonthlyContractedSalesDollars Then ShowThisRecord = False  '69 to 56
						If total_CustMonthlyContractedSalesDollars - total_ThreePPAvgSales < (100 * record_num) Then
							x = total_CustMonthlyContractedSalesDollars - total_ThreePPAvgSales
							If (x / total_CustMonthlyContractedSalesDollars) * 100 < 10 Then ShowThisRecord = False ' 56 to 50
						End If
					End If

					If ShowThisRecord <> False Then
						child_id = child_id + 1 ' child table id.
						
						
						' Chain information
						Response.Write("<tr id=""CUST" & total_ChainID & """")
						Response.Write("class = 'view'>")
						Response.Write("<td class='details-control'></td>") ' chain filter icon
						Response.Write("<td class='smaller-detail-line'>" & total_ChainID & "</td>") ' Chain ID
						Response.Write("<td class='smaller-detail-line'>" & total_ChainName & "</td>") ' Chain Name
						Response.Write("<td align='center' class='smaller-detail-line'> --- </td>") ' primary slsmn
						Response.Write("<td align='center' class='smaller-detail-line'> --- </td>") ' secondary slsmn
						Response.Write("<td align='center' class='smaller-detail-line'> --- </td>") ' install date
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(total_Month1Sales_NoRent,0) & "</td>") ' last third month
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(total_Month2Sales_NoRent,0) & "</td>") ' last second month
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(total_Month3Sales_NoRent,0) & "</td>") ' last first month
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency((total_Month1Sales_NoRent + total_Month2Sales_NoRent + total_Month3Sales_NoRent)/3, 0) & "</td>") ' 3 month average
						
						If Not IsNumeric(total_ShortageHolder) Then total_ShortageHolder = 0
						If total_ShortageHolder < 0 Then
							Response.Write("<td align='right' class='negative-thin smaller-detail-line'>" & FormatCurrency(total_ShortageHolder ,0,0,0) & "</td>") ' shortage last 3 month
						Else
							Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(total_ShortageHolder, 0,0,0) & "</td>")
						End If

						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(total_CurrentHolder, 0) & "</td>") ' this month sales
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(total_Month3GP, 0) & "</td>") ' GP
						Response.Write("<td align='right' style='border-left: 2px solid rgb(85, 85, 85) !important;' class='smaller-detail-line'>" & FormatCurrency(total_CustMonthlyContractedSalesDollars, 0) & "</td>") ' MCS
						
						If total_VarianceHolder < 1 Then 
							If ABS(total_VarianceHolder) < 1 Then
								Response.Write("<td align='right' class='negative-thin not-as-small-detail-line'>" & "-$0"  & "</td>") ' handle variance less than 1 whole dollar
							Else
								Response.Write("<td align='right' class='negative-thin not-as-small-detail-line'>" & FormatCurrency(total_VarianceHolder ,0,0,0) & "</td>")
							End If
						Else
							Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(total_VarianceHolder, 0, 0, 0) & "</td>") ' last montth MCS variance
						End If

						If total_CurrentMonthVarianceHolder < 1 Then 
							If ABS(total_CurrentMonthVarianceHolder) < 1 Then
								Response.Write("<td align='right' class='negative-thin not-as-small-detail-line'>" & "-$0"  & "</td>") ' handle variance less than 1 whole dollar
							Else
								Response.Write("<td align='right' class='negative-thin not-as-small-detail-line'>" & FormatCurrency(total_CurrentMonthVarianceHolder ,0,0,0) & "</td>")
							End If
						Else
							Response.Write("<td align='right' class='not-as-small-detail-line'>" & FormatCurrency(total_CurrentMonthVarianceHolder ,0,0,0) & "</td>")
						End If

						Response.Write("<td align='center' class='smaller-detail-line'> --- </td>") ' Enrollment Date
						Response.Write("<td align='right' style='border-right: 2px solid rgb(85, 85, 85) !important;' class='smaller-detail-line'>" & FormatCurrency(total_PendingLVFHolder, 2) & "</td>") ' Pending LVF
						
						If total_RentalHolder < 0 Then
							Response.Write("<td align='right' class='negative-thin smaller-detail-line'>" & FormatCurrency(total_RentalHolder ,0) & "</td>")
						Else
							Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(total_RentalHolder ,0) & "</td>")				
						End If

						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(total_LVFHolder, 2) & "</td>") ' LVF
						Response.Write("<td align='center' class='smaller-detail-line'> --- </td>") ' Max LVF
						Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(total_TotalEquipmentValue, 0) & "</td>") ' Equip value
						Response.Write("<td align='center' class='smaller-detail-line'> --- </td>") ' acction
						Response.Write("<td align='center' class='smaller-detail-line'> --- </td>") ' additional infor
						Response.Write("</tr>")
						'Customers informations of each Chain
						Response.Write("<tr class = 'fold'>")
						Response.Write("<td colspan = '24' style = 'padding-left: 0px; padding-right: 0px;'>")
						'Response.Write("<div class='fold-content'>")
						'Response.Write("<h3> Chain Information </h3>")
						'Response.Write("<p> This table shows customer's information of Chain")
						'Response.Write("</div>")
						Response.Write("<table id=""tableSuperSum_child" & child_id & """")
						Response.Write("class='display1' style = 'width:100%;border-bottom: 1px solid #111;'>")
						' child table
				%>
						<thead>				
							<tr>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-left: 2px solid #555 !important; border-top: 2px solid #555 !important; width:55px;" id="salesColumn"><br></th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:67px;" id="salesColumn"><br>Acct</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important; width:64px;" id="salesColumn"><br>Client</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:66px;" id="salesColumn">Primary<br> Slsmn</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:85px;" id="salesColumn">Secondary<br> Slsmn</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:54px;" id="salesColumn">Install<br> Date</th>
	
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-left: 2px solid #555 !important;width:47px; border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-3,ReportDate)),1) %><br><%= Year(DateAdd("m",-3,ReportDate))%></th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:45px;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-2,ReportDate)),1) %><br><%= Year(DateAdd("m",-2,ReportDate))%></th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:45px;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %><br><%= Year(DateAdd("m",-1,ReportDate))%></th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important; width: 47px;" id="salesColumn">3 Prior<br>mos avg $</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important; width: 74px;" id="salesColumn">Shortage<br>Last 3 mos</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:51px;" id="salesColumn"><%=MonthName(Month(ReportDate),1) %><br>Sales $</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:73px;" id="salesColumn">GP$<br><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %>&nbsp;<%= Year(DateAdd("m",-1,ReportDate))%></th>
	
								
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;width:51px;" id="salesColumn"><br>MCS</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:73px;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %>&nbsp;MCS<br>Variance</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:73px;" id="salesColumn"><%=MonthName(Month(ReportDate),1) %>&nbsp;MTD<br>Variance</th>					
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:86px;" id="salesColumn">Enrollment<br>Date</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width: 72px; border-right: 2px solid #555 !important;" id="salesColumn" >Pending<br>LVF</th>	
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:67px;" id="salesColumn"><%= MonthName(Month(DateAdd("m",-1,ReportDate))) %><br>Rental $</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:74px;" id="salesColumn">LVF<br><%=MonthName(Month(DateAdd("m",-2,ReportDate)),1) %>&nbsp;<%= Year(DateAdd("m",-3,ReportDate))%></th>
												
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:42px;" id="salesColumn">Max LVF<br>&nbsp;</th>
								
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-right: 2px solid #555 !important; border-top: 2px solid #555 !important; border-left: 2px solid #555 !important;width:104px;" id="salesColumn">Eqp Value</th>
								
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-top: 2px solid #555 !important;width:57px;" id="salesColumn"><br>Action</th>
								<th class="td-align sorttable_numeric smaller-header child-head" style="display:none; border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;width:85px;" id="salesColumn">Additional<br>Info</th>
								
							</tr>
						</thead>
	
				<%
						
						Response.Write("<tbody>")

						If Not rs.Eof Then
							Do While Not rs.EOF
								If rs("ChainID") <> total_ChainID Then
								Else 
									ShowThisRecord = True
										
									If ShowThisRecord <> False Then			
										PrimarySalesMan =  ""
										SecondarySalesMan =  ""
										SelectedCustomerID = rs("CustNum")
										CustName = rs("Name")
										CustMonthlyContractedSalesDollars = 0
										InstallDate = ""
										EnrollmentDate = ""
										PrimarySalesMan = rs("Salesman")
										SecondarySalesMan = rs("SecondarySalesman")
										CustMonthlyContractedSalesDollars = rs("MonthlyContractedSalesDollars")
										If CustMonthlyContractedSalesDollars > 0  Then 
										Else
											CustMonthlyContractedSalesDollars = 0.00
										End If
										InstallDate = rs("InstallDate")
										MaxMCSCharge = rs("MaxMCSCharge")
										EnrollmentDate =  rs("MCSEnrollmentDate")
										If Len(EnrollmentDate) < 2 Then EnrollmentDate = ""

										'Decide if this record meets the filter criteria
										If FilterSlsmn1 <> "" And FilterSlsmn1 <> "All" Then
											If CInt(FilterSlsmn1) <> Cint(rs("Salesman")) Then ShowThisRecord = False
										End If
										If FilterSlsmn2 <> "" And FilterSlsmn2 <> "All" Then
											If CInt(FilterSlsmn2) <> Cint(rs("SecondarySalesman")) Then ShowThisRecord = False
										End If
									End If
	
									Month3Sales_NoRent = rs("Month3Sales_NoRent") - rs("Month3Cat21Sales") 
									VarianceHolder = Month3Sales_NoRent - CustMonthlyContractedSalesDollars 
									CurrentHolder = rs("CurrentHolder")
									CurrentMonthVarianceHolder = CurrentHolder - CustMonthlyContractedSalesDollars				
	
									If ShowThisRecord <> False Then									
										Month1Sales_NoRent = rs("Month1Sales_NoRent") - rs("Month1Cat21Sales") 
										Month2Sales_NoRent = rs("Month2Sales_NoRent") - rs("Month2Cat21Sales") 										
										ThreePPSales = Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent										
										Month3Cost_NoRent = rs("Month3Cost_NoRent") 										
										Month3GP = Month3Sales_NoRent - Month3Cost_NoRent

										If Not IsNumeric(Month3GP) Then Month3GP  = 0	

										ThreePPAvgSales = (Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent) / 3										
										ShortageHolder = ThreePPSales - (CustMonthlyContractedSalesDollars * 3)																				
										LVFHolder = rs("LVFHolder") 										
										LVFHolderCurrent = rs("LVFHolderCurrent") 																				
										TotalEquipmentValue = rs("TotalEquipmentValue")										
										TotalCustsReported = TotalCustsReported + 1	
										Response.Write("<tr id=""CUST" & SelectedCustomerID & """")
										Response.Write(">")
										Response.Write("<td class='smaller-detail-line' style='width:48px;'></td>")
										Response.Write("<td class='smaller-detail-line' style='width:67px;'><a href='../CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>"& SelectedCustomerID  & "</a></td>")
										Response.Write("<td class='smaller-detail-line' style = 'width:64px;'><a href='../CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>"& CustName & "</a></td>")	
										PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(PrimarySalesMan)
										SecondarySalesPerson = GetSalesmanNameBySlsmnSequence(SecondarySalesman)

										If Instr(PrimarySalesPerson ," ") <> 0 Then
											Response.Write("<td class='smaller-detail-line' style='width:65px;'>" & Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")+1) & "</td>")
										Else
											Response.Write("<td class='smaller-detail-line' style='width:65px;'>" & PrimarySalesPerson & "</td>")
										End If

										If Instr(SecondarySalesPerson," ") <> 0 Then
											Response.Write("<td class='smaller-detail-line' style='width:83px;'>" & Left(SecondarySalesPerson,Instr(SecondarySalesPerson," ")+1) & "</td>")
										Else
											Response.Write("<td class='smaller-detail-line' style='width:83px;'>" & SecondarySalesPerson & "</td>")
										End If
										
										If Not IsDate(InstallDate) Then InstallDate = "01/01/2000"

										InstallDate = cDate(InstallDate) 
										iYear = Year(InstallDate)

										If Month(InstallDate) < 10 Then iMonth = "0" & Month(InstallDate) else iMonth = Month(InstallDate)

										If Day(InstallDate) < 10 Then iDay = "0" & Day(InstallDate) else iDay = Day(InstallDate)

										Response.Write("<td align='right' class='smaller-detail-line' style='width:54px;'><span class='hidden'>" & iYear & iMonth & iDay & "</span>" & Left(InstallDate,Len(InstallDate)-4) & Right(InstallDate,2) & "</td>")	
										MissedMonth1 = False : MissedMonth2 = False : MissedMonth3 = False

										If Month3Sales_NoRent - CustMonthlyContractedSalesDollars < 1 Then MissedMonth3 = True

										If Month2Sales_NoRent - CustMonthlyContractedSalesDollars < 1 Then MissedMonth2 = True

										If Month1Sales_NoRent - CustMonthlyContractedSalesDollars < 1 Then MissedMonth1 = True
										
										If MissedMonth3  = True AND MissedMonth2  = True AND MissedMonth1  = True Then
											Response.Write("<td align='right' class='smaller-detail-line' style='width:47px;'><mark>" & FormatCurrency(Month1Sales_NoRent,0) & "</mark></td>")
										Else
											Response.Write("<td align='right' class='smaller-detail-line' style='width:47px;'>" & FormatCurrency(Month1Sales_NoRent,0) & "</td>")				
										End If
										
										If MissedMonth3  = True AND MissedMonth2  = True Then
											Response.Write("<td align='right' class='smaller-detail-line' style='width:45px;'><mark>" & FormatCurrency(Month2Sales_NoRent,0) & "</mark></td>")
											Response.Write("<td align='right' class='smaller-detail-line' style='width:45px;'><mark>" & FormatCurrency(Month3Sales_NoRent,0) & "</mark></td>")
										Else
											Response.Write("<td align='right' class='smaller-detail-line' style='width:45px;'>" & FormatCurrency(Month2Sales_NoRent,0) & "</td>")
											Response.Write("<td align='right' class='smaller-detail-line' style='width:45px;'>" & FormatCurrency(Month3Sales_NoRent,0) & "</td>")
										End If										
	
										Response.Write("<td align='right' class='smaller-detail-line' style='width: 47px;'>" & FormatCurrency(ThreePPAvgSales,0) & "</td>")
										
										If ShortageHolder < 0 Then
											Response.Write("<td align='right' class='negative-thin smaller-detail-line' style='width: 74px;'> " & FormatCurrency(ShortageHolder ,0,0,0) & "</td>")
										Else
											Response.Write("<td align='right' class='smaller-detail-line' style='width: 74px;'>" & FormatCurrency(ShortageHolder ,0,0,0) & "</td>")				    
										End If																				
										
										' Calc under by the current month recovered the deficit
										If VarianceHolder < 0 Then 'Meaning they have a variance
											If CurrentHolder >= CustMonthlyContractedSalesDollars + ABS(VarianceHolder)  Then
												Response.Write("<td align='right' class='smaller-detail-line' style='width:51px;'><font color='blue'><b>" & FormatCurrency(CurrentHolder,0)  & "</b></foont></td>")
											Else
												Response.Write("<td align='right' class='smaller-detail-line' style='width:51px;'><font color='black'>" & FormatCurrency(CurrentHolder,0)  & "</foont></td>")
											End If
										Else
											Response.Write("<td align='right' class='smaller-detail-line' style='width:51px;'><font color='black'>" & FormatCurrency(CurrentHolder,0) & "</foont></td>")
										End If																			
										
										Response.Write("<td align='right' class='smaller-detail-line' style='width:71px;'>" &  FormatCurrency(Month3GP,0)  & "</td>")										
										Response.Write("<td align='right' class='smaller-detail-line' style='border-left: 2px solid #555 !important; width:49px;'>" & FormatCurrency(CustMonthlyContractedSalesDollars,0) & "</td>")
	
										If VarianceHolder < 1 Then 
											If ABS(VarianceHolder) < 1 Then
												Response.Write("<td align='right' class='negative-thin not-as-small-detail-line' style='width:72px;'>" & "-$0"  & "</td>") ' handle variance less than 1 whole dollar
											Else
												Response.Write("<td align='right' class='negative-thin not-as-small-detail-line' style='width:72px;'>" & FormatCurrency(VarianceHolder ,0,0,0) & "</td>")
											End If
										Else
											Response.Write("<td align='right' class='not-as-small-detail-line' style='width:72px;'>" & FormatCurrency(VarianceHolder ,0,0,0) & "</td>")
										End If
	
										If CurrentMonthVarianceHolder < 1 Then 
											If ABS(CurrentMonthVarianceHolder) < 1 Then
												Response.Write("<td align='right' class='negative-thin not-as-small-detail-line' style='width:72px;'>" & "-$0"  & "</td>") ' handle variance less than 1 whole dollar
											Else
												Response.Write("<td align='right' class='negative-thin not-as-small-detail-line' style='width:72px;'>" & FormatCurrency(CurrentMonthVarianceHolder ,0,0,0) & "</td>")
											End If
										Else
											Response.Write("<td align='right' class='not-as-small-detail-line' style='width:72px;'>" & FormatCurrency(CurrentMonthVarianceHolder ,0,0,0) & "</td>")
										End If																				

										If EnrollmentDate <> "" Then
											'EnrollmentDate Date
											EnrollmentDate = cDate(EnrollmentDate) 
											eYear = Year(EnrollmentDate)

											If Month(EnrollmentDate) < 10 Then eMonth = "0" & Month(EnrollmentDate) else eMonth = Month(EnrollmentDate)

											If Day(EnrollmentDate) < 10 Then eDay = "0" & Day(EnrollmentDate) else eDay = Day(EnrollmentDate)

											EnrollmentDispayableDate = eMonth & "/" & eDay  & "/" & eYear
											EnrollmentDispayableDate  = cDate(EnrollmentDispayableDate) 
											Response.Write("<td align='right' class='smaller-detail-line' style='width:85px;'><span class='hidden'>" & eYear & eMonth & eDay & "</span>" & Left(EnrollmentDispayableDate,Len(EnrollmentDispayableDate)-4) & Right(EnrollmentDispayableDate,2) & "</td>")	
										Else
											Response.Write("<td align='right' class='smaller-detail-line' style='width:85px;'><span class='hidden'> --- </td>")
										End If
										PendingLVFHolder = rs("PendingLVF")
										Response.Write("<td align='right' class='smaller-detail-line' style='border-right: 2px solid #555 !important; width:70px;'>" &  FormatCurrency(PendingLVFHolder,2)  & "</td>")
										RentalHolder = rs("RentalHolder")

										IF rs("Month3XSF") > 0 Then RentalHolder = RentalHolder  + rs("Month3XSF")
										
										If RentalHolder < 0 Then
											Response.Write("<td align='right' class='negative-thin smaller-detail-line' style='width:67px;'>" & FormatCurrency(RentalHolder ,0) & "</td>")
										Else
											Response.Write("<td align='right' class='smaller-detail-line' style='width:67px;'>" & FormatCurrency(RentalHolder ,0) & "</td>")				
										End If
										
										Response.Write("<td align='right' class='smaller-detail-line' style='width:72px;'>" &  FormatCurrency(LVFHolder,2)  & "</td>")					
										MaxLVFPerMachineHolder = MaxMCSCharge

										If Not IsNumeric(MaxMCSCharge) Then 
											Response.Write("<td align='right' class='smaller-detail-line' style='width:42px;'>" & MaxLVFPerMachineHolder & "</td>")
										Else
											Response.Write("<td align='right' class='smaller-detail-line' style='width:42px;'>" & FormatCurrency(MaxLVFPerMachineHolder,2) & "</td>")				
										End If																				
										
										If TotalEquipmentValue > 0 Then											
											LCPGP = 0 											
											If TotalEquipmentValue <> 0 Then %>
												<td align="right" class="smaller-detail-line" style="width:104px;" >
												<a data-toggle="modal" data-show="true" href="#" data-cust-id="<%= SelectedCustomerID %>" data-lcp-gp="<%= LCPGP %>" data-target="#modalEquipmentVPC" data-tooltip="true" data-title="View Customer Equipment" ><%= FormatCurrency(TotalEquipmentValue,0) %></a>    
												</td>
											<% Else %>
												<%= FormatCurrency(TotalEquipmentValue,0) %>
											<% End If %>											
										<%
										Else
											Response.Write("<td align='right' class='smaller-detail-line' style='width:104px;'>No Equipment</td>")
										End If
	
										'Action
										Response.Write("<td align='right' class='smaller-detail-line' style='width:57px;'>")
										btncolor = "btn-success"
										
										if GetMCSNotesStatus(SelectedCustomerID, MonthName(Month(DateAdd("m",-1,ReportDate)))) Then 
											if GetMCSNotesNoActionStatus(SelectedCustomerID, MonthName(Month(DateAdd("m",-1,ReportDate)))) = 2 Then
												btncolor = "btn-default noaction"
											Else
												btncolor = "btn-default"
											End If					
										End if

										Response.Write "<button type=""button"" class=""" & btncolor & """ id=""btn" & SelectedCustomerID & """ data-toggle=""modal"" data-target=""#modalGeneralNotesGroupM"" data-cust-id=""" & SelectedCustomerID & """ data-cust-name=""" &CustName & """ data-mcs-variance=""" & VarianceHolder & """ data-mcs-salespersonid1=""" & PrimarySalesMan & """ data-mcs-salespersonid2=""" & SecondarySalesMan & """  data-mcs-salesperson1=""" & PrimarySalesPerson & """ data-mcs-salesperson2=""" & SecondarySalesPerson & """ data-mcs-month=""" & MonthName(Month(DateAdd("m",-1,ReportDate))) & """ data-mcs-userno=""" & Session("userNo") & """ data-maxmcscharge=""" & MaxMCSCharge & """ data-mcsdollars=""" & CustMonthlyContractedSalesDollars & """ >Action</button>"
										Response.Write("</td>")										
										'Additional Info / Notes										
										'Allow for a note here as a way to put in a note for the customer in general
										'Use -2 as the category number for MCS notes
										
										If UserHasAnyUnviewedNotes(SelectedCustomerID) Then
											'Pulsing icon
											Response.Write("<td align='center' class='smaller-detail-line' style='width:85px;'>")
											Response.Write("<a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='" & SelectedCustomerID & "' class='ole' rel='tooltip' style='cursor:pointer;'><i class='fa fa-file-text-o faa-pulse animated fa-2x' aria-hidden='true'></i></a>")																	
											Response.Write("</td>")
										Else
											'Regular icon
											Response.Write("<td align='center' class='smaller-detail-line' style='width:85px;'>")
											Response.Write("<a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='" & SelectedCustomerID  & "' class='ole' rel='tooltip' style='cursor:pointer;'><i class='fa fa-file-text-o' aria-hidden='true'></i></a>")											
											Response.Write("</td>")
										End If	
	
										Response.Write("</tr>")
										
									End If
								End If	

								rs.movenext
									
							Loop
							rs.movefirst
							Response.Write("</tbody>")
							Response.Write("</table>")
						Else	
							Response.Write("Nothing To Report")
						End If
	
						Response.Write("</td>")
						Response.Write("<td style='display: none'>" & total_ChainID & "</td>") ' Chain ID
						Response.Write("<td style='display: none'>" & total_ChainName & "</td>") ' Chain Name
						Response.Write("<td style='display: none'></td>")
						Response.Write("<td style='display: none'></td>")
						Response.Write("<td style='display: none'></td>")
						Response.Write("<td style='display: none'>" & FormatCurrency(total_Month1Sales_NoRent,0) & "</td>") ' last third month
						Response.Write("<td style='display: none'>" & FormatCurrency(total_Month2Sales_NoRent,0) & "</td>") ' last second month
						Response.Write("<td style='display: none'>" & FormatCurrency(total_Month3Sales_NoRent,0) & "</td>") ' last first month
						Response.Write("<td style='display: none'>" & FormatCurrency((total_Month1Sales_NoRent + total_Month2Sales_NoRent + total_Month3Sales_NoRent)/3, 0) & "</td>") ' 3 month average
	
						If total_ShortageHolder < 0 Then
							Response.Write("<td style='display: none'>" & FormatCurrency(total_ShortageHolder, 0,0,0) & "</td>") ' shortage last 3 month
						Else
							Response.Write("<td style='display: none'>" & FormatCurrency(total_ShortageHolder, 0,0,0) & "</td>")
						End If

						'Response.Write("<td style='display: none'></td>")
						Response.Write("<td style='display: none'>" & FormatCurrency(total_CurrentHolder, 0) & "</td>") ' this month sales
						Response.Write("<td style='display: none'>" & FormatCurrency(total_Month3GP, 0) & "</td>") ' GP
						Response.Write("<td style='display: none'>" & FormatCurrency(total_CustMonthlyContractedSalesDollars, 0) & "</td>") ' MCS

						If total_VarianceHolder < 1 Then 
							If ABS(total_VarianceHolder) < 1 Then
								Response.Write("<td style='display: none'>" & "-$0"  & "</td>") ' handle variance less than 1 whole dollar
							Else
								Response.Write("<td style='display: none'>" & FormatCurrency(total_VarianceHolder ,0,0,0) & "</td>")
							End If
						Else
							Response.Write("<td style='display: none'>" & FormatCurrency(total_VarianceHolder, 0, 0, 0) & "</td>") ' last montth MCS variance
						End If

						If total_CurrentMonthVarianceHolder < 1 Then 
							If ABS(total_CurrentMonthVarianceHolder) < 1 Then
								Response.Write("<td style='display: none'>" & "-$0"  & "</td>") ' handle variance less than 1 whole dollar
							Else
								Response.Write("<td style='display: none'>" & FormatCurrency(total_CurrentMonthVarianceHolder ,0,0,0) & "</td>")
							End If
						Else
							Response.Write("<td style='display: none'>" & FormatCurrency(total_CurrentMonthVarianceHolder ,0,0,0) & "</td>")
						End If

						Response.Write("<td style='display: none'></td>")
						Response.Write("<td style='display: none'>" & FormatCurrency(total_PendingLVFHolder, 2) & "</td>") ' Pending LVF

						If total_RentalHolder < 0 Then
							Response.Write("<td style='display: none'>" & FormatCurrency(total_RentalHolder ,0) & "</td>")
						Else
							Response.Write("<td style='display: none'>" & FormatCurrency(total_RentalHolder ,0) & "</td>")				
						End If

						Response.Write("<td style='display: none'>" & FormatCurrency(total_LVFHolder, 2) & "</td>") ' LVF
						Response.Write("<td style='display: none'></td>")
						Response.Write("<td style='display: none'>" & FormatCurrency(total_TotalEquipmentValue, 0) & "</td>") ' Equip value
						Response.Write("<td style='display: none'></td>")
						Response.Write("<td style='display: none'></td>")
						Response.Write("</tr>")
					End If	
				Else	
				End If

				rs1.movenext

			Loop
	
			Response.Write("</tbody>")
			Response.Write("</table>")		
			Response.Write("</div>")

		Else
			Response.Write("Nothing To Report")
		End If	
	
	%>				
				 </table> 
			</div>
		</div>
	<!-- eof responsive tables !-->	
	
	<!-- eof row !-->
	
	<!-- row !-->
	<div class="row">
	<%
	   Response.Write("<input type='hidden' id='number_entire' style='display:none;' value="& child_id & " />")
	'Response.Write("<div class='col-lg-12'><h3>" & "Total Customers Listed:" & TotalCustsReported  & "</h3></div>")
	%>
	<div class="col-lg-12"><hr></div>
	</div>
	<!-- eof row !-->
	