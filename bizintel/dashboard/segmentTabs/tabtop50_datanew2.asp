
<div class="row">

<div class="tab-content">

<table id="tableDataTopNew" class="display compact" style="width:100%;">

					<thead>
						<tr>	
							<th colspan="2" class="sorttable numeric smaller-header"></th>
							<th class="td-align1 vpc-variance-header" colspan="4" style="border-right: 2px solid #555 !important;">Variances</th>
							<th class="td-align1 vpc-3pavg-header" colspan="7" style="border-right: 2px solid #555 !important;">Sales</th>
							<th class="td-align1 vpc-lcp-header" colspan="5" style="border-right: 2px solid #555 !important;">MCS</th>
							<th class="td-align1 vpc-current-header" colspan="3" style="border-right: 2px solid #555 !important;">EQUIP ROI</th>
							<th class="td-align1 gen-info-header" colspan="4" style="border-right: 2px solid #555 !important;">General</th>
							
						</tr>

							<% '
							'Setup PP1 & PP2 descriptions
							
							PP1Var = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated -1) & "<br>" & GetPeriodYearBySeq(PeriodSeqBeingEvaluated-1) & "&nbsp;$"
							PP1VarShort = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated -1)
							PP2Var = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated -2) & "<br>" & GetPeriodYearBySeq(PeriodSeqBeingEvaluated-2) & "&nbsp;$"
							PP2VarShort = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated -2)
							PVarShort = "Period&nbsp;" & GetPeriodBySeq(PeriodSeqBeingEvaluated)
							%>
						
						<tr>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Acct</th> 
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Client</th> 
							<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><%= PVarShort %> vs<br>3P avg $</th> 
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Day<br>Impact</th>  
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>ADS</th> 
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%= PVarShort %> vs<br>12P avg $</th>
							
							
							<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br><%= PP1VarShort %></th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br><%= PP2VarShort %></th>							
							
							<th class="td-align sorttable_numeric smaller-header not-as-small-detail-line" style="border-left: 2px solid #555 !important; border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><%= PVarShort %> $</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>3P avg $</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>12P avg $</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Current $</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>SPLY $</th> 
							<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>MCS</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%= PVarShort %> <br>vs MCS</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">3P avg vs<br> MCS</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">12P avg vs<br> MCS</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Current vs<br> MCS</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><%= PVarShort %><br>ROI</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">3P avg<br>ROI</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">Equipment<br>Value</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Primary<br> Slsmn</th>
							<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Customer<br>Type</th>
							<th class="td-align smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Customer<br>Notes</th>
							<th class="td-align smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;">Rules</th>
						</tr>
					</thead>

					
<tbody>
<%
	Segment = Request.QueryString("p")

	ShowPercentageColumns = False

	Select Case MUV_READ("LOHVAR")
		Case "Secondary"
	
			SQL = "SELECT TOP 50 * FROM BI_DashboardSegmentTabs WHERE Tab = 'ALL' AND SecondarySalesmanNumber = " & Segment & " ORDER BY TwelvePAvgSales DESC"
			
		Case "Primary"
	
			SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
			SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
			SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
			SQL = SQL & " FROM CustCatPeriodSales_ReportData "
			SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
			SQL = SQL & " AND PrimarySalesman = " & Segment 
			
		Case "CustType"
	
			SQL = "SELECT Distinct CustCatPeriodSales_ReportData.CustNum,LCPTotSalesAllCats as LCPSales, Total3PPAvgAllCats, TotalCostAllCats, TotalTPLYAllCats "
			SQL = SQL & ",Total3PPSalesAllCats AS ThreePPSales "
			SQL = SQL & ", Total12PPSalesAllCats As TwelvePPSales "
			SQL = SQL & " FROM CustCatPeriodSales_ReportData "
			SQL = SQL & " WHERE ThisPeriodSequenceNumber = " & PeriodSeqBeingEvaluated 
			SQL = SQL & " AND CustType = " & CustType 
			
	End Select	
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.ConnectionTimeout = 120
	cnn8.open (Session("ClientCnnString"))
	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	Set rs = cnn8.Execute(SQL)
	

		GrandTotLCPvs3PAvgSales = 0
				
		Do While Not rs.EOF

			Response.Write("<tr>")
		
			
			Response.Write("<td><a href='../tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & rs("CustID") & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>" & rs("CustID") & "</a></td>")
			
			Response.Write("<td><a href='../tools/CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & rs("CustID") & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>" & rs("CustName") & "</a></td>")
			
			Response.Write("<td>" & FormatCurrency(rs("LCPv3PAvg"),0,-2,0) & "</td>")
	
			Response.Write("<td>" & FormatCurrency(rs("DayImpact"),0) & "</td>")
			
			Response.Write("<td>" & FormatCurrency(rs("ADS"),0) & "</td>")
			
			Response.Write("<td>" & FormatCurrency(rs("LCPv12PAvg"),0) & "</td>")
			
			Response.Write("<td>" & FormatCurrency(rs("PP1Sales"),0,-2,0)  & "</td>")
			
			Response.Write("<td>" & FormatCurrency(rs("PP2Sales"),0,-2,0)  & "</td>")
			
			Response.Write("<td>" & FormatCurrency(rs("LCPSales"),0,-2,0)  & "</td>")
			
			Response.Write("<td>" & FormatCurrency(rs("ThreePAvgSales"),0,-2,0)  & "</td>")
			
			Response.Write("<td>" & FormatCurrency(rs("TwelvePAvgSales"),0,-2,0)  & "</td>")
			
			Response.Write("<td>"& FormatCurrency(rs("CPSales"),0,-2,0)  & "</td>")
			
			Response.Write("<td>" & FormatCurrency(rs("SPLYSales"),0,-2,0)  & "</td>")
			
			If Not IsNull(rs("MCS")) Then
				Response.Write("<td>" &  FormatCurrency(rs("MCS"),0)  & "</td>")
			Else
				Response.Write("<td>&nbsp;</td>")
			End If
			If Not IsNull(rs("LCPvMCS")) Then
				Response.Write("<td>" &  FormatCurrency(rs("LCPvMCS"),0,-2,0)  & "</td>")
			Else
				Response.Write("<td>&nbsp;</td>")
			End If
			If Not IsNull(rs("ThreePAvgvMCS")) Then
				Response.Write("<td>" &  FormatCurrency(rs("ThreePAvgvMCS"),0,-2,0)  & "</td>")
			Else
				Response.Write("<td>&nbsp;</td>")
			End If
			If Not IsNull(rs("TwelvePAvgvMCS")) Then
				Response.Write("<td>" &  FormatCurrency(rs("TwelvePAvgvMCS"),0,-2,0)  & "</td>")
			Else
				Response.Write("<td>&nbsp;</td>")
			End If
			If Not IsNull(rs("CPvMCS")) Then
				Response.Write("<td>" &  FormatCurrency(rs("CPvMCS"),0,-2,0)  & "</td>")
			Else
				Response.Write("<td>&nbsp;</td>")
			End If
			If rs("EqpValue")> 0 Then	
				If IsNumeric(rs("LCPROI")) Then
					Response.Write("<td>" &   FormatNumber(rs("LCPROI"),1)  & "</td>")
				Else
					Response.Write("<td>No Sales</td>")
				End If
				If IsNumeric(rs("ThreePAvgROI")) Then
					Response.Write("<td>" & FormatNumber(rs("ThreePAvgROI"),1) & "</td>")
				Else
					Response.Write("<td>&nbsp;</td>")
				End If
				' Write equipment value regardless of ROI
				'Response.Write("<td>" & FormatCurrency(rs("EqpValue"),0) & "</td>")
				Response.Write("<td><a data-toggle='modal' data-show='true' href='#' data-cust-id='" & rs("CustID") & "' data-lcp-gp='0' data-target='#modalEquipmentVPC' data-tooltip='true' data-title='View Customer Equipment'>" & FormatCurrency(rs("EqpValue"),0) & "</a></td>")
			Else
				Response.Write("<td>&nbsp;</td>")
				Response.Write("<td>&nbsp;</td>")
				Response.Write("<td>&nbsp;</td>")								
			End If
			Select Case MUV_READ("LOHVAR")
					Case "Secondary"
					    If Instr(rs("PrimarySalesmanName") ," ") <> 0 Then
							Response.Write("<td>" & Left(rs("PrimarySalesmanName"),Instr(rs("PrimarySalesmanName")," ")+1) & "</td>")
						Else
							Response.Write("<td>" & rs("PrimarySalesmanName")& "</td>")
						End If
					Case "Primary"
					    If Instr(rs("SecondarySalesmanName")," ") <> 0 Then
							Response.Write("<td>" & Left(rs("SecondarySalesmanName"),Instr(rs("SecondarySalesmanName")," ")+1) & "</td>")
						Else
							Response.Write("<td>" & rs("SecondarySalesmanName")& "</td>")
						End If
					Case "CustType"
					    If Instr(rs("SecondarySalesmanName")," ") <> 0 Then
							Response.Write("<td>" & Left(rs("SecondarySalesmanName"),Instr(rs("SecondarySalesmanName")," ")+1) & "</td>")
						Else
							Response.Write("<td>" & rs("SecondarySalesmanName")& "</td>")
						End If
			End Select	
			Response.Write("<td>" & rs("CustomerTypeName")& "</td>")
			
			'Response.Write("<td>" & UserHasAnyUnviewedNotes(rs("CustID")) & "</td>")
		
            If (UserHasAnyUnviewedNotes(rs("CustID"))) Then
            	Response.Write("<td><a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='" & rs("CustID") & "' class='ole' rel='tooltip' style='cursor:pointer;'><i class='fa fa-file-text-o faa-pulse animated fa-2x' aria-hidden='true'></i></a></td>")
           	Else
           	 	Response.Write("<td><a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='" & rs("CustID") & " class='ole' rel='tooltip' style='cursor:pointer;'><i class='fa fa-file-text-o' aria-hidden='true'></i></a></td>")
           	End If
			
			
			Response.Write("<td>"& "123abc" & "</td>")
				
			rs.movenext
				
		Loop


%>
</tbody>

					<tfoot>
						<tr>
							<td>&nbsp;</td>
							<td>Total</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td class="border-left border-right">&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
							<td>&nbsp;</td>
						</tr>
					</tfoot>


</table>
</div>
</div>