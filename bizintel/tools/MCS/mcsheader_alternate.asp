<!--<div class='table-responsive' style="border:1px #ddd solid;">-->
<div class='table-responsive'>
	<table class='table table-condensed table-top2'>
		<tbody>
			<tr>


				<!----- BOX 2 ----->
				<td width="20%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<td width="50%"># of MCS accounts</td>
									<td width="50%" align="center"><strong><%= TotalMCSClients %></strong></td>
								</tr>
								<tr>
									<td width="50%">Commitment</td>
									<td width="50%" align="center"><strong><%= FormatCurrency(TotalMCSCommitment,0) %> </strong></td>
								</tr>
								<tr>
									<% If Not IsNumeric(TotalSalesAllMCSCustomers) Then TotalSalesAllMCSCustomers = 0 %>
									<td width="50%">Actual <%= MonthName(Month(DateAdd("m",-1,ReportDate))) %> sales</td>
									<td width="50%" align="center"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers,0) %> </strong></td>
								</tr>
								<tr>
									<td width="50%">Variance</td>
									<% If Round(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) = 0 Then %>
										<td width="50%" align="center" class="neutral"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %> </strong></td>
									<% ElseIf Round(TotalMCSCommitment - TotalSalesAllMCSCustomers,0) > 0 Then %>
										<td width="50%" align="center" class="negative"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %> </strong></td>									
									<% Else %>
										<td width="50%" align="center" class="positive"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %> </strong></td>									
									<% End If %>
								</tr>
							</tbody>
						</table>
					</div>
				</td>
				<!----- END BOX 2 ----->



				<!----- BOX 3 ----->
				<td width="20%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<td width="50%">Exceeded or Met</td>
									<td width="50%" align="center"><strong><%= TotalCustomersOver %>&nbsp;&nbsp;<font class="positive"><%= FormatCurrency(TotalOverDollars,0) %></font></strong></td>
								</tr>
								<tr>
									<td width="50%">Missed&nbsp;<small>(not including zero)</small></td>
									<td width="50%" align="center"><strong><%= TotalCustomersUnder %>&nbsp;&nbsp;<font class="negative"><%= FormatCurrency(TotalUnderDollars * -1,0) %></font></strong></td>
								</tr>
								<tr>
									<td width="50%">Zero</td>
									<td width="50%" align="center"><strong><%= TotalCustomersZeroSales %>&nbsp;&nbsp;<font class="negative">(<%= FormatCurrency(TotalZeroSalesCommitment ,0) %>)</font></strong></td>
								</tr>
								<tr>
									<td width="50%">Overcomes</td>
									<td width="50%" align="center"><strong><font color='blue'><%= TotalCustomersUnderButRecovered %></font></strong></td>
								</tr>
							</tbody>
						</table>
					</div>
				</td>
				<!----- END BOX 3 ----->

				<!----- BOX 4 ----->
				<td width="20%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<% If Not IsNumeric(TotalLVFLastMonth) Then TotalLVFLastMonth = 0 %>
									<td width="50%">LVF charges in <%=MonthName(Month(DateAdd("m",-1,ReportDate)))%></td>
									<td width="50%" align="center" class="neutral"><strong><%= FormatCurrency(TotalLVFLastMonth,0) %></strong></td>
								</tr>
								<tr>
									<% If Not IsNumeric(TotalPendingLVF) Then TotalPendingLVF = 0 %>
									<td width="50%">LVF Pending</td>
									<td width="50%" align="center" class="neutral"><strong><%= FormatCurrency(TotalPendingLVF,0) %></strong></td>
								</tr>
							</tbody>
						</table>
					</div>
				</td>
				<!----- END BOX 4 ----->
				
				<!----- BOX 4 ----->
				<td width="10%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
					</div>
				</td>
				<!----- END BOX 4 ----->

				<!----- BOX 5 ----->
				<td width="30%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<td width="70%">										
										<input type="checkbox" name="chkIncludeDeficitCovered" id="chkIncludeDeficitCovered" <% IF IncludeDeficitCovered THEN Response.write ("checked=""true""") %>>&nbsp;Include overcomes<br>
									</td>
									<td width="30%">										
										<input type="checkbox" name="chkHideNoActionNeeded" id="chkHideNoActionNeeded" <% IF HideNoActionNeeded THEN Response.write ("checked=""true""") %>>&nbsp;Hide No Action Needed<br>
									</td>	
								</tr>
								<tr>
									<td width="70%">
										<input type="checkbox" name="chkShowZeroSalesCusts" id="chkShowZeroSalesCusts" <% If ShowZeroSalesCusts= 1 Then Response.Write("checked") %>>&nbsp;Show only customers with $0 (or less) sales in <%= MonthName(Month(DateAdd("m",-1,ReportDate))) %><br>
									</td>
									<td width="30%">
										<input type="checkbox" name="chkShowAllCusts" id="chkShowAllCusts" <% If ShowAllCusts = 1 Then Response.Write("checked") %>>&nbsp;Show all customers<br>
									</td>									
								</tr>
								<tr>
									<td width="100%" colspan="2">
										<input type="checkbox" name="chkApplyRule" id="chkApplyRule" <% If ApplyRule = 1 Then Response.Write("checked") %>>&nbsp;Apply the $100 - 10% rule&nbsp;&nbsp;(<strong><%= TotalApplyRuleCount%></strong>&nbsp;<%= GetTerm("customers")%> filtered out)
										<br><small>Variance &lt $100 <u>and</u> represents &lt 10% of the customer's mcs</small>
									</td>
								</tr>
							</tbody>
						</table>
					</div>
				</td>
				<!----- END BOX 5 ----->
				
 					 			
			</tr>
		</tbody>
	</table>
</div>
