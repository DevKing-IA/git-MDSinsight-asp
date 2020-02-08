<div class='table-responsive table-top'>
	<table class='table table-condensed'>
		<tbody>
			<tr>


				<!----- BOX 2 ----->
				<td width="33%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<td width="70%">Number of <%= GetTerm("customers")%> in MCS program</td>
									<td width="30%" align="right"><strong><%= TotalMCSClients %></strong></td>
								</tr>
								<tr>
									<td width="70%">Monthly dollar commitment for all MCS <%= GetTerm("customers")%></td>
									<td width="30%" align="right"><strong><%= FormatCurrency(TotalMCSCommitment,0) %> </strong></td>
								</tr>
								<tr>
									<% If Not IsNumeric(TotalSalesAllMCSCustomers) Then TotalSalesAllMCSCustomers = 0 %>
									<td width="70%">Actual <%= MonthName(Month(DateAdd("m",-1,ReportDate))) %> sales for all MCS <%= GetTerm("customers")%></td>
									<td width="30%" align="right"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers,0) %> </strong></td>
								</tr>
								<tr>
									<td width="70%">Overall variance of sales vs MCS expectation</td>
									<% If Round(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) = 0 Then %>
										<td width="30%" align="right" class="neutral"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %> </strong></td>
									<% ElseIf Round(TotalMCSCommitment - TotalSalesAllMCSCustomers,0) > 0 Then %>
										<td width="30%" align="right" class="negative"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %> </strong></td>									
									<% Else %>
										<td width="30%" align="right" class="positive"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %> </strong></td>									
									<% End If %>
								</tr>
								<tr>
									<td width="70%">Customers under MCS but <%= MonthName(Month(ReportDate)) %> sales covered deficit</td>
									<td width="30%" align="right"><strong><font color='blue'><%= TotalCustomersUnderButRecovered %></font></strong></td>
								</tr>

							</tbody>
						</table>
					</div>
				</td>
				<!----- END BOX 2 ----->



				<!----- BOX 3 ----->
				<td width="33%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<td width="70%"><%= GetTerm("Customers")%> who met or exceeded their MCS goal</td>
									<td width="30%" align="right"><strong><%= TotalCustomersOver %></strong></td>
								</tr>
								<tr>
									<td width="70%">Dollar variance of over performers</td>
									<td width="30%" align="right" class="positive"><strong><%= FormatCurrency(TotalOverDollars,0) %></strong></td>
								</tr>
								<tr>
									<td width="70%"><%= GetTerm("Customers")%> who <u>did not</u> meet their MCS goal</td>
									<td width="20%" align="right"><strong><%= TotalCustomersUnder %></strong></td>
								</tr>
								<tr>
									<td width="70%">Dollar variance of under performers</td>
									<td width="30%" align="right" class="negative"><strong><%= FormatCurrency(TotalUnderDollars * -1,0) %></strong></td>
								</tr>
								<tr>
									<% If Not IsNumeric(TotalLVFLastMonth) Then TotalLVFLastMonth = 0 %>
									<td width="70%">Total LVF dollars charged in <%=MonthName(Month(DateAdd("m",-1,ReportDate)))%></td>
									<td width="30%" align="right" class="neutral"><strong><%= FormatCurrency(TotalLVFLastMonth,0) %></strong></td>
								</tr>
								<tr>
									<% If Not IsNumeric(TotalPendingLVF) Then TotalPendingLVF = 0 %>
									<td width="70%">Pending LVF dollars for <%=MonthName(Month(ReportDate))%></td>
									<td width="30%" align="right" class="neutral"><strong><%= FormatCurrency(TotalPendingLVF,0) %></strong></td>
								</tr>

							</tbody>
						</table>
					</div>
				</td>
				<!----- END BOX 3 ----->



				<!----- BOX 4 ----->
				<td width="33%">
					<div class="table-striped table-condensed table-hover account-info-table inner-table">
						<table class="table table-striped table-condensed table-hover">
							<tbody>
								<tr>
									<td width="70%">MCS <%= GetTerm("customers")%> with $0 sales in <%= MonthName(Month(DateAdd("m",-1,ReportDate))) %></td>
									<td width="30%" align="right"><strong><%= TotalCustomersZeroSales %></strong></td>
								</tr>
								<tr>
									<td width="70%">MCS commitment of <%= GetTerm("customers")%> with $0 sales</td>
									<td width="30%" align="right" class="negative"><strong>(<%= FormatCurrency(TotalZeroSalesCommitment ,0) %>)</strong></td>
								</tr>
								<tr>
									<td width="70%">										
										<input type="checkbox" name="chkIncludeDeficitCovered" id="chkIncludeDeficitCovered" <% IF IncludeDeficitCovered THEN Response.write ("checked=""true""") %>>&nbsp;Include customers where <%= MonthName(Month(ReportDate)) %> sales covered deficit<br>
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
				<!----- END BOX 4 ----->
				
 					 			
			</tr>
		</tbody>
	</table>
</div>
