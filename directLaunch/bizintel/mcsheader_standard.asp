				<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Number of <%= GetTerm("customers")%> in MCS program</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalMCSClients %></strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetTerm("Customers")%> who met or exceeded their MCS goal</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustomersOver %></strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">MCS <%= GetTerm("customers")%> with $0 sales in <%= MonthName(Month(DateAdd("m",-1,ReportDate))) %></font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustomersZeroSales %></strong></font></td>
				</tr>
				<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Monthly dollar commitment for all MCS <%= GetTerm("customers")%></font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalMCSCommitment,0) %></strong> </font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Dollar variance of over performers</font></td>
					<td width="3%" align="right" class="positive"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalOverDollars,0) %></strong> </font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">MCS commitment of <%= GetTerm("customers")%> with $0 sales</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalZeroSalesCommitment ,0) %></strong></font></td>
				</tr>
				<tr>
					<% If Not IsNumeric(TotalSalesAllMCSCustomers) Then TotalSalesAllMCSCustomers = 0 %>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Actual <%= MonthName(Month(DateAdd("m",-1,ReportDate))) %> sales for all MCS <%= GetTerm("customers")%></font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers,0) %></strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= GetTerm("Customers")%> who <u>did not</u> meet their MCS goal</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustomersUnder %></strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
				</tr>
				<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Overall variance of sales vs MCS expectation</font></td>
					<% If Round(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) = 0 Then %>
						<td width="3%" align="right" class="neutral"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %></strong> </font></td>
					<% ElseIf Round(TotalMCSCommitment - TotalSalesAllMCSCustomers,0) > 0 Then %>
						<td width="3%" align="right" class="negative"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %></strong> </font></td>									
					<% Else %>
						<td width="3%" align="right" class="positive"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %></strong> </font></td>									
					<% End If %>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Dollar variance of under performers</font></td>
					<td width="3%" align="right" class="negative"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalUnderDollars * -1,0) %></strong></font></td>						
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
				<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Customers under MCS but <%= MonthName(Month(ReportDate)) %> sales covered deficit</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><font color='blue'><strong><%= TotalCustomersUnderButRecovered %></strong></font></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<% If Not IsNumeric(TotalLVFLastMonth) Then TotalLVFLastMonth = 0 %>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Total LVF dollars charged in <%=MonthName(Month(DateAdd("m",-1,ReportDate)))%></font></td>
					<td width="3%" align="right" class="neutral"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalLVFLastMonth,0) %></strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
							
				<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<% If Not IsNumeric(TotalPendingLVF) Then TotalPendingLVF = 0 %>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Pending LVF dollars for <%=MonthName(Month(ReportDate))%></font></td>
					<% If TotalPendingLVF < 0 Then %>
						<td width="3%" align="right" class="negative"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalPendingLVF,0) %></strong></font></td>
					<% Else %>
						<td width="3%" align="right" class="neutral"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalPendingLVF,0) %></strong></font></td>
					<% End If %>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
				<tr>
					<td colspan ="8">&nbsp;</td>
				</tr>
