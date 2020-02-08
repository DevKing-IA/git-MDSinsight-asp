<%
Colwidth1 = 13
Colwidth2 = 8
Colwidth3 = 8
Colwidth4 = 14
Colwidth5 = 10
Colwidth6 = 8
Colwidth7 = 14
Colwidth8 = 8
Colwidth9 = 13
Colwidth10 = 42
%>
				<tr>
					<td width="<%=Colwidth1%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"># of MCS accounts</font></td>
					<td width="<%=Colwidth2%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalMCSClients %></strong></font></td>
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="<%=Colwidth4%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Exceeded or met</font></td>
					<td width="<%=Colwidth5%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustomersOver %></strong></font>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt" class="positive"><strong>&nbsp;<%= FormatCurrency(TotalOverDollars,0) %></strong></font>
					</td>
					<td width="<%=Colwidth6%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth7%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">LVF charges in <%= MonthName(Month(DateAdd("m",-1,ReportDate))) %></font></td>	
					<td width="<%=Colwidth8%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= FormatCurrency(TotalLVFLastMonth,0) %></font></td>	
					<td width="<%=Colwidth9%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>	
				</tr>
				
				<tr>
					<td width="<%=Colwidth1%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Commitment</font></td>
					<td width="<%=Colwidth2%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalMCSCommitment,0) %></strong> </font></td>
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="<%=Colwidth4%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Missed<small>(not including zero)</small></font></td>
					<td width="<%=Colwidth5%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustomersUnder %></strong></font>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt" class="negative"><strong>&nbsp;<%= FormatCurrency(TotalUnderDollars * -1,0) %></strong></font>					
					</td>
					<td width="<%=Colwidth6%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth7%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">LVF Pending</font></td>	
					<td width="<%=Colwidth8%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><%= FormatCurrency(TotalPendingLVF,0) %></font></td>	
					<td width="<%=Colwidth9%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>	
				</tr>
				
				<tr>
					<% If Not IsNumeric(TotalSalesAllMCSCustomers) Then TotalSalesAllMCSCustomers = 0 %>
					<td width="<%=Colwidth1%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Actual <%= MonthName(Month(DateAdd("m",-1,ReportDate))) %> sales</font></td>
					<td width="<%=Colwidth2%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers,0) %></strong></font></td>
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Zero</font></td>
					<td width="<%=Colwidth5%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustomersZeroSales %></strong></font>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"class="negative"><strong>(<%= FormatCurrency(TotalZeroSalesCommitment ,0) %>)</strong></font>
					</td>
					<td width="<%=Colwidth10%>%" colspan="4"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
				</tr>
				
				<tr>
					<td width="<%=Colwidth1%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Variance</font></td>
					<% If Round(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) = 0 Then %>
						<td width="<%=Colwidth2%>%" align="center" class="neutral"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %></strong> </font></td>
					<% ElseIf Round(TotalMCSCommitment - TotalSalesAllMCSCustomers,0) > 0 Then %>
						<td width="<%=Colwidth2%>%" align="center" class="negative"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %></strong> </font></td>									
					<% Else %>
						<td width="<%=Colwidth2%>%" align="center" class="positive"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalSalesAllMCSCustomers - TotalMCSCommitment,0) %></strong> </font></td>									
					<% End If %>
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Overcomes</font></td>
					<td width="<%=Colwidth5%>" align="center" class="negative"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt" color="blue"><strong><%= TotalCustomersUnderButRecovered %></strong></font>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt" color="blue">&nbsp;<strong><%= FormatCurrency(TotalUnderButRecoveredDeficitDollars,0) %></strong></font>
					</td>						
					<td width="<%=Colwidth10%>%" colspan="4"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
				
				<tr>
					<td colspan ="9">&nbsp;</td>
				</tr>
