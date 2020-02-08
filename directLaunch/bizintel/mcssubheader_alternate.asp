<%
Colwidth1 = 10
Colwidth2 = 10
Colwidth3 = 5
Colwidth4 = 14
Colwidth5 = 10
Colwidth6 = 5
Colwidth7 = 14
Colwidth8 = 10
Colwidth9 = 22


Colwidth1 = 13
Colwidth2 = 8
Colwidth3 = 8
Colwidth4 = 16
Colwidth5 = 10


%>

				<tr>
					<td width="<%=Colwidth1%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">ADDS</font></td>
					<td width="<%=Colwidth2%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustsMCSAdded %></strong></font>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"  class="positive"><strong>&nbsp;<%= FormatCurrency(TotalCustsMCSAddedDollars,0)%></strong></font>
					</td>
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>	
					<td width="<%=Colwidth4%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Missed Goal (not overcome)</font></td>
					<td width="<%=Colwidth5%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= (TotalCustomersUnder + TotalCustomersZeroSales) - TotalCustomersUnderButRecovered %></strong></font>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt" class="negative"><strong>&nbsp;<%= FormatCurrency((TotalUnderDollars + TotalZeroSalesCommitment)-TotalUnderButRecoveredDeficitDollars,0) %></strong></font>
					</td>					
					<td width="<%=Colwidth6%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="<%=Colwidth7%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth8%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth9%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
				</tr>
				<tr>
					<td width="<%=Colwidth1%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">REMOVES</font></td>
					<td width="<%=Colwidth2%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustsMCSRemoved %></strong></font>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt" class="negative"><strong>&nbsp;<%=FormatCurrency(TotalCustsMCSRemovedDollars,0) %></strong></font>
					</td>
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="<%=Colwidth4%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Not Reviewed</font></td>
					<td width="<%=Colwidth5%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= (TotalCustomersUnder + TotalCustomersZeroSales ) - (TotalActedUpon + TotalCustomersUnderButRecovered)%></strong></font>
						<!--<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong>&nbsp;<%= FormatCurrency(TotalNotActedUponMCSDollars,0) %></strong></font>-->
						<%
						'Formula: y/x=p%
						y = ((TotalCustomersUnder + TotalCustomersZeroSales ) - TotalActedUpon)
						x = (TotalCustomersUnder + TotalCustomersZeroSales)
						p = (y/x)
						p = p * 100
						P = Round(p,0)
						%>	
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong>&nbsp;(<%= p %>%)</strong></font> 
					</td>					
					<td width="<%=Colwidth6%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="<%=Colwidth7%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth8%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth9%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
				<tr>
					<td width="<%=Colwidth1%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth2%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<!--<td width="14%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">No Action Necessary</font></td>
					<td width="14%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalNoAction %></strong></font></td>-->
					<td width="<%=Colwidth4%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth5%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth6%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="<%=Colwidth7%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth8%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="<%=Colwidth9%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
				<tr>
					<td width="<%=Colwidth1%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth2%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong>&nbsp;</strong> </font></td>
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<!--<td width="14%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Sales rep followup</font></td>
					<td width="14%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalFollowup %></strong></font></td>-->
					<td width="<%=Colwidth4%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth5%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth6%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="<%=Colwidth7%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth8%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="<%=Colwidth9%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
				<tr>
					<td width="<%=Colwidth1%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth2%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong>&nbsp;</strong></font></td>
					<td width="<%=Colwidth3%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<!--<td width="14%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">LVFS To Invoice</font></td>
					<td width="14%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalNumCustsInvoiced %></strong></font>
						<font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong>&nbsp;<%= FormatCurrency(TotalLVFInvoicedAmount,0) %></strong></font>
					</td>-->
					<td width="<%=Colwidth4%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth5%>%" align="center"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth6%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="<%=Colwidth7%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="<%=Colwidth8%>%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="<%=Colwidth9%>%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
							
				<tr>
					<td colspan ="9">&nbsp;</td>					
				</tr>
				<tr>
					<td colspan ="9">&nbsp;</td>
				</tr>