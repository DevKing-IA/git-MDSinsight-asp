<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Number of <%= GetTerm("customers")%> added to MCS program</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustsMCSAdded %></strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>	
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Net increase / decrease to total MCS commitment</font></td>
					<% If TotalNetMCSChange = 0 Then %>
						<td width="3%" align="right" class="neutral"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalNetMCSChange ,0)%></strong></font></td>
					<% ElseIf TotalNetMCSChange < 0 Then %>
						<td width="3%" align="right" class="negative"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalNetMCSChange ,0)%>&darr;</strong></font></td>					
					<% Else %>
						<td width="3%" align="right" class="positive"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalNetMCSChange ,0)%>&uarr;</strong></font></td>					
					<% End If %>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Number of reported <%= GetTerm("customers")%>&nbsp;reviewed</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalActedUpon %></strong></font></td>
				</tr>
				<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Number of <%= GetTerm("customers")%> removed from  MCS program</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustsMCSRemoved %></strong> </font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Net increase / decrease to total LVFs</font></td>
					<% If TotalNetLVFChange = 0 Then %>
						<td width="3%" align="right" class="neutral"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalNetLVFChange ,0)%></strong></font></td>
					<% ElseIf TotalNetMCSChange < 0 Then %>
						<td width="3%" align="right" class="negative"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalNetLVFChange ,0)%>&darr;</strong></font></td>					
					<% Else %>
						<td width="3%" align="right" class="positive"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalNetLVFChange ,0)%>&uarr;</strong></font></td>					
					<% End If %>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Number of reported <%= GetTerm("customers")%>&nbsp;<u>not</u> reviewed</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalCustomersUnder - TotalActedUpon %></strong></font></td>
				</tr>
				<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong>&nbsp;</strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Number of <%= GetTerm("customers")%> marked No Action Necessary</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalNoAction %></strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
				</tr>
				<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong>&nbsp;</strong> </font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Number of sales rep follow-ups</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalFollowup %></strong></font></td>						
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
				<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong>&nbsp;</strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Number of <%= GetTerm("customers")%> to be invoiced for LVF</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= TotalNumCustsInvoiced %></strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
							
				<tr>
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>										
					<td width="31%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">Total LVF amount to be invoiced</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt"><strong><%= FormatCurrency(TotalLVFInvoicedAmount,0) %></strong></font></td>
					<td width="1%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
					<td width="27%"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>
					<td width="3%" align="right"><font face="Consolas" style="font-size: <%= FontSizeVar %>pt">&nbsp;</font></td>					
				</tr>
				<tr>
					<td colspan ="8">&nbsp;</td>
				</tr>