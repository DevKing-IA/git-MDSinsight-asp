<style type="text/css">
	.row-common{
		border: 1px solid #dbdece;
		padding-top: 10px;
		padding-bottom: 10px;
		margin-bottom: 10px;
		font-size: 12px;
	}
</style>

<% ' Lookup the customer record to get the other stuff we need

SQL = "SELECT * FROM " & MUV_Read("SQL_Owner")  & ".AR_Customer WHERE CustNum = '" & SelectedCustomer & "'"
						
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
If not rs.Eof Then
	tmpStatus = rs("AcctStatus")
	tmpChain = rs("ChainNum")
	tmpAssociatedNumber = rs("ArOldAcctNum")
	tmpSalesman = rs("Salesman")
	tmpSalesman2 = rs("SecondarySalesman")	
	tmpReferral = rs("ReferalCode")	
	tmpARrep = rs("ArRep")		
	tmpCustType = rs("CustType")
	tmpRetType = rs("ReturnType")
	tmpRetTime = rs("ReturnTime")
	tmpCustName = rs("Name")
	tmpAddr1 = rs("Addr1")
	tmpAddr2 = rs("Addr2")
	tmpCSZ = rs("CityStateZip")
	tmpPhone = rs("Phone")
	If Not IsNull(rs("LastBuyDate")) Then tmpLastBuy = FormatDateTime(rs("LastBuyDate"),2)
	If Not IsNull(rs("InstallDate")) Then tmpInstall = FormatDateTime(rs("InstallDate"),2)
	If Not IsNull(rs("CancelDate")) Then  tmpCancel = FormatDateTime(rs("CancelDate"),2)
	If Not IsNull(rs("ReturnDate")) Then  tmpReturn = FormatDateTime(rs("ReturnDate"),2) Else tmpReturn = Null
End IF
rs.close
%>
 
<div class="col-lg-6 box">
 
	<table style="width:100%;">
	<tr><td align="left"><b><%=GetTerm("Account")%> #:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=SelectedCustomer%>

	<input type="hidden" id="txtAccount" name="txtAccount" value="<%=SelectedCustomer %>">
	<input type="hidden" id="txtCustID" name="txtCustID" value="<%=SelectedCustomer %>"></td></tr>	
	<tr><td align="left"><b>Name:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpCustName %></td></tr>
	<tr><td align="left"><b>Addr1:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpAddr1 %></td></tr>
	<% If tmpAddr2  <> "" Then %>
		<tr><td align="left"><b>Addr2:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpAddr2 %></td></tr>
	<%End If%>
	<tr><td align="left"><b>City,State,Zip:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpCSZ %></td></tr>
	<tr><td align="left"><b>Phone:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpPhone %></td></tr>
<!-- 	
	<tr><td align="left"><b>Status:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpStatus %></td></tr>
	<tr><td align="left"><b>Chain ID:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpChain %></td></tr>
	<tr><td align="left"><b>Associated <%=GetTerm("Account")%> #:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpAssociatedNumber %></td></tr>
	<tr><td align="left"><b>Last Buy Date:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpLastBuy %></td></tr>
	<tr><td align="left"><b>Install Date:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpInstall %></td></tr>
	<tr><td align="left"><b>Cancel Date:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpCancel %></td></tr>
	<tr>
		<td align="left"><b><%=GetTerm("Primary Salesman")%>:</b></td><td>&nbsp;&nbsp;&nbsp;</td>
		<td align="left">
		<% 
		tmpVar = GetSalesmanNameAndEmailBySlsmnSequence(tmpSalesman)
		If tmpVar <> "*Not Found*" Then
			Response.Write(Left(tmpVar,Instr(tmpVar ,"~")-1))
			If Right(tmpVar,len(tmpVar)-Instr(tmpVar ,"~")) <> "Not Found" Then
				Response.Write("&nbsp;<a href='mailto:" & Right(tmpVar,len(tmpVar)-Instr(tmpVar ,"~")) & "'><i class='fa fa-envelope-o'></i></a>")
			End If
		Else
			Response.Write("*Not Found*")
		End If%>
		</td>
	</tr>
	<tr>
		<td align="left"><b><%=GetTerm("Secondary Salesman")%>:</b></td><td>&nbsp;&nbsp;&nbsp;</td>
		<td align="left">
		<% 
		tmpVar = GetSalesmanNameAndEmailBySlsmnSequence(tmpSalesman2)
		If tmpVar <> "*Not Found*" Then
			Response.Write(Left(tmpVar,Instr(tmpVar ,"~")-1))
			If Right(tmpVar,len(tmpVar)-Instr(tmpVar ,"~")) <> "Not Found" Then
				Response.Write("&nbsp;<a href='mailto:" & Right(tmpVar,len(tmpVar)-Instr(tmpVar ,"~")) & "'><i class='fa fa-envelope-o'></i></a>")
			End If
		Else
			Response.Write("*Not Found*")
		End If%>
		</td>
	</tr>
	<tr>
	<tr>
		<td align="left"><b>Referral:</b></td><td>&nbsp;&nbsp;&nbsp;</td>
		<td align="left">
			<%=GetReferralNameByCode(tmpReferral)%>
		</td>
	</tr>
	<tr>
		<td align="left"><b><%=GetTerm("A/r")%> Rep:</b></td><td>&nbsp;&nbsp;&nbsp;</td>
		<td align="left">
			<% 
			tmpVar = GetSalesmanNameAndEmailBySlsmnSequence(tmpARrep)
			If tmpVar <> "*Not Found*" Then
				Response.Write(Left(tmpVar,Instr(tmpVar ,"~")-1))
				If Right(tmpVar,len(tmpVar)-Instr(tmpVar ,"~")) <> "Not Found" Then
					Response.Write("&nbsp;<a href='mailto:" & Right(tmpVar,len(tmpVar)-Instr(tmpVar ,"~")) & "'><i class='fa fa-envelope-o'></i></a>")
				End If
			Else
				Response.Write("*Not Found*")
			End If%>
		</td>
	</tr>
	<tr>
		<td align="left"><b><%=GetTerm("Account")%> Type:</b></td><td>&nbsp;&nbsp;&nbsp;</td>
		<td align="left">
			<%=GetCustTypeByCode(tmpCustType )%>
		</td>
	</tr>
	!-->
	</table>
 </div>

 
