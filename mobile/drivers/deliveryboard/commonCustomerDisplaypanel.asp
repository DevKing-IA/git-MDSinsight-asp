<style type="text/css">
	.row-common{
		border: 1px solid #dbdece;
		padding-top: 10px;
		padding-bottom: 10px;
		margin-bottom: 10px;
		font-size: 12px;
	}
</style>

<%
'Lookup the customer record to get the other stuff we need

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

<!-- row !-->
<div class="row row-common">
	<div class="col-lg-6">
		<table style="width:100%;">
			<tr><td align="right"><b><%=GetTerm("Account")%> #:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=SelectedCustomer%>
			<!-- Need both these hiddwn fields for forms written before this include file existed !-->
			<input type="hidden" id="txtAccount" name="txtAccount" value="<%=SelectedCustomer %>">
			<input type="hidden" id="txtCustID" name="txtCustID" value="<%=SelectedCustomer %>"></td></tr>	
			<tr><td align="right"><b>Name:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpCustName %></td></tr>
			<tr><td align="right"><b>Addr1:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpAddr1 %></td></tr>
			<% If tmpAddr2  <> "" Then %>
			<tr><td align="right"><b>Addr2:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpAddr2 %></td></tr>
			<%End If%>
			<tr><td align="right"><b>City,State,Zip:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpCSZ %></td></tr>
			<tr><td align="right"><b>Phone:</b></td><td>&nbsp;&nbsp;&nbsp;</td><td align="left"><%=tmpPhone %></td></tr>
		</table>
	</div>
</div>
<!-- eof row !-->