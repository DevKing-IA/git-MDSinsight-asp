<tr class="for-edit editable selected" onclick="javascript: selectRow(this);">
	<td data-id="FilterData">
		<select class="form-control" id="FilterData">
		<%
		Set cnnCustFiltersList = Server.CreateObject("ADODB.Connection")
		cnnCustFiltersList.open (Session("ClientCnnString"))
		DIM var_sql_ajax
		var_sql_ajax="SELECT * FROM IC_Filters ORDER BY displayOrder"
		Set rsFilterList = Server.CreateObject("ADODB.Recordset")
		Set rsFilterList= cnnCustFiltersList.Execute(var_sql_ajax)
		%>
		<%DO WHILE NOT rsFilterList.EOF%>
			<option value="<%=rsFilterList("InternalRecordIdentifier")%>"><%="(" & rsFilterList("FilterID") & ") " & rsFilterList("Description")%></option>
			<%rsFilterList.MoveNExt%>
		<%LOOP%>
		<%
		
		cnnCustFiltersList.Close
		
		%>
		</select>
	</td>
	<td data-id="location"><input type="text" class="form-control" id="location" placeholder="Location of Unit"></td>
	<td data-id="qty"><select class="form-control" id="qty"></select></td>
	<td data-id="FrequencyType">
		<select class="form-control" id="FrequencyType">
			<option value="D" selected>Days</option>
			<option value="W">Weeks</option>
			<option value="M">Months</option>
		</select>
	</td>
	<td data-id="FrequencyTime">
		<select class="form-control" id="FrequencyTime"></select>
	</td>
	<td data-id="Price"><input type="text" class="form-control" id="Price"></td>
	<td data-id="LastChangedData"><input type="text" class="form-control" id="LastChangedData" autocomplete="off"></td>
</tr>