<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<!--#include file="InSightFuncs_Orders.asp"-->
<!--#include file="rePostings_inc.asp"-->
<%
'***************************************************
'List of all the AJAX functions & subs
'***************************************************
'Sub rePostOrderToBackend()
'Sub returnsUMSForUnmappedProductCode()
'***************************************************
'End List of all the AJAX functions & subs
'***************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'ALL AJAX MODAL SUBROUTINES AND FUNCTIONS BELOW THIS AREA

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


action = Request("action")

Select Case action
	Case "rePostOrderToBackend"
		rePostOrderToBackend()
	Case "returnsUMSForUnmappedTaxableProductCode"
		returnsUMSForUnmappedTaxableProductCode()
	Case "returnsUMSForUnmappedNonTaxableProductCode"
		returnsUMSForUnmappedNonTaxableProductCode()
End Select




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub returnsUMSForUnmappedTaxableProductCode() 

	prodSKU = Request.Form("prodSKU")
	UM = Request.Form("UM")
	%>

	Unit of Measure: 
	<select class="C_Country_Modal form-control" id="txtUnmappedTaxableUM" name="txtUnmappedTaxableUM" style="width:50px;"> 
		<% 
		  	SQLProductsTableUM = "SELECT Distinct(prodCasePricing) FROM IC_Product WHERE prodSKU = '" & prodSKU & "'"
		
			Set cnnProductsTableUM = Server.CreateObject("ADODB.Connection")
			cnnProductsTableUM.open (Session("ClientCnnString"))
			Set rsProductsTableUM = Server.CreateObject("ADODB.Recordset")
			rsProductsTableUM.CursorLocation = 3 
			Set rsProductsTableUM = cnnProductsTableUM.Execute(SQLProductsTableUM)
				
			If not rsProductsTableUM.EOF Then
			
				If rsProductsTableUM("prodCasePricing") = "N" Then
					%><option value="N" selected="selected">N</option><%
				ElseIf rsProductsTableUM("prodCasePricing") = "U" Then
					%><option value="U" <% If UM = rsProductsTableUM("prodCasePricing") Then Response.Write("selected='selected'") %>>U</option><%
					%><option value="C" <% If UM = rsProductsTableUM("prodCasePricing") Then Response.Write("selected='selected'") %>>C</option><%
				ElseIf rsProductsTableUM("prodCasePricing") = "C" Then
					%><option value="U" <% If UM = rsProductsTableUM("prodCasePricing") Then Response.Write("selected='selected'") %>>U</option><%
					%><option value="C" <% If UM = rsProductsTableUM("prodCasePricing") Then Response.Write("selected='selected'") %>>C</option><%
				End If 
			Else
			%>
				<option value="U">U</option>
     			<option value="C">C</option>
     			<option value="N">N</option>
			<%
			End If
			
			set rsProductsTableUM = Nothing
			cnnProductsTableUM.close
			set cnnProductsTableUM = Nothing
		%>									
	</select>

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub returnsUMSForUnmappedNonTaxableProductCode() 

	prodSKU = Request.Form("prodSKU")
	%>

	Unit of Measure: 
	<select class="C_Country_Modal form-control" id="txtUnmappedNonTaxableUM" name="txtUnmappedNonTaxableUM" style="width:50px;"> 
		<% 
		  	SQLProductsTableUM = "SELECT Distinct(prodCasePricing) FROM IC_Product WHERE prodSKU = '" & prodSKU & "'"
		
			Set cnnProductsTableUM = Server.CreateObject("ADODB.Connection")
			cnnProductsTableUM.open (Session("ClientCnnString"))
			Set rsProductsTableUM = Server.CreateObject("ADODB.Recordset")
			rsProductsTableUM.CursorLocation = 3 
			Set rsProductsTableUM = cnnProductsTableUM.Execute(SQLProductsTableUM)
				
			If not rsProductsTableUM.EOF Then
			
				If rsProductsTableUM("prodCasePricing") = "N" Then
					%><option value="N">N</option><%
				ElseIf rsProductsTableUM("prodCasePricing") = "U" Then
					%><option value="U">U</option><%
					%><option value="C">C</option><%
				ElseIf rsProductsTableUM("prodCasePricing") = "C" Then
					%><option value="U">U</option><%
					%><option value="C">C</option><%
				End If 
			Else
			%>
				<option value="U">U</option>
     			<option value="C">C</option>
     			<option value="N">N</option>
			<%
			End If
			
			set rsProductsTableUM = Nothing
			cnnProductsTableUM.close
			set cnnProductsTableUM = Nothing
		%>									
	</select>

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub rePostOrderToBackend()
	
	If Request.Form("IntRecID") <> "" Then
		

		IntRecID = Request.Form("IntRecID") 
	
		Set cnnPostOrderToBackend = Server.CreateObject("ADODB.Connection")
		cnnPostOrderToBackend.open (Session("ClientCnnString"))

		Set rsRepost = Server.CreateObject("ADODB.Recordset")
		rsRepost.CursorLocation = 3 
		
		SQLrsRepost = "SELECT * FROM API_OR_OrderHeader WHERE InternalRecordIdentifier = " & IntRecID 

		Set rsRepost = cnnPostOrderToBackend.Execute(SQLrsRepost)

		If Not rsRepost.Eof Then
			Call rePostOrderToBackend1(rsRepost("OrderID"),"UPSERT")
		End If
		
		Set rs = Nothing
		cnnPostOrderToBackend.Close
		Set cnnPostOrderToBackend = Nothing
		
	End If


End Sub


Function Number_Of_Lines (passedOrderHeaderRecID)

	resultNumber_Of_Lines = ""
	
	Set cnnNumber_Of_Lines = Server.CreateObject("ADODB.Connection")
	cnnNumber_Of_Lines.open (Session("ClientCnnString"))
	
	Set rsNumber_Of_Lines = Server.CreateObject("ADODB.Recordset")
	rsNumber_Of_Lines.CursorLocation = 3 

	SQLNumber_Of_Lines = "SELECT COUNT(*) as Expr1 FROM API_OR_OrderDetail WHERE OrderHeaderRecID = " & passedOrderHeaderRecID

	Set rsNumber_Of_Lines = cnnNumber_Of_Lines.Execute(SQLNumber_Of_Lines)
	
	If Not rsNumber_Of_Lines.EOF Then resultNumber_Of_Lines = rsNumber_Of_Lines("Expr1")
	
	Set rsNumber_Of_Lines = Nothing
	cnnNumber_Of_Lines.Close
	Set cnnNumber_Of_Lines = Nothing

	Number_Of_Lines = resultNumber_Of_Lines

End Function

%>