<!--#include file="../../inc/subsandfuncs.asp"-->
<!--#include file="../../inc/Insightfuncs.asp"-->
<%
dummy = MUV_Write("ClientID","") ' Need this here

'baseURL should always have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)
sURL = Request.ServerVariables("SERVER_NAME")


Username = Request.QueryString("u")
ClientKey = Request.QueryString("cl")
UserNo = Request.QueryString("un")
UnitUPCData = Request.QueryString("uupc")
CaseUPCData = Request.QueryString("cupc")
InventoriedItem = Request.QueryString("i")
PickableItem = Request.QueryString("p")
ProductCategoriesForInventoryReport = Request.QueryString("c")

'Response.Write("Username:" & Username & "<br>")
'Response.Write("ClientKey:" & ClientKey & "<br>")
'Response.Write("UserNo:" & UserNo & "<br>")

'**************************************************************************************
'Build WHERE Clause For Unit UPC, Case UPC, Inventoried and Pickable Items
'**************************************************************************************

WHERE_CLAUSE_UNITUPC = ""
WHERE_CLAUSE_CASEUPC = ""
WHERE_CLAUSE_INVENTORIEDITEM = ""
WHERE_CLAUSE_PICKABLEITEM = ""
WHERE_CLAUSE_CATEGORY = ""

If UnitUPCData = "NOTEMPTY" Then
	WHERE_CLAUSE_UNITUPC = " AND (prodUnitUPC <> '') "
ElseIf  UnitUPCData = "EMPTY" Then
	WHERE_CLAUSE_UNITUPC = " AND (prodUnitUPC = '' OR prodUnitUPC IS NULL) "
Else
	WHERE_CLAUSE_UNITUPC = ""
End If


If CaseUPCData = "NOTEMPTY" Then
	WHERE_CLAUSE_CASEUPC = " AND (prodCaseUPC <> '') "
ElseIf  CaseUPCData = "EMPTY" Then
	WHERE_CLAUSE_CASEUPC = " (AND prodCaseUPC = '' OR prodCaseUPC IS NULL) "
Else
	WHERE_CLAUSE_CASEUPC = ""
End If
			
			
If InventoriedItem = "YES" Then
	WHERE_CLAUSE_INVENTORIEDITEM = " AND (prodInventoriedItem = 1) "
ElseIf InventoriedItem = "NO" Then
	WHERE_CLAUSE_INVENTORIEDITEM = " AND (prodInventoriedItem = 0) "
Else
	WHERE_CLAUSE_INVENTORIEDITEM = " "
End If
			
If PickableItem = "YES" Then
	WHERE_CLAUSE_PICKABLEITEM = " AND (prodPickableItem = 1) "
ElseIf InventoriedItem = "NO" Then
	WHERE_CLAUSE_PICKABLEITEM = " AND (prodPickableItem = 0) "
Else
	WHERE_CLAUSE_PICKABLEITEM = ""
End If

CategoryArray = ""
CategoryArray = Split(ProductCategoriesForInventoryReport,",")


For z = 0 to UBound(CategoryArray)
	If z = 0 AND UBound(CategoryArray) = 0 Then
		WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & " AND (prodCategory = '" & CategoryArray(z) & "' "
	ElseIf z = 0 AND UBound(CategoryArray) > 0 Then
		WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & " AND ((prodCategory = '" & CategoryArray(z) & "') "
	Else
		WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & " OR (prodCategory = '" & CategoryArray(z) & "')"
	End If
Next	
	
If WHERE_CLAUSE_CATEGORY <> "" Then
	WHERE_CLAUSE_CATEGORY = WHERE_CLAUSE_CATEGORY & ")"
End If






SQL = "SELECT * FROM tblServerInfo WHERE clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open InsightCnnString
'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3
'Response.Write("SQL:" & SQL& "<br>")

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Product UPC Report<%
	Response.End
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	Recordset.close
	Connection.close	
End If	

SQL = "SELECT * FROM Settings_CompanyID"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)

If not rs.EOF Then
	CompanyName = rs("Stmt_CompanyName")
	CompanyIdentityColor1 = rs("CompanyIdentityColor1")
	CompanyIdentityColor2 = rs("CompanyIdentityColor2")
	If CompanyIdentityColor1 = "" Then CompanyIdentityColor1 = "#6c7271"
	If CompanyIdentityColor2 = "" Then CompanyIdentityColor2 = "#6c7271"
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing

%>

<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
		<title>Product UPC Report</title>
		
		<style type="text/css">
		body{
			margin:30px;
 			font-family: Arial;
			font-size: 13px;
			overflow-x: hidden;
			text-align: left;
		}
		
		.line{
			width: 100%;
			float: left;
		}
		
		  table{
	 	   border-collapse: collapse;
	  
 	   }

		</style>
		
	</head>
	
	<body>
		
		<!-- main table starts here !-->
 		<table width="1150" align="center">
			<tbody >
				<tr>
					<td width="100%">
		
		<!-- logo / address / account starts here !-->
		<table width="1150" style="margin-bottom:20px;">
			<tbody>
				<tr>
					
					<!-- logo !-->
					<th scope="col" align="left">
							<img src="../../clientfiles/<%=MUV_Read("ClientID")%>/logos/logo.png">
					</th>
					<!-- eof logo !-->
					
  				</tr>
			</tbody>
		</table>
		
		<!-- monthly consolidated invoice title !-->
		<table width="1150" border="1" bordercolor="#111111"  cellpadding="4" style="margin-bottom:-1px;">
			<tbody>
				<tr>
					<th scope="col" >
						<h3 style="line-height:1; margin-top:10px; margin-bottom:10px;" align="center">Product UPC Report</h3>

					</th>
				</tr>
			</tbody>
		</table>
		<!-- eof monthly consolidated invoice title !-->
		
		<!-- the table with statements starts here !-->
		<table width="1150" border="1" bordercolor="#111111" cellpadding="4" style="margin-bottom:-1px;">
			<thead>
				<tr bgcolor="<%= CompanyIdentityColor1 %>" style="color:#fff;">
				  <th scope="col">SKU</th>
				  <th scope="col">Desc</th>
				  <th scope="col">Category</th> 
				  <th scope="col">Inventoried</th>
				  <th scope="col">Pickable</th>
				  <th scope="col">Case Pricing</th>
				  <th scope="col">Case Desc</th>
				  <th scope="col">Unit UPC</th>
				  <th scope="col">Case UPC</th>
				</tr>
			</thead>
			
			<tbody>
			<% 
				SQLInvReport = " SELECT prodUnitUPC, prodCaseUPC, prodSKU, prodDescription, prodCategory, "
				SQLInvReport = SQLInvReport & "prodCasePricing, prodCaseDescription, prodInventoriedItem, prodPickableItem "
				SQLInvReport = SQLInvReport & "FROM " & MUV_Read("SQL_Owner") & ".IC_Product "
				SQLInvReport = SQLInvReport & " WHERE prodSKU <> '' " & WHERE_CLAUSE_UNITUPC
				SQLInvReport = SQLInvReport & " " & WHERE_CLAUSE_CASEUPC
				SQLInvReport = SQLInvReport & " " & WHERE_CLAUSE_INVENTORIEDITEM
				SQLInvReport = SQLInvReport & " " & WHERE_CLAUSE_PICKABLEITEM
				SQLInvReport = SQLInvReport & " " & WHERE_CLAUSE_CATEGORY
				SQLInvReport = SQLInvReport & " ORDER BY prodCategory, prodSKU ASC "

				'Response.Write(SQLInvReport & "<br>")

				Set cnnProdInvReport = Server.CreateObject("ADODB.Connection")
				cnnProdInvReport.open (Session("ClientCnnString"))
				Set rsProdInvReport = Server.CreateObject("ADODB.Recordset")
				rsProdInvReport.CursorLocation = 3 

				Set rsProdInvReport = cnnProdInvReport.Execute(SQLInvReport)

				If not rsProdInvReport.Eof Then
					TotalAmt = 0
					Do While not rsProdInvReport.Eof
					
						prodSKU = rsProdInvReport("prodSKU")
						prodDescription = rsProdInvReport("prodDescription")
						prodCategoryID = rsProdInvReport("prodCategory")

						If prodCategoryID <> "" Then
							prodCategoryName = GetCategoryByID(prodCategoryID)
						End If

						prodInventoriedItem = rsProdInvReport("prodInventoriedItem")

						If prodInventoriedItem = 1 OR prodInventoriedItem = vbtrue Then
							prodInventoriedItemDisplay = "YES"
						Else
							prodInventoriedItemDisplay = "NO"																		
						End If

						prodPickableItem = rsProdInvReport("prodPickableItem")

						If prodPickableItem = 1 OR prodPickableItem = vbtrue Then
							prodPickableItemDisplay = "YES"
						Else
							prodPickableItemDisplay = "NO"																		
						End If

						prodCasePricing = rsProdInvReport("prodCasePricing")
						prodCaseDescription = rsProdInvReport("prodCaseDescription")
						prodUnitUPC = rsProdInvReport("prodUnitUPC")
						prodCaseUPC = rsProdInvReport("prodCaseUPC")					

					%>

						<!-- line !-->
						<tr>

							<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
								<%= prodSKU %>
							</th>

							<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
								<%= prodDescription %>
							</th>
							
							<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
								<%= prodCategoryName %>
							</th>
							
							<th scope="col" style="font-size:12px; font-weight:normal;" align="center">
								<%= prodInventoriedItemDisplay %>
							</th>
							
							<th scope="col" style="font-size:12px; font-weight:normal;" align="center">
								<%= prodPickableItemDisplay %>
							</th>
							
							<th scope="col" style="font-size:12px; font-weight:normal;" align="center">
								<%= prodCasePricing %>
							</th>
							
							<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
								<%= prodCaseDescription %>
							</th>
							
							<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
								<%= prodUnitUPC %>
							</th>
							
							<th scope="col" style="font-size:12px; font-weight:normal;" align="left">
								<%= prodCaseUPC %>
							</th>
						</tr>
						<%rsProdInvReport.movenext
					Loop
				End If
				Set rsProdInvReport = Nothing
				cnnProdInvReport.Close
				Set cnnProdInvReport = Nothing
			%>
<!-- eof line !-->
			</tbody>
		</table>
		<!-- the table with statements ends here !-->
		

		
		 
		
		</td>
		</tr>
		</tbody>
		</table>
		<!-- main table ends here !-->
		
	</body>
	
</html>
