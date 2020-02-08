<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->

<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
<%
Response.Write("Start Page:" & Now() & "<br>")
Server.ScriptTimeout = 25000

'Build webfulfillment page
'Designed to be launched via a scheduled process
'Self contained page will connect to the OCSAccess web server & build the Insight webfulfillment table
'Usage = "http://{xxx}.{domain}.com/directLaunch/invoicing/BuildWebFulfillmentTable.asp?runlevel=run_now

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)
If Request.QueryString("runlevel") <> "run_now" then response.end

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles alerts for ALL clients
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and exit
If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF

		PassPhrase = TopRecordset.Fields("directLaunchPassphrase")
		ClientKey = TopRecordset.Fields("clientkey")

		Response.Write("******** Processing " & ClientKey  & "************<br>")
		Call SetClientCnnString ' Also gets the OCSAccess SQL values
		
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
		
		If MUV_READ("cnnStatus") = "OK" and MUV_READ("webFulfillmentModule") = "Enabled" Then ' dev sites ARE included in this build
		

			' Connect to OCSAccess and build a temporary table on insight containing all the order ids
			

			
			'To make it quicker, only get the last 45 days
			SQLOCSAccess = "SELECT OrderID, CustID, OrderDate, merchTotal, Comments FROM tblOrders WHERE  (OrderDate > DateAdd(day,-45,getdate())) "
			SQLOCSAccess = SQLOCSAccess & " ORDER BY OrderID"

			Set cnnOCSAccess = Server.CreateObject("ADODB.Connection")
			cnnOCSAccess.open (MUV_READ("ClientCnnString_OCSAccess"))
			Set rsOCSAccess = Server.CreateObject("ADODB.Recordset")
			rsOCSAccess.CursorLocation = 3 
			
			Set rsOCSAccess = cnnOCSAccess.Execute(SQLOCSAccess)

			If not rsOCSAccess.EOF Then
			
				Set cnnWebFulfillment = Server.CreateObject("ADODB.Connection")
				cnnWebFulfillment.open (Session("ClientCnnString"))
				Set rsWebFulfillment = Server.CreateObject("ADODB.Recordset")
				rsWebFulfillment.CursorLocation = 3 

				' The table might not be there, so create it but use an error branch in case it's there already
				On Error Resume Next
				SQLWebFulfillment = "CREATE TABLE zIN_WebFulfillment_0 ([OrderID] [varchar](50) NULL, [OrderDate] [datetime] NULL, [OrderAmount] [money] NULL, [CustID] [varchar](50) NULL, [OrderComments] [varchar](8000) NULL, [CustTypeNum] int NULL)" ' Uses 0 becuase technically there is no user number
				Set rsWebFulfillment = cnnWebFulfillment.Execute(SQLWebFulfillment)
				On Error Goto 0


				'Delete all fron the insight work table
				SQLWebFulfillment = "DELETE FROM zIN_WebFulfillment_0" ' Uses 0 becuase technically there is no user number
				Set rsWebFulfillment = cnnWebFulfillment.Execute(SQLWebFulfillment)
				
				'Now fill 'er up

				Do
				
					SQLWebFulfillment = "INSERT INTO zIN_WebFulfillment_0 (OrderID,OrderDate,OrderAmount,CustID,OrderComments) VALUES "
					SQLWebFulfillment = SQLWebFulfillment + "('" & rsOCSAccess("OrderID") & "','" & rsOCSAccess("OrderDate") & "'," & rsOCSAccess("merchTotal") & ",'" & rsOCSAccess("CustID") & "' "
					SQLWebFulfillment = SQLWebFulfillment + ",'" & rsOCSAccess("Comments") & "')"
					
					Set rsWebFulfillment = cnnWebFulfillment.Execute(SQLWebFulfillment)
					
					rsOCSAccess.Movenext
				Loop Until rsOCSAccess.EOF

			
				'The table now holds ALL OrderIDs, so cut it back only to include those we
				'don't have in the IN_WebFulfillment table yet
				
				SQLWebFulfillment = "DELETE FROM zIN_WebFulfillment_0 WHERE OrderID IN (SELECT OCSAccessOrderID FROM IN_WebFulfillment)"
				Set rsWebFulfillment = cnnWebFulfillment.Execute(SQLWebFulfillment)
	

				'Now begin filling in the missing records
				
				Set rsTelSel = Server.CreateObject("ADODB.Recordset")
				rsTelSel.CursorLocation = 3 
				Set rsWebFulfillmentForUpdating = Server.CreateObject("ADODB.Recordset")
				rsWebFulfillmentForUpdating.CursorLocation = 3 
	
	
				SQLWebFulfillment = "SELECT * FROM zIN_WebFulfillment_0 ORDER BY OrderID"
				Set rsWebFulfillment = cnnWebFulfillment.Execute(SQLWebFulfillment)
	
				If Not rsWebFulfillment.EOF Then

					Do While not rsWebFulfillment.EOF
					
						OCSAccessOrderID = "" : OCSAccessOrderDate = "" : CustID = "" : CustClassCode = "" : CustTypeNum = "" : MDSInvoiceID = "" 
						MDSInvoiceDate = "" : OCSAccessMerchTotal = "" : MDSInvoiceTotal = 0 : OCSAccessOrderComments = ""
						
		
						OCSAccessOrderID = rsWebFulfillment("OrderID")
						OCSAccessOrderDate = rsWebFulfillment("OrderDate")
						OCSAccessMerchTotal  = rsWebFulfillment("OrderAmount")
						OCSAccessOrderComments = rsWebFulfillment("OrderComments")
						CustID = rsWebFulfillment("CustID")
						CustClassCode = GetCustClassByCustID(CustID)
						CustTypeNum = GetCustTypeCodeByCustID(CustID)
						
						'Now try to find the info in TelSel
						SQLTelSel = "SELECT InvoiceNo, InvoiceDate, InvoiceTotaL - (GstTax + SalesTaxCharge + Deposit) As InvoiceTotal from TelSel WHERE InvoiceComment3 = 'Order #  " & Trim(OCSAccessOrderID) & "' "
						SQLTelSel = SQLTelSel & " AND InvoiceNo <> '***DELETE'"
						Set rsTelSel = cnnWebFulfillment.Execute(SQLTelSel)
						
						If Not rsTelSel.EOF Then
							If rsTelSel("InvoiceNo") <> "***DELETE" Then
							MDSInvoiceID = rsTelSel("InvoiceNo")
							MDSInvoiceDate = rsTelSel("InvoiceDate")
							MDSInvoiceTotal = rsTelSel("InvoiceTotal")
							End If
						End If


						If NOT IsNumeric(CustTypeNum) Then CustTypeNum = 0

	
						' Insert into the live IN_WebFulfillment tbale
						SQLWebFulfillmentForUpdating = "INSERT INTO IN_WebFulfillment (OCSAccessOrderID, OCSAccessOrderDate, CustID, CustClassCode, CustTypeNum, MDSInvoiceID, MDSInvoiceDate, OCSAccessMerchTotal, MDSInvoiceTotal,OCSAccessOrderComments) "
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + " VALUES ("
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + "'" & OCSAccessOrderID & "'"
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + ",'" & OCSAccessOrderDate & "'"
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + ",'" & CustID & "'"
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + ",'" & CustClassCode & "'"
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + "," & CustTypeNum
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + ",'" & MDSInvoiceID & "'"
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + ",'" & MDSInvoiceDate & "'"
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + "," & OCSAccessMerchTotal  
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + "," & MDSInvoiceTotal 
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + ", CustTypeNum = " & CustTypeNum & " "						
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + ",'" & OCSAccessOrderComments & "'"
						SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + ")"	
						
						'Response.Write("<BR>" & SQLWebFulfillmentForUpdating )
						
						Set rsWebFulfillmentForUpdating = cnnWebFulfillment.Execute(SQLWebFulfillmentForUpdating)
					
						rsWebFulfillment.MoveNext
					Loop
							
				End If
				
				Set rsWebFulfillmentForUpdating = Nothing
				Set rsTelSel = Nothing
				Set rsWebFulfillment = Nothing
				cnnWebFulfillment.Close
				Set cnnWebFulfillment = Nothing
			
			End If
			
			Set rsOCSAccess = Nothing
			cnnOCSAccess.Close
			Set cnnOCSAccess = Nothing

			
			'We are done with the first part, which is to insert all the new order data
			'Now we have to go through the second pass which is to see if any orders
			'already in the file have now been invoiced by Metroplex

			Response.Write("SECOND PASS***<BR>")
			
			Set cnnWebFulfillment = Server.CreateObject("ADODB.Connection")
			cnnWebFulfillment.open (Session("ClientCnnString"))
			Set rsWebFulfillment = Server.CreateObject("ADODB.Recordset")
			rsWebFulfillment.CursorLocation = 3 

			SQLWebFulfillment = "SELECT * FROM IN_WebFulFillment WHERE MDSInvoiceID = ''"
			
			Set rsWebFulfillment = cnnWebFulfillment.Execute(SQLWebFulfillment)
			
			If Not rsWebFulfillment.EOF Then
			
				Set rsTelSel = Server.CreateObject("ADODB.Recordset")
				rsTelSel.CursorLocation = 3 
				Set rsWebFulfillmentForUpdating = Server.CreateObject("ADODB.Recordset")
				rsWebFulfillmentForUpdating.CursorLocation = 3 
				

				Do While Not rsWebFulfillment.EOF
				
					'For each of the orders with blank info, see if it have now been invoiced
					
					MDSInvoiceID = "" : MDSInvoiceDate = "" : MDSInvoiceTotal = 0 
	
					OCSAccessOrderID = rsWebFulfillment("OCSAccessOrderID")
					
					'Now try to find the info in TelSel
					SQLTelSel = "SELECT InvoiceNo, InvoiceDate, InvoiceTotaL - (GstTax + SalesTaxCharge + Deposit) As InvoiceTotal, CustNum AS CustID from TelSel WHERE InvoiceComment3 = 'Order #  " & Trim(OCSAccessOrderID) & "'"
					SQLTelSel = SQLTelSel & " AND InvoiceNo <> '***DELETE'"

					Set rsTelSel = cnnWebFulfillment.Execute(SQLTelSel)

					If Not rsTelSel.EOF Then
						If rsTelSel("InvoiceNo") <> "***DELETE" Then
						
							MDSInvoiceID = rsTelSel("InvoiceNo")
							MDSInvoiceDate = rsTelSel("InvoiceDate")
							MDSInvoiceTotal = rsTelSel("InvoiceTotal")
							CustTypeNum = GetCustTypeCodeByCustID(rsTelSel("CustID"))
							'Response.Write("<BR>" & CustTypeNum )
							'Response.Write("<BR>" & rsTelSel("CustID"))
	
							' Insert into the live IN_WebFulfillment tbale
							SQLWebFulfillmentForUpdating = "UPDATE IN_WebFulfillment "
							SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + " SET "
							SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + " MDSInvoiceID = '" & MDSInvoiceID & "'"
							SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + ", MDSInvoiceDate = '" & MDSInvoiceDate &  "'"
							SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + ", MDSInvoiceTotal = " & MDSInvoiceTotal & " "
							SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + " WHERE InternalRecordIdentifier = " & rsWebFulfillment("InternalRecordIdentifier")						
							
							'Response.Write("<BR>" & SQLWebFulfillmentForUpdating)
							
							Set rsWebFulfillmentForUpdating = cnnWebFulfillment.Execute(SQLWebFulfillmentForUpdating)
							
						End If
					End If

					rsWebFulfillment.MoveNext
				Loop
				
				Set rsTelSel = Nothing
				Set rsWebFulfillmentForUpdating = Nothing
			
			End If
			
			
			' New third part - just set all the NULL in DontIncludeOnreport to 0
			SQLWebFulfillmentForUpdating = "UPDATE IN_WebFulfillment "
			SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + " SET "
			SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + "DontIncludeOnReport = 0 "
			SQLWebFulfillmentForUpdating = SQLWebFulfillmentForUpdating + " WHERE DontIncludeOnReport IS Null"
			
			Set rsWebFulfillmentForUpdating = cnnWebFulfillment.Execute(SQLWebFulfillmentForUpdating)
		
			
			Set rsWebFulfillment = Nothing
			cnnWebFulfillment.Close
			Set cnnWebFulfillment = Nothing
			

		'******************************************	
		Response.Write("******** DONE Processing " & ClientKey  & "************<br>")
	End If				
	TopRecordset.movenext
	
	Loop

	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If
Response.Write("End Page:" & Now() & "<br>")
Response.write("<script type=""text/javascript"">closeme();</script>")	
'Response.End
'*************************
'*************************
'Subs and funcs begin here


Sub SetClientCnnString

	dummy=MUV_WRITE("cnnStatus","")

	SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString
	
	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	
	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and exit
	If Recordset.recordcount <= 0 then
		Recordset.close
		Connection.close
		set Recordset=nothing
		set Connection=nothing
	Else
		ClientCnnString = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
		ClientCnnString = ClientCnnString & ";Database=" & Recordset.Fields("dbCatalog")
		ClientCnnString = ClientCnnString & ";Uid=" & Recordset.Fields("dbLogin")
		ClientCnnString = ClientCnnString & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
		dummy = MUV_Write("ClientCnnString",ClientCnnString)
		dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		dummy = MUV_Write("webFulfillmentModule",Recordset.Fields("webFulfillmentModule"))
		
		ClientCnnString_OCSAccess = "Driver={SQL Server};Server=" & Recordset.Fields("OCSAccess_dbServer")
		ClientCnnString_OCSAccess= ClientCnnString_OCSAccess& ";Database=" & Recordset.Fields("OCSAccess_dbCatalog")
		ClientCnnString_OCSAccess= ClientCnnString_OCSAccess& ";Uid=" & Recordset.Fields("OCSAccess_dbLogin")
		ClientCnnString_OCSAccess= ClientCnnString_OCSAccess& ";Pwd=" & Recordset.Fields("OCSAccess_dbPassword") & ";"
		dummy = MUV_Write("ClientCnnString_OCSAccess",ClientCnnString_OCSAccess)
		dummy = MUV_Write("SQL_Owner_OCSAccess",Recordset.Fields("OCSAccess_dbLogin"))

		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub



%>