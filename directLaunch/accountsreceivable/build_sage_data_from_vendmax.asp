<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->

<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
 
<%

Default_GL_AR_Account = "12500"
Default_GL_Account = "41400"


'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page
'Usage = "http://{xxx}.{domain}.com/directLaunch/accountsreceivable/build_sage_data_from_vendmax.asp?runlevel=run_now
Server.ScriptTimeout = 90000

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)

If Request.QueryString("runlevel") <> "run_now" then
	Response.Write("Improper usage, no run level was specified in the query string")	
	response.end
End IF 

'This single page loops through and handles alerts for ALL clients
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1 AND Backend = 'Streamware'"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
	
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3


If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF
	
		ClientKey = TopRecordset.Fields("clientkey")
	
		'To begin with, see if this client uses the A/R module
		'If they don't then don't bother running for then
		
		Response.Write("Seeing if we need to run this for " & ClientKey & "<br>")
		
		If TopRecordset.Fields("arModule") = 1 Then
	
			'The IF statement below makes sure that when run from DEV it only deos client keys with a d
			'and when run from LIVE it only does client keys without a d
			'Pretty smart, huh
			
			If (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") = 0 AND Instr(ucase(ClientKey),"D") = 0)_
			or (Instr(ucase(Request.ServerVariables("SERVER_NAME")),"DEV") <> 0 AND Instr(ucase(ClientKey),"D") <> 0) Then

												
				Call SetClientCnnString
				
				Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
				

				'**************************************************************
				'Get next Entry Thread for use in the SC_AuditLogDLaunch table
				On Error Goto 0
				Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
				cnnAuditLog.open MUV_READ("ClientCnnString") 
				Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
				rsAuditLog.CursorLocation = 3 
				Set rsAuditLog = cnnAuditLog.Execute("Select TOP 1 * from SC_AuditLogDLaunch order by EntryThread desc")
				If Not rsAuditLog.EOF Then 
					If IsNull(rsAuditLog("EntryThread")) Then EntryThread =1 Else EntryThread = rsAuditLog("EntryThread") + 1
				Else
					EntryThread = 1
				End If
				set rsAuditLog = nothing
				cnnAuditLog.close
				set cnnAuditLog = nothing
					
	
					CreateAuditLogEntry "Build Sage Data From VendMax Launch","Build Sage Data From VendMax Launch","Minor",0,"Build Sage Data From VendMax Launch ran."					
	
					WriteResponse "<font color='purple' size='24'>Start processing " & ClientKey  & "</font><br>"
	
					
					If MUV_READ("cnnStatus") = "OK" Then ' else it loops
					
						
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
							
				
							Set cnnSageVmax = Server.CreateObject("ADODB.Connection")
							cnnSageVmax.open (Session("ClientCnnString"))
							Set rsSageVmax = Server.CreateObject("ADODB.Recordset")
							Set rsVmaxLocations = Server.CreateObject("ADODB.Recordset")
							Set rsInvoiceHistHeader  = Server.CreateObject("ADODB.Recordset")
							Set rsVmaxPointsOfSale = Server.CreateObject("ADODB.Recordset")
							Set rsInvoiceHistDetail  = Server.CreateObject("ADODB.Recordset")
							Set rsVmaxOrderH  = Server.CreateObject("ADODB.Recordset")
							Set rsVmaxOrderI  = Server.CreateObject("ADODB.Recordset")

							
							SQLSageVmax = "DELETE FROM AR_Customer"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)
							
							
							SQLSageVmax = "INSERT INTO AR_Customer (CustNum, Name, Addr1, Addr2, CityStateZip, Phone, Fax,  AcctStatus, "
							SQLSageVmax = SQLSageVmax & "City, [State], zip) "
							SQLSageVmax = SQLSageVmax & "SELECT cus_id, [description], addr2, addr3, city + ' ' + [state] + ' ' + zip, main_phone, "
							SQLSageVmax = SQLSageVmax & "fax, 'A', city, [state], zip "
							SQLSageVmax = SQLSageVmax & "FROM VMAX_Customers"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)
							
							
							SQLSageVmax = "DELETE FROM IC_Product"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)
							
							SQLSageVmax = "INSERT INTO IC_Product (prodSKU, prodDescription, prodCategory) "
							SQLSageVmax = SQLSageVmax & "SELECT pro_id, [description], cat "
							SQLSageVmax = SQLSageVmax & "FROM VMAX_products"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)
							
							
							SQLSageVmax = "DELETE FROM AR_CustomerShipTo"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)
															
							SQLSageVmax = "INSERT INTO AR_CustomerShipTo (CustNum, ShipName, Addr1, Addr2, City, [State], Zip, Phone, Fax, Email,BackendShipToIDIfApplicable) "
							SQLSageVmax = SQLSageVmax & "SELECT cus_id, addr1, addr2, addr3, city, [state], zip, main_phone, fax, e_mail, loc_id "
							SQLSageVmax = SQLSageVmax & "FROM VMAX_locations"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)


							SQLSageVmax = "DELETE FROM AR_CustomerBillTo"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)
															
							SQLSageVmax = "INSERT INTO AR_CustomerBillTo (CustNum, BillName, Addr1, Addr2, City, [State], Zip, Phone, Fax, Email,BackendBillToIDIfApplicable) "
							SQLSageVmax = SQLSageVmax & "SELECT cus_id, ma_addr1, ma_addr2, ma_addr3, ma_city, ma_state, ma_zip, main_phone, fax, e_mail, loc_id "
							SQLSageVmax = SQLSageVmax & "FROM VMAX_locations"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)


								
							'Update the customer mapping table based on data from the vendmax locations file
							'Get the Streamware partner record ID
							PartnerIntRecID = ""
							SQLSageVmax = "SELECT InternalRecordIdentifier FROM IC_Partners WHERE partnerCompanyName = 'Streamware'"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)
							
							If rsSageVmax.EOF Then
								WriteResponse "No IC_Partners record for Streamware<br><br>"
								Response.End
							Else
								PartnerIntRecID =rsSageVmax("InternalRecordIdentifier")
							End If
							
							
							'Delete the ones where Streamware is the recordSource
							SQLSageVmax = "DELETE FROM AR_CustomerMapping WHERE partnerRecID = " & PartnerIntRecID & " AND RecordSource <> 'Insight'"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)
							
							'Now put back any that don't already exist in the file
							'If they are already in there, Insight owns them
							SQLSageVmax = "INSERT INTO AR_CustomerMapping (partnerCustID,ourCustID,partnerRecID,partnerShipToID) "
							SQLSageVmax = SQLSageVmax & "SELECT user7, cus_id, " &  PartnerIntRecID & ", loc_id FROM VMAX_locations WHERE user7 IS NOT NULL AND "
							SQLSageVmax = SQLSageVmax & "loc_id NOT IN (SELECT ourCustID FROM AR_CustomerMapping WHERE partnerRecID = " & PartnerIntRecID & ")"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)


							'Begin the big multistep process of building the invoice history header file	
							SQLSageVmax = "DELETE FROM IN_InvoiceHistHeader"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)
							
							
							'Initially inserts it with the pos_id in the CustID field
							SQLSageVmax = "INSERT INTO IN_InvoiceHistHeader (InvoiceID, CustID, InvoiceCreationDate, InvoiceType) "
' Top 1000							
							SQLSageVmax = SQLSageVmax & "SELECT TOP 1000 transaction_id, pos_id, datetime_printed, 'Invoice' "
							SQLSageVmax = SQLSageVmax & "FROM VMAX_printed_invoices"
							SQLSageVmax = SQLSageVmax & " ORDER BY datetime_printed DESC"
							
							WriteResponse SQLSageVmax & "<br><br>"
							cnnSageVmax.CommandTimeout = 90000 
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)

							'Now change the pos_id to the loc_id
							SQLSageVmax = "SELECT CustId,InternalRecordIdentifier FROM IN_InvoiceHistHeader "
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax = cnnSageVmax.Execute(SQLSageVmax)
							
							If NOT rsSageVmax.EOF Then
								Do While NOT rsSageVmax.EOF
								
									CustID = NULL : ShipToID = NULL : BillToID = NULL
									HistHeaderIntRecId = rsSageVmax("InternalRecordIdentifier")
								
									SQLVmaxPointsOfSale = "SELECT * FROM VMAX_points_of_sale WHERE pos_id = " & rsSageVmax("CustID")
									Set rsVmaxPointsOfSale = cnnSageVmax.Execute(SQLVmaxPointsOfSale)
									
									If Not rsVmaxPointsOfSale.EOF Then
									
										SQLVmaxLocations = "SELECT * FROM VMAX_Locations WHERE loc_id = " & rsVmaxPointsOfSale("loc_id")
										Set rsVmaxLocations = cnnSageVmax.Execute(SQLVmaxLocations)
										

										If NOT rsVmaxLocations.EOF Then
											CustID = rsVmaxLocations("cus_id")
											ShipToID = rsVmaxLocations("loc_id")
											BillToID = rsVmaxLocations("loc_id")
										End If
									
									End If
									
								
									SQLInvoiceHistHeader = "UPDATE IN_InvoiceHistHeader SET "
									SQLInvoiceHistHeader = SQLInvoiceHistHeader & "CustID = " & CustID
									SQLInvoiceHistHeader = SQLInvoiceHistHeader & ", BackendShipToIDIfApplicable = " & ShipToID 
									SQLInvoiceHistHeader = SQLInvoiceHistHeader & ", BackendBillToIDIfApplicable = " & BillToID
									SQLInvoiceHistHeader = SQLInvoiceHistHeader & " WHERE InternalRecordIdentifier = " & HistHeaderIntRecId
									
									Response.Write("<br>" & SQLInvoiceHistHeader  & "<br>")
									
									Set rsInvoiceHistHeader = cnnSageVmax.Execute(SQLInvoiceHistHeader)
									
									rsSageVmax.movenext
									
								Loop 
								
							End If


							'Now update the Alternate Cust ID in the Invoice History Header from the Customer Mapping table
							
							SQLInvoiceHistHeader = "UPDATE IN_InvoiceHistHeader "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "SET IN_InvoiceHistHeader.AlternateCustID = AR_CustomerMapping.partnerCustID "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "FROM IN_InvoiceHistHeader INNER JOIN "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "AR_CustomerMapping ON IN_InvoiceHistHeader.custid = AR_CustomerMapping.ourcustID AND "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "IN_InvoiceHistHeader.BackendShipToIDIfApplicable = AR_CustomerMapping.partnerShipToID"
							
							Set rsInvoiceHistHeader = cnnSageVmax.Execute(SQLInvoiceHistHeader)
							

							' Set additional fields that we need to get from the vendmax orders table
							SQLInvoiceHistHeader = "UPDATE IN_InvoiceHistHeader SET "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "IN_InvoiceHistHeader.PONumber = VMAX_orders.po_number "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.Terms = VMAX_orders.terms "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.Orderdate = VMAX_orders.datetime_created "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "FROM IN_InvoiceHistHeader INNER JOIN "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "VMAX_orders ON IN_InvoiceHistHeader.InvoiceID = VMAX_orders.transaction_id "

							Set rsInvoiceHistHeader = cnnSageVmax.Execute(SQLInvoiceHistHeader)


							' Update Ship To Information
							SQLInvoiceHistHeader = "UPDATE IN_InvoiceHistHeader SET "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "IN_InvoiceHistHeader.ShipToName = AR_CustomerShipTo.ShipName "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.ShipToAddr1 = AR_CustomerShipTo.Addr1 "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.ShipToAddr2 = AR_CustomerShipTo.Addr2 "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.ShipToCity = AR_CustomerShipTo.City "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.ShipToState = AR_CustomerShipTo.State "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.ShipToPostalCode = AR_CustomerShipTo.Zip "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.ShipToContact = AR_CustomerShipTo.Contact "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.ShipToDescription = AR_CustomerShipTo.Description "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "FROM IN_InvoiceHistHeader INNER JOIN "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "AR_CustomerShipTo ON IN_InvoiceHistHeader.BackendShipToIDIfApplicable = AR_CustomerShipTo.BackendShipToIDIfApplicable"


							Set rsInvoiceHistHeader = cnnSageVmax.Execute(SQLInvoiceHistHeader)


							' Update Bill To Information
							SQLInvoiceHistHeader = "UPDATE IN_InvoiceHistHeader SET "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "IN_InvoiceHistHeader.BillToName = AR_CustomerBillTo.BillName "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.BillToAddr1 = AR_CustomerBillTo.Addr1 "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.BillToAddr2 = AR_CustomerBillTo.Addr2 "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.BillToCity = AR_CustomerBillTo.City "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.BillToState = AR_CustomerBillTo.State "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.BillToPostalCode = AR_CustomerBillTo.Zip "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.BillToContact = AR_CustomerBillTo.Contact "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",IN_InvoiceHistHeader.BillToDescription = AR_CustomerBillTo.Description "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "FROM IN_InvoiceHistHeader INNER JOIN "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "AR_CustomerBillTo ON IN_InvoiceHistHeader.BackendBillToIDIfApplicable = AR_CustomerBillTo.BackendBillToIDIfApplicable"


							Set rsInvoiceHistHeader = cnnSageVmax.Execute(SQLInvoiceHistHeader)

							
							' Any blank BillTo's are Same As Sold To, so copy in the ShipTo info
							SQLInvoiceHistHeader = "UPDATE IN_InvoiceHistHeader SET "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "BillToName = ShipToName "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",BillToAddr1 = ShipToAddr1 "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",BillToAddr2 = ShipToAddr2 "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",BillToCity = ShipToCity "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",BillToState = ShipToState "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",BillToPostalCode = ShipToPostalCode "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",BillToContact = ShipToContact "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & ",BillToDescription = ShipToDescription "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & " WHERE BillToName IS NULL AND BillToAddr1 IS NULL"

							
							Set rsInvoiceHistHeader = cnnSageVmax.Execute(SQLInvoiceHistHeader)


							' Fillin invoice due date
							SQLInvoiceHistHeader = "UPDATE IN_InvoiceHistHeader SET "
							SQLInvoiceHistHeader = SQLInvoiceHistHeader & "InvoiceDueDate = dateadd(d,30,InvoiceCreationDate) "
							
							Set rsInvoiceHistHeader = cnnSageVmax.Execute(SQLInvoiceHistHeader)
							
 

'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
' N O W		D O		L I N E		I T E M S 
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''

							'Begin the big multistep process of building the invoice history item file	
							SQLSageVmax = "DELETE FROM IN_InvoiceHistDetail"
							WriteResponse SQLSageVmax & "<br><br>"
							Set rsSageVmax= cnnSageVmax.Execute(SQLSageVmax)


							SQLInvoiceHistHeader= "SELECT * FROM IN_InvoiceHistHeader "
							WriteResponse SQLInvoiceHistHeader & "<br><br>"
							Set rsInvoiceHistHeader  = cnnSageVmax.Execute(SQLInvoiceHistHeader)
							
							If NOT rsInvoiceHistHeader.EOF Then
								Do While NOT rsInvoiceHistHeader.EOF
								
''''This is where I left off					
								'Lookup this invoice in the vmax order table to get the orderid
								SQLVmaxOrderH  = "SELECT ord_id FROM VMAX_orders WHERE transaction_id = '" & rsInvoiceHistHeader("InvoiceID") & "'"
								WriteResponse SQLVmaxOrderH  & "<br><br>"
								Set rsVmaxOrderH  = cnnSageVmax.Execute(SQLVmaxOrderH)
								
									If Not rsVmaxOrderH.EOF Then
									
										vmax_ord_id = 0
										vmax_ord_id = rsVmaxOrderH("ord_id")
	
										' That got us the ord_id, now insert all the line items into the IN_InvoiceHistDetailTable
										SQLVmaxOrderI  = "INSERT INTO IN_InvoiceHistDetail ( "
										SQLVmaxOrderI  = SQLVmaxOrderI  & "InvoiceID, CustID, InvoiceCreationDate, LineNumber, prodSKU, "
										SQLVmaxOrderI  = SQLVmaxOrderI  & "PricePerUnitSold, QtyOrdered, QtyShipped, Taxable, ThisLineNotAProduct, TotalTaxForLine "
										SQLVmaxOrderI  = SQLVmaxOrderI  & ") "
										SQLVmaxOrderI  = SQLVmaxOrderI  & " SELECT '" & rsInvoiceHistHeader("InvoiceID") & "','" & rsInvoiceHistHeader("CustID") & "', "
										SQLVmaxOrderI  = SQLVmaxOrderI  &	"'" & rsInvoiceHistHeader("InvoiceCreationDate") & "', "
										SQLVmaxOrderI  = SQLVmaxOrderI  &	"line_num, "
										SQLVmaxOrderI  = SQLVmaxOrderI  &	"pkp_id, "
										SQLVmaxOrderI  = SQLVmaxOrderI  &	"price, "		
										SQLVmaxOrderI  = SQLVmaxOrderI  &	"quantity, "
										SQLVmaxOrderI  = SQLVmaxOrderI  &	"quantity, "																									
										SQLVmaxOrderI  = SQLVmaxOrderI  &	"CASE WHEN tax_rate = 0 THEN 0 ELSE 1 END, "
										SQLVmaxOrderI  = SQLVmaxOrderI  &	"0, "
										SQLVmaxOrderI  = SQLVmaxOrderI  &	"total_tax "
										SQLVmaxOrderI  = SQLVmaxOrderI  &	" FROM VMAX_order_items WHERE VMAX_order_items.ord_id = '" & vmax_ord_id & "'"
										
										WriteResponse SQLVmaxOrderI & "<br><br>"
										
										Set rsVmaxOrderH  = cnnSageVmax.Execute(SQLVmaxOrderI)

									End If

									'INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  
									'INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  
									'Now lookup additional line items that may be in the INVOICE ITEMS table from Vendmax.
									'INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  
									'INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  INVOICE ITEMS  									

									SQLVmaxOrderH  = "SELECT ord_id FROM VMAX_orders WHERE transaction_id = '" & rsInvoiceHistHeader("InvoiceID") & "'"
									WriteResponse SQLVmaxOrderH  & "<br><br>"
									Set rsVmaxOrderH  = cnnSageVmax.Execute(SQLVmaxOrderH)
									
										If Not rsVmaxOrderH.EOF Then
										
											vmax_ord_id = 0
											vmax_ord_id = rsVmaxOrderH("ord_id")
		
											' That got us the ord_id, now insert all the line items into the IN_InvoiceHostDetailTable
											SQLVmaxOrderI  = "INSERT INTO IN_InvoiceHistDetail ( "
											SQLVmaxOrderI  = SQLVmaxOrderI  & "InvoiceID, CustID, InvoiceCreationDate, LineNumber, prodSKU, "
											SQLVmaxOrderI  = SQLVmaxOrderI  & "PricePerUnitSold, QtyOrdered, QtyShipped, Taxable, ThisLineNotAProduct "
											SQLVmaxOrderI  = SQLVmaxOrderI  & ") "
											SQLVmaxOrderI  = SQLVmaxOrderI  & " SELECT '" & rsInvoiceHistHeader("InvoiceID") & "','" & rsInvoiceHistHeader("CustID") & "', "
											SQLVmaxOrderI  = SQLVmaxOrderI  &	"'" & rsInvoiceHistHeader("InvoiceCreationDate") & "', "
											SQLVmaxOrderI  = SQLVmaxOrderI  &	"line_num, "
											SQLVmaxOrderI  = SQLVmaxOrderI  &	"pkp_id, "
											SQLVmaxOrderI  = SQLVmaxOrderI  &	"price, "		
											SQLVmaxOrderI  = SQLVmaxOrderI  &	"quantity, "
											SQLVmaxOrderI  = SQLVmaxOrderI  &	"quantity, "																									
											SQLVmaxOrderI  = SQLVmaxOrderI  &	"CASE WHEN tax_rate = 0 THEN 0 ELSE 1 END, "
											SQLVmaxOrderI  = SQLVmaxOrderI  &	"0 "
											SQLVmaxOrderI  = SQLVmaxOrderI  &	" FROM VMAX_order_items WHERE VMAX_order_items.ord_id = '" & vmax_ord_id & "'"
											
											'WriteResponse SQLVmaxOrderI & "<br><br>"
											
											'Set rsVmaxOrderH  = cnnSageVmax.Execute(SQLVmaxOrderI)
	
										End If

																	
									rsInvoiceHistHeader.movenext
									
								Loop 
								


								
								
								
							End If
							

							
							
							
							
							
							
							
							
							
							
							
							' Take pkp_id and find it in Vmax_packaged_products and change it to pro_id
							SQLVmaxOrderI  = "UPDATE IN_InvoiceHistDetail SET "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "IN_InvoiceHistDetail.prodSKU = VMAX_packaged_products.pro_id "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "FROM IN_InvoiceHistDetail INNER JOIN "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "VMAX_packaged_products ON IN_InvoiceHistDetail.prodSKU = VMAX_packaged_products.pkp_id"
							Set rsVmaxOrderH  = cnnSageVmax.Execute(SQLVmaxOrderI)

							
							' Now use the pro_id to lookup the actual sku in VMAX_Products
							SQLVmaxOrderI  = "UPDATE IN_InvoiceHistDetail SET "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "IN_InvoiceHistDetail.prodSKU = VMAX_Products.code "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "FROM IN_InvoiceHistDetail INNER JOIN "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "VMAX_Products ON IN_InvoiceHistDetail.prodSKU = VMAX_Products.pro_id"
							Set rsVmaxOrderH  = cnnSageVmax.Execute(SQLVmaxOrderI)

							
						
							' Now update all the product descriptions
							
							SQLVmaxOrderI  = "UPDATE IN_InvoiceHistDetail SET "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "IN_InvoiceHistDetail.prodDescription = VMAX_products.Description "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "FROM IN_InvoiceHistDetail INNER JOIN "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "VMAX_products ON IN_InvoiceHistDetail.prodSKU = VMAX_products.code"

							Set rsVmaxOrderH  = cnnSageVmax.Execute(SQLVmaxOrderI)

							
							'Update the GL Account Numbers to default values
							SQLVmaxOrderI  = "UPDATE IN_InvoiceHistDetail SET "
							SQLVmaxOrderI  = SQLVmaxOrderI  & " GL_AR_Account = '" & Default_GL_AR_Account & "'"
							SQLVmaxOrderI  = SQLVmaxOrderI  & ",GL_Account = '" & Default_GL_Account & "'"
							Set rsVmaxOrderH  = cnnSageVmax.Execute(SQLVmaxOrderI)
							
							'Now redo GL Account based in user field 5 from the product file
							SQLVmaxOrderI  = "UPDATE IN_InvoiceHistDetail SET "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "IN_InvoiceHistDetail.GL_AR_Account = VMAX_Products.user5 "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "FROM IN_InvoiceHistDetail INNER JOIN "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "VMAX_Products ON IN_InvoiceHistDetail.prodSKU = Cast(VMAX_Products.pro_id as varchar(255)) "
							SQLVmaxOrderI  = SQLVmaxOrderI  & "WHERE user5 IS NOT NULL AND user5 <> '" & Default_GL_Account & "'"

							Set rsVmaxOrderH  = cnnSageVmax.Execute(SQLVmaxOrderI)

							
							

							On error resume next
							cnnSageVmax.Close
							Set rsSageVmax = Nothing
							Set rsVmaxLocations = Nothing
							Set rsVmaxPointsOfSale  = Nothing
							Set rsInvoiceHistHeader = Nothing
							Set cnnSageVmax = Nothing



''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
''''''''''''''''''
						
	
										
						WriteResponse ("******** DONE Processing " & ClientKey  & "************<br>")
				
				End If
			
		End If	
		
	Else ' is the ar  module enabled
	
		Call SetClientCnnString
				
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
			
		'Get next Entry Thread for use in the SC_AuditLogDLaunch table
		On Error Goto 0
		Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
		cnnAuditLog.open MUV_READ("ClientCnnString") 
		Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
		rsAuditLog.CursorLocation = 3 
		Set rsAuditLog = cnnAuditLog.Execute("Select TOP 1 * from SC_AuditLogDLaunch order by EntryThread desc")
		If Not rsAuditLog.EOF Then 
		If IsNull(rsAuditLog("EntryThread")) Then EntryThread =1 Else EntryThread = rsAuditLog("EntryThread") + 1
		Else
		EntryThread = 1
		End If
		set rsAuditLog = nothing
		cnnAuditLog.close
		set cnnAuditLog = nothing


		WriteResponse ("Skipping the client " & ClientKey & " because the ar module is not enabled.<BR>")
		
	End If ' is the Service  module enabled
	
	TopRecordset.movenext
	
	Loop
	
	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")

'************************************************************************************
'************************************************************************************
'Subs and funcs begin here
'************************************************************************************

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
		Session("SQL_Owner") = Recordset.Fields("dbLogin")
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub


Sub WriteResponse(passedLogEntry)

	response.write(Now() & "&nbsp;&nbsp;&nbsp;" & passedLogEntry)
	
	passedLogEntry = Replace(passedLogEntry,"'","''")
	
	SQL = "INSERT INTO SC_AuditLogDLaunch (EntryThread, DirectLaunchName, DirectLaunchFile, LogEntry)"
	SQL = SQL &  " VALUES (" & EntryThread & ""
	SQL = SQL & ",'MCS Rebuild Helper'"
	SQL = SQL & ",'/directlaunch/bizintel/mcs_rebuild_helper_launch.asp'"
	SQL = SQL & ",'"  & passedLogEntry & "'"
	SQL = SQL & ")"
	
	'Response.write("<BR>" & SQL & "<BR>")
	
	Set cnnAuditLog = Server.CreateObject("ADODB.Connection")
	cnnAuditLog.open Session("ClientCnnString") 
	Set rsAuditLog = Server.CreateObject("ADODB.Recordset")
	rsAuditLog.CursorLocation = 3 
	
	Set rsAuditLog = cnnAuditLog.Execute(SQL)

	set rsAuditLog = nothing
	cnnAuditLog.close
	set cnnAuditLog = nothing

End Sub


Sub DelayResponse(numberOfseconds)
 Dim WshShell
 Set WshShell=Server.CreateObject("WScript.Shell")
 WshShell.Run "waitfor /T " & numberOfSecond & "SignalThatWontHappen", , True
End Sub

Function CheckRemoteURL(fileURL)
    ON ERROR RESUME NEXT
    Dim xmlhttp

    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

    xmlhttp.open "GET", fileURL, False
    xmlhttp.send
    If(Err.Number<>0) then
        Response.Write "Could not connect to remote server"
    else
        Select Case Cint(xmlhttp.status)
            Case 200, 202, 302
                Set xmlhttp = Nothing
                CheckRemoteURL = True
            Case Else
                Set xmlhttp = Nothing
                CheckRemoteURL = False
        End Select
    end if
    ON ERROR GOTO 0
End Function


'************************************************************************************
'************************************************************************************
'Subs and funcs end here
'************************************************************************************


%>