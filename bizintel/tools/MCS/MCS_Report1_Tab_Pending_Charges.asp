<!-- row !-->
<div class="row" style="width:98%; margin-left:10px; margin-top:20px;">

<div class="container-fluid">
    <div class="row">
    
   	   <button type="button" class="generate-pdf" id="btnGeneratePDFPendingCharges"><i class="fas fa-file-pdf"></i>&nbsp;Generate PDF</button>
   	   
       <table id="tableSuperSumPendingCharges" class="display compact" style="width:100%;">
              <thead>
                  <tr>	
            		<th class="td-align1 gen-info-header" colspan="3" style="border-right: 2px solid #555 !important;">General</th>
					<th class="td-align1 vpc-current-header" colspan="2" style="border-right: 2px solid #555 !important;">Action</th>
				</tr>
				
                <tr>
					<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>Acct</th>
					<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Client</th>
					<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Primary<br> Slsmn</th>
					<th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">Action Notes</th>
					<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important; border-right: 2px solid #555 !important;" id="salesColumn">Pending Charges</th>
                </tr>
          </thead>

		<tbody>
		<%		


			SQL = "SELECT BI_MCSActions.CustID, BI_MCSActions.ActionNotes, AR_Customer.Salesman "
			SQL = SQL & " FROM  BI_MCSActions INNER JOIN "
			SQL = SQL & "  AR_Customer ON BI_MCSActions.CustID = AR_Customer.CustNum "
			SQL = SQL & " WHERE      (BI_MCSActions.MCSMonth = "
			SQL = SQL & " (SELECT      TOP (1) MCSMonth "
			SQL = SQL & "  FROM         BI_MCSActions AS BI_MCSActions_1 "
			SQL = SQL & "  ORDER BY InternalRecordIdentifier DESC)) AND (BI_MCSActions.Action LIKE '%invoice%') AND (YEAR(BI_MCSActions.RecordCreationDateTime) = "
			SQL = SQL & " (SELECT      TOP (1) YEAR(RecordCreationDateTime) AS Expr1 "
			SQL = SQL & "  FROM         BI_MCSActions AS BI_MCSActions_1 "
			SQL = SQL & "  WHERE      (Action LIKE '%invoice%') "
			SQL = SQL & "  ORDER BY InternalRecordIdentifier DESC)) "
			SQL = SQL & " ORDER BY AR_Customer.Salesman "
			
			Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.open (Session("ClientCnnString"))
			
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3
			Set rs = cnn8.Execute(SQL)
		
			If Not rs.Eof Then
					
				Do While Not rs.EOF
		
					PrimarySalesMan =  ""
					SelectedCustomerID = rs("CustID")
					CustName = GetCustNameByCustNum(rs("CustID"))
					
					PrimarySalesMan = rs("Salesman")
					
					If PrimarySalesMan <> "" Then
						PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(PrimarySalesMan)
					Else
						PrimarySalesPerson = ""
					End If
					
					ActionNotes = rs("ActionNotes")

	
					'***************************************************************
					'Get Current Pending Charges For This Account
					'***************************************************************
					
					SQLPendingLVF = "SELECT * FROM AR_Customer INNER JOIN BI_MCSData ON BI_MCSData.CustID = AR_Customer.CustNum WHERE AR_Customer.CustNum = " & SelectedCustomerID 
				
					Set cnnPendingLVF = Server.CreateObject("ADODB.Connection")
					cnnPendingLVF.open (Session("ClientCnnString"))
					
					Set rsPendingLVF = Server.CreateObject("ADODB.Recordset")
					rsPendingLVF.CursorLocation = 3
					Set rsPendingLVF = cnnPendingLVF.Execute(SQLPendingLVF)
				
					If Not rsPendingLVF.EOF Then
						PendingLVFHolder = rsPendingLVF("PendingLVF")
					Else
						PendingLVFHolder = 0
					End If
					
					PendingLVFHolder = FormatCurrency(PendingLVFHolder,2)
					'***************************************************************

					Response.Write("<tr id=""PENDINGCUST" & SelectedCustomerID & """")
	
					Response.Write(">")
				    Response.Write("<td class='smaller-detail-line'><a href='../CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>" & SelectedCustomerID  & "</a></td>")
				    Response.Write("<td class='smaller-detail-line'><a href='../CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>" & CustName & "</a></td>")	

				    If Instr(PrimarySalesPerson ," ") <> 0 Then
						Response.Write("<td class='smaller-detail-line'>" & Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")+1) & "</td>")
					Else
						Response.Write("<td class='smaller-detail-line'>" & PrimarySalesPerson & "</td>")
					End If
					
					
					Response.Write("<td align='left' class='smaller-detail-line'>" & ActionNotes & "</td>")
					Response.Write("<td align='left' class='smaller-detail-line' style='border-right: 2px solid #555 !important; color:red;'>" & PendingLVFHolder & "</td>")	
	
	
	
				    Response.Write("</tr>")
					    
					rs.movenext
						
				Loop
				
				Response.Write("</tbody>")
				Response.Write("</table>")		
				Response.Write("</div>")
		Else
		
			Response.Write("Nothing To Report")
		End If
		
		%>
		
		
            </table>
    </div>
         
          </div>
<!-- eof responsive tables !-->



<!-- eof row !-->

<!-- row !-->
<div class="row">
   <%
'Response.Write("<div class='col-lg-12'><h3>" & "Total Customers Listed:" & TotalCustsReported  & "</h3></div>")
%>

<div class="col-lg-12"><hr></div>
</div>
<!-- eof row !-->
