<!--#include file="../../inc/InsightFuncs_AR_AP.asp"-->
<%
If Session("Userno") = "" Then
	Response.End
Else
        InternalRecordIdentifier = Request.QueryString("i")
        PartnerInternalRecordIdentifier = Request.QueryString("i")
        FirstLetter = Request.QueryString("letter")
        currentPage=Int(Request.QueryString("page"))
        filterdata=Request.QueryString("filterdata")
        rowPerPage=Int(Request.QueryString("pagesize"))
        IF LCASE(FirstLetter)="all" OR LEN(FirstLetter)=0 THEN
            SQL9 = "SELECT COUNT(CustNum) as TotalCustomerCount FROM AR_Customer WHERE AcctStatus = 'A'" 
			Set cnn9 = Server.CreateObject("ADODB.Connection")
			cnn9.open (Session("ClientCnnString"))
			Set rs9 = Server.CreateObject("ADODB.Recordset")
			rs9.CursorLocation = 3 
			Set rs9 = cnn9.Execute(SQL9)
			If not rs9.EOF Then
				TotalCustomerCount = rs9("TotalCustomerCount")
			Else
				TotalCustomerCount = 0
			End If
			
			SQL9 = "SELECT COUNT(partnerCustID) as TotalEquivalentPartnerCustCount FROM AR_CustomerMapping WHERE partnerRecID = " & InternalRecordIdentifier
			Set rs9 = cnn9.Execute(SQL9)
			If not rs9.EOF Then
				TotalEquivalentPartnerCustCount = rs9("TotalEquivalentPartnerCustCount")
			Else
				TotalEquivalentPartnerCustCount = 0
			End If
            rs9.close
			set rs9 = Nothing
			cnn9.close
			set cnn9 = Nothing
            ELSE
                    SQL9 = "SELECT COUNT(CustNum) as TotalCustomerCount FROM AR_Customer WHERE LEFT(Name,1) = '" & FirstLetter & "' AND AcctStatus = 'A'"
					Set cnn9 = Server.CreateObject("ADODB.Connection")
					cnn9.open (Session("ClientCnnString"))
					Set rs9 = Server.CreateObject("ADODB.Recordset")
					rs9.CursorLocation = 3 
					Set rs9 = cnn9.Execute(SQL9)
					If not rs9.EOF Then
						TotalCustomerCount = rs9("TotalCustomerCount")
					Else
						TotalCustomerCount = 0
					End If
					
					SQL9 = "SELECT COUNT(partnerCustID) as TotalEquivalentPartnerCustCount, "
					SQL9 = SQL9 & " Count(ourCustID) as OurCustCOunt FROM AR_CustomerMapping where ourCustID in "
					SQL9 = SQL9 & " (select custnum from  AR_Customer "
					SQL9 = SQL9 & " WHERE LEFT(AR_Customer.Name,1) = '" & FirstLetter & "' AND AR_Customer.AcctStatus = 'A') "
										
					Set rs9 = cnn9.Execute(SQL9)
					'Response.Write(SQL9)
					If not rs9.EOF Then
						TotalEquivalentPartnerCustCount = rs9("TotalEquivalentPartnerCustCount")
					Else
						TotalEquivalentPartnerCustCount = 0
					End If
                    rs9.close
                    set rs9 = Nothing
			        cnn9.close
			        set cnn9 = Nothing
        END IF
END IF

%>
<%
	Set cnnCustomerTable = Server.CreateObject("ADODB.Connection")
	cnnCustomerTable.open (Session("ClientCnnString"))
	Set rsCustomerTable = Server.CreateObject("ADODB.Recordset")
	rsCustomerTable.CursorLocation = 3 
	
	Set cnnEquivalentCustomers = Server.CreateObject("ADODB.Connection")
	cnnEquivalentCustomers.open (Session("ClientCnnString"))
	Set rsEquivalentCustomers = Server.CreateObject("ADODB.Recordset")
	rsEquivalentCustomers.CursorLocation = 3 
	
    DIM pageCount,currentPage,rowPerPage,i_count,recordCount
    i_count=0
   
    
    
	If FirstLetter = "all" Then				
		SQLCustomersTable = "SELECT a.CustNum,a.Addr1, a.Addr2,a.CityStateZip, a.Phone,b.partnerCustID,a.Name FROM AR_Customer a "
        SQLCustomersTable =SQLCustomersTable +" LEFT JOIN AR_CustomerMapping b ON b.ourCustID=a.CustNum AND b.partnerRecID="& PartnerInternalRecordIdentifier 
        'IF LEN(filterdata)>0 THEN
        '    SQLCustomersTable =SQLCustomersTable + " AND b.partnerCustID like '%%"+filterdata+"%%'"
        'END IF

        SQLCustomersTable =SQLCustomersTable + " WHERE a.AcctStatus = 'A'"
        IF LEN(filterdata)>0 THEN
            SQLCustomersTable =SQLCustomersTable + " AND (a.CustNum like '%%"+filterdata+"%%' OR  b.partnerCustID like '%%"+filterdata+"%%')"
        END IF
        SQLCustomersTable =SQLCustomersTable + " ORDER BY CONVERT(int, a.CustNum) ASC"
	Else
        SQLCustomersTable = "SELECT a.CustNum,a.Addr1, a.Addr2,a.CityStateZip, a.Phone,b.partnerCustID,a.Name FROM AR_Customer a "
        SQLCustomersTable =SQLCustomersTable +" LEFT JOIN AR_CustomerMapping b ON b.ourCustID=a.CustNum AND b.partnerRecID="& PartnerInternalRecordIdentifier 
        IF LEN(filterdata)>0 THEN
            SQLCustomersTable =SQLCustomersTable + " AND b.partnerCustID like '%%"+filterdata+"%%'"
        END IF
		 SQLCustomersTable =SQLCustomersTable +" WHERE LEFT(a.Name,1) = '" & FirstLetter & "' AND a.AcctStatus = 'A'"
        IF LEN(filterdata)>0 THEN
             SQLCustomersTable =SQLCustomersTable + " AND (a.CustNum like '%%"+filterdata+"%%' OR  b.partnerCustID like '%%"+filterdata+"%%')"
        END IF
        SQLCustomersTable =SQLCustomersTable + " ORDER BY CONVERT(int, a.CustNum) ASC"
	End If
	
	'Response.write(SQLCustomersTable)
    rsCustomerTable.PageSize=rowPerPage
	rsCustomerTable.Open SQLCustomersTable,cnnCustomerTable,3,3
	'Set rsCustomerTable = cnnCustomerTable.Execute(SQLCustomersTable)
    If NOT rsCustomerTable.EOF Then
                          
        recordCount=rsCustomerTable.PageCount
        rsCustomerTable.AbsolutePage=currentPage
    END IF
                %>

<div class="container">
    <div class="row">
        <div class="col-lg-5 col-md-5 col-sm-5 col-xs-5">
            <ul class="nav nav-pills" role="tablist" style="margin: 12px 0;">
                <li role="presentation" ><a href="#">Total Accounts: <span class="badge total-customer"><%= TotalCustomerCount %></span></a></li>
                <li role="presentation"><a href="#">Partner Accounts Defined:<span class="badge"><%= TotalEquivalentPartnerCustCount %></span></a></li>
            </ul>
        </div>
        <div class="col-lg-7 col-md-7 col-sm-7 col-xs-7 text-right">
            
             <%IF recordCount>1 THEN %>
       

                <ul class="pagination pagination-sm">
                    <li>
                        <a href="#" aria-label="Previous" onclick="javascript:gotoPage(1);">
                            <span aria-hidden="true">First</span>
                        </a>
                    </li>
                    <li>
                        <a href="#" aria-label="Previous">
                            <span aria-hidden="true">&laquo;</span>
                        </a>
                    </li>
                    <%IF currentPage<9 THEN %>
                        <%IF recordCount>9 THEN %>
                            <%For jPage=1 TO 10 %>
                                <%IF jPage=currentPage THEN %>
                                    <li class="active"><a href="#"><%=jPage %></a></li>
                                    <%ELSE %>
                                        <li><a href="#" onclick="javascript:gotoPage(<%=jPage%>);"><%=jPage %></a></li>

                                <%END IF %>
                            <%Next %>
                            <%ELSE %>
                                    <%For jPage=1 TO recordCount %>
                                        <%IF jPage=currentPage THEN %>
                                        <li class="active"><a href="#"><%=jPage %></a></li>
                                        <%ELSE %>
                                            <li><a href="#" onclick="javascript:gotoPage(<%=jPage%>);"><%=jPage %></a></li>

                                        <%END IF %>
                                    <%NEXT %>
                        <%END IF %>
                        <%ELSE %>
                            <%
                                pagingStart=currentPage-4
                                
                                %>
                            <%For jPage=pagingStart TO pagingStart+9 %>
                                       <%IF jPage=currentPage THEN %>
                                    <li class="active"><a href="#"><%=jPage %></a></li>
                                    <%ELSE %>
                                        <li><a href="#" onclick="javascript:gotoPage(<%=jPage%>);"><%=jPage %></a></li>

                                <%END IF %>
                                <%IF jPage>= recordCount THEN  %>
                                    <% jPage=pagingStart+9%>
                                <%END IF %>
                            <%Next %>
                    <%END IF %>
                    
                    <li>
                      <a href="#" aria-label="Next">
                        <span aria-hidden="true">&raquo;</span>
                      </a>
                    </li>
                     <li>
                        <a href="#" onclick="javascript:gotoPage(<%=recordCount %>);" aria-label="Previous">
                            <span aria-hidden="true">Last</span>
                        </a>
                    </li>
                </ul>
          
       

            <%END IF %>
        </div>
        <div class="clearfix"></div>
         <%IF recordCount>1 THEN %>
        <div class="alert alert-warning alert-dismissible show" data-dismiss="alert" role="alert">
                To edit any entry, double click inside the Equivalent Cust ID box next to the appropriate customer.
                <button type="button" class="close" data-dismiss="alert" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                </button>
            </div>
        <%END IF %>
        <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12">
            
              <table class="table table-striped table-bordered table-responsive">
                  <thead>
                    <tr>
                        <th>Our Cust ID</th>
                        <th>Account Name</th>
                        <th>Account Address</th>
                        <th>Equivalent Cust ID</th>
                    </tr> 
                  </thead>
                  <tbody >
						<%
                        
						If NOT rsCustomerTable.EOF Then
                          
                            recordCount=rsCustomerTable.PageCount
                            rsCustomerTable.AbsolutePage=currentPage
                            Do While Not rsCustomerTable.EOF AND i_count<rowPerPage
								i_count=i_count+1
								customerID = rsCustomerTable("CustNum")
								customerName = rsCustomerTable("Name") 
								customerAddr1 = rsCustomerTable("Addr1") 
								customerAddr2 = rsCustomerTable("Addr2") 
								customerCityStateZip = rsCustomerTable("CityStateZip") 
								customerPhone = rsCustomerTable("Phone")
									
								
								
								partnerEquivalentCustID = rsCustomerTable("partnerCustID")
								
						        %>
	                          
								<tr data-toggle="tooltip" data-placement="top" title="Click to select row, double click to edit." onclick="javascript:selectRow(this);" ondblclick="javascript:editRow(this);" class="for-select" data-id="txtPartnerEquivalentCustomer*<%= customerID %>*<%= PartnerInternalRecordIdentifier %>">
		                            <td><%= customerID %></td>
		                            <td><strong><%= customerName %></strong></td>
		                            <td class="description">
		                            	<%= customerAddr1 %>
		                            	<% If customerAddr2 <> "" Then Response.Write("<br>" & customerAddr2 & "<br>") %>
		                            	<%= customerCityStateZip %><br>
		                            	<%= customerPhone %>
		                            </td>
                                    <td style="width:30%;">
                                        <%= partnerEquivalentCustID %>
		                            </td>
	                            </tr>		                            
	                            		                          
			        			<%
								rsCustomerTable.movenext
							Loop
							
						End If
						
						set rsCustomerTable = Nothing
						cnnCustomerTable.close
						set cnnCustomerTable = Nothing
						
						
						
			            %>
 
                  
                        </tbody>
                </table>

        </div>
        <div class="clearfix"></div>
        <%IF recordCount>1 THEN %>
        <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 text-right">

                <ul class="pagination pagination-sm">
                    <li>
                        <a href="#" aria-label="Previous" onclick="javascript:gotoPage(1);">
                            <span aria-hidden="true">First</span>
                        </a>
                    </li>
                    <li>
                        <a href="#" aria-label="Previous">
                            <span aria-hidden="true">&laquo;</span>
                        </a>
                    </li>
                     <%IF currentPage<9 THEN %>
                        <%IF recordCount>9 THEN %>
                            <%For jPage=1 TO 10 %>
                                <%IF jPage=currentPage THEN %>
                                    <li class="active"><a href="#"><%=jPage %></a></li>
                                    <%ELSE %>
                                        <li><a href="#" onclick="javascript:gotoPage(<%=jPage%>);"><%=jPage %></a></li>

                                <%END IF %>
                            <%Next %>
                            <%ELSE %>
                                    <%For jPage=1 TO recordCount %>
                                        <%IF jPage=currentPage THEN %>
                                        <li class="active"><a href="#"><%=jPage %></a></li>
                                        <%ELSE %>
                                            <li><a href="#" onclick="javascript:gotoPage(<%=jPage%>);"><%=jPage %></a></li>

                                        <%END IF %>
                                    <%NEXT %>
                        <%END IF %>
                        <%ELSE %>
                            <%
                                pagingStart=currentPage-4
                                
                                %>
                            <%For jPage=pagingStart TO pagingStart+9 %>
                                       <%IF jPage=currentPage THEN %>
                                    <li class="active"><a href="#"><%=jPage %></a></li>
                                    <%ELSE %>
                                        <li><a href="#" onclick="javascript:gotoPage(<%=jPage%>);"><%=jPage %></a></li>

                                <%END IF %>
                                <%IF jPage>= recordCount THEN  %>
                                    <% jPage=pagingStart+9%>
                                <%END IF %>
                            <%Next %>
                    <%END IF %>
                    
                    <li>
                      <a href="#" aria-label="Next">
                        <span aria-hidden="true">&raquo;</span>
                      </a>
                    </li>
                     <li>
                        <a href="#" onclick="javascript:gotoPage(<%=recordCount %>);" aria-label="Previous">
                            <span aria-hidden="true">Last</span>
                        </a>
                    </li>
                </ul>
          
        </div>

        <%END IF %>
    </div>
</div>