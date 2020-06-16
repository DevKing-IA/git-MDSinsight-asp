<div class="row" style="width:98%; margin-left:10px; margin-top:20px;">


    <div class="container-fluid">
        <div class="row">
               <table id="tableSuperSumClients" class="display compact" style="width:100%;">
                  <thead>
                      <tr>	
                            <th class="td-align1 gen-info-header" colspan="5" style="border-right: 2px solid #555 !important;">General</th>
                            <th class="td-align1 vpc-3pavg-header" colspan="7" style="border-right: 2px solid #555 !important;">Sales<small>&nbsp;(Excluding Rent, XSFs, LVF & Category 21)</small></th>
                            <th class="td-align1 vpc-lcp-header" colspan="5" style="border-right: 2px solid #555 !important;">MCS</th>
                            <th class="td-align1 vpc-misc-header" colspan="3" style="border-right: 2px solid #555 !important;">MISC</th>
                            <th class="td-align1 vpc-current-header" colspan="1" style="border-right: 2px solid #555 !important;">Equipment</th>
                            <th class="td-align1 activities-header" colspan="2" style="border-right: 2px solid #555 !important;">Activities</th>
                    </tr>
                    
                    <tr>
                        <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>Acct</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Client</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Primary<br> Slsmn</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Secondary<br> Slsmn</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Install<br> Date</th>
    
                        <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-3,ReportDate)),1) %><br><%= Year(DateAdd("m",-3,ReportDate))%></th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-2,ReportDate)),1) %><br><%= Year(DateAdd("m",-2,ReportDate))%></th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %><br><%= Year(DateAdd("m",-1,ReportDate))%></th>
                        <!--<th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Current $</th>-->
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">3 Prior<br>mos avg $</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Shortage<br>Last 3 mos</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(ReportDate),1) %><br>Sales $</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">GP$<br><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %>&nbsp;<%= Year(DateAdd("m",-1,ReportDate))%></th>
    
                        
                        <th class="td-align sorttable_numeric smaller-header" style="border-left: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn"><br>MCS</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(DateAdd("m",-1,ReportDate)),1) %>&nbsp;MCS<br>Variance</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%=MonthName(Month(ReportDate),1) %>&nbsp;MTD<br>Variance</th>					
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Enrollment<br>Date</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important; border-right: 2px solid #555 !important;" id="salesColumn" >Pending<br>LVF</th>	
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><%= MonthName(Month(DateAdd("m",-1,ReportDate))) %><br>Rental $</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">LVF<br><%=MonthName(Month(DateAdd("m",-2,ReportDate)),1) %>&nbsp;<%= Year(DateAdd("m",-3,ReportDate))%></th>
                                        
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn">Max LVF<br>&nbsp;</th>
                        
                        <th class="td-align sorttable_numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important; border-left: 2px solid #555 !important;" id="salesColumn">Eqp Value</th>
                        
                        <th class="td-align sorttable_numeric smaller-header" style="border-top: 2px solid #555 !important;" id="salesColumn"><br>Action</th>
                        <th class="td-align sorttable_numeric smaller-header" style="border-right: 2px solid #555 !important; border-top: 2px solid #555 !important;" id="salesColumn">Additional<br>Info</th>
                        
                    </tr>
                  </thead>
                  
                
    
    <%		
            Response.Write("<tbody>")
    
    
        SQL = "SELECT * FROM AR_Customer INNER JOIN BI_MCSData ON BI_MCSData.CustID = AR_Customer.CustNum WHERE MonthlyContractedSalesDollars <> 0 and ChainID = 0" 
        
        Set cnn8 = Server.CreateObject("ADODB.Connection")
        cnn8.open (Session("ClientCnnString"))
        
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.CursorLocation = 3
        Set rs = cnn8.Execute(SQL)
    
        If Not rs.Eof Then
                
            Do While Not rs.EOF
    
                ShowThisRecord = True
    
                    
                If ShowThisRecord <> False Then			
                
                    PrimarySalesMan =  ""
                    SecondarySalesMan =  ""
                    SelectedCustomerID = rs("CustNum")
                    CustName = rs("Name")
                    CustMonthlyContractedSalesDollars = 0
                    InstallDate = ""
                    EnrollmentDate = ""
                    
                    PrimarySalesMan = rs("Salesman")
                    SecondarySalesMan = rs("SecondarySalesman")
                    CustMonthlyContractedSalesDollars = rs("MonthlyContractedSalesDollars")
                    InstallDate = rs("InstallDate")
                    MaxMCSCharge = rs("MaxMCSCharge")
                    EnrollmentDate =  rs("MCSEnrollmentDate")
                    
                    'Decide if this record meets the filter criteria
                    If FilterSlsmn1 <> "" And FilterSlsmn1 <> "All" Then
                        If CInt(FilterSlsmn1) <> Cint(rs("Salesman")) Then ShowThisRecord = False
                    End If
                    If FilterSlsmn2 <> "" And FilterSlsmn2 <> "All" Then
                        If CInt(FilterSlsmn2) <> Cint(rs("SecondarySalesman")) Then ShowThisRecord = False
                    End If
            
                End If
                
    
                Month3Sales_NoRent = rs("Month3Sales_NoRent") - rs("Month3Cat21Sales") 
    
                If ShowAllCusts <> 1 Then
                    If Month3Sales_NoRent >= CustMonthlyContractedSalesDollars Then ShowThisRecord = False
                End If
    
                If ShowZeroSalesCusts = 1 Then
                    If Month3Sales_NoRent > 0 Then ShowThisRecord = False
                End If
    
                VarianceHolder = Month3Sales_NoRent - CustMonthlyContractedSalesDollars 
                CurrentHolder = rs("CurrentHolder")
    
    
                CurrentMonthVarianceHolder = CurrentHolder - CustMonthlyContractedSalesDollars 
                
                ' Calc under by the current month recovered the deficit
                If VarianceHolder < 0 Then 'Meaning they have a variance
                    If CurrentHolder >= CustMonthlyContractedSalesDollars + ABS(VarianceHolder)  Then
                        If IncludeDeficitCovered <> 1 Then ShowThisRecord = False
                    End If
                End If
                
                
                 If ABS(VarianceHolder) < 100 Then
                    If Month3Sales_NoRent <> 0 Then
                        VariancePercentHolder = 100 - ((Month3Sales_NoRent/CustMonthlyContractedSalesDollars) * 100) 
                    End If
                    VariancePercentHolder  = VariancePercentHolder  * -1
                    If ApplyRule = 1 Then
                        If ABS(VariancePercentHolder) < 10 Then
                            ARCount = ARCount + 1
                            ShowThisRecord = False
                        End If
                    End If
                End If
    
                If ShowThisRecord <> False Then
    
                    Month1Sales_NoRent = rs("Month1Sales_NoRent") - rs("Month1Cat21Sales") 
                    Month2Sales_NoRent = rs("Month2Sales_NoRent") - rs("Month2Cat21Sales") 
                    
                    ThreePPSales = Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent
                    
                    Month3Cost_NoRent = rs("Month3Cost_NoRent") 
                    
                    Month3GP = Month3Sales_NoRent - Month3Cost_NoRent
                    If Not IsNumeric(Month3GP) Then Month3GP  = 0
                
                    ThreePPAvgSales = (Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent) / 3
    
                    ' New Rule 5 per david /12/6/19			
                    If ThreePPAvgSales >= CustMonthlyContractedSalesDollars Then ShowThisRecord = False  '69 to 56
                    
                    'Now see if the shortage is less than $100
                    If CustMonthlyContractedSalesDollars - ThreePPAvgSales < 100 Then
                        ' See if it is less than 10%
                        x = CustMonthlyContractedSalesDollars - ThreePPAvgSales
                        If ApplyRule <> 1 Then
                        Else
                            If (x / CustMonthlyContractedSalesDollars) * 100 < 10 Then ShowThisRecord = False ' 56 to 50
                        End If
                    End IF
                    
                End If
     
    
    
                If ShowThisRecord <> False Then
                
                    Month1Sales_NoRent = rs("Month1Sales_NoRent") - rs("Month1Cat21Sales") 
                    Month2Sales_NoRent = rs("Month2Sales_NoRent") - rs("Month2Cat21Sales") 
                    
                    ThreePPSales = Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent
                    
                    Month3Cost_NoRent = rs("Month3Cost_NoRent") 
                    
                    Month3GP = Month3Sales_NoRent - Month3Cost_NoRent
                    If Not IsNumeric(Month3GP) Then Month3GP  = 0
                
                    ThreePPAvgSales = (Month1Sales_NoRent + Month2Sales_NoRent + Month3Sales_NoRent) / 3
                    
                    ShortageHolder = ThreePPSales - (CustMonthlyContractedSalesDollars * 3)
                    
                    
                    LVFHolder = rs("LVFHolder") 
                    
                    LVFHolderCurrent = rs("LVFHolderCurrent") 
                    
                    
                    TotalEquipmentValue = rs("TotalEquipmentValue")
                    
                    TotalCustsReported = TotalCustsReported + 1
    
                    Response.Write("<tr id=""CUST" & SelectedCustomerID & """")
    
                    Response.Write(">")
                    Response.Write("<td class='smaller-detail-line'><a href='../CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>"& SelectedCustomerID  & "</a></td>")
                    Response.Write("<td class='smaller-detail-line'><a href='../CatAnalByPeriod/CatAnalByPeriod_SingleCustomer.asp?CID=" & SelectedCustomerID & "&ZDC=0&VB=3Periods&oon=new' target='_blank'>"& CustName & "</a></td>")	
                    PrimarySalesPerson = GetSalesmanNameBySlsmnSequence(PrimarySalesMan)
                    SecondarySalesPerson = GetSalesmanNameBySlsmnSequence(SecondarySalesman)
                    If Instr(PrimarySalesPerson ," ") <> 0 Then
                        Response.Write("<td class='smaller-detail-line'>" & Left(PrimarySalesPerson,Instr(PrimarySalesPerson ," ")+1) & "</td>")
                    Else
                        Response.Write("<td class='smaller-detail-line'>" & PrimarySalesPerson & "</td>")
                    End If
                    If Instr(SecondarySalesPerson," ") <> 0 Then
                        Response.Write("<td class='smaller-detail-line'>" & Left(SecondarySalesPerson,Instr(SecondarySalesPerson," ")+1) & "</td>")
                    Else
                        Response.Write("<td class='smaller-detail-line'>" & SecondarySalesPerson & "</td>")
                    End If
                    
                    If Not IsDate(InstallDate) Then InstallDate = "01/01/2000"
                    InstallDate = cDate(InstallDate) 
                    iYear = Year(InstallDate)
                    If Month(InstallDate) < 10 Then iMonth = "0" & Month(InstallDate) else iMonth = Month(InstallDate)
                    If Day(InstallDate) < 10 Then iDay = "0" & Day(InstallDate) else iDay = Day(InstallDate)
                    Response.Write("<td align='right' class='smaller-detail-line'><span class='hidden'>" & iYear & iMonth & iDay & "</span>" & Left(InstallDate,Len(InstallDate)-4) & Right(InstallDate,2) & "</td>")
                    
                    MissedMonth1 = False : MissedMonth2 = False : MissedMonth3 = False
                                     
                    If Month3Sales_NoRent - CustMonthlyContractedSalesDollars < 1 Then MissedMonth3 = True
                    If Month2Sales_NoRent - CustMonthlyContractedSalesDollars < 1 Then MissedMonth2 = True
                    If Month1Sales_NoRent - CustMonthlyContractedSalesDollars < 1 Then MissedMonth1 = True
    
                    
                    If MissedMonth3  = True AND MissedMonth2  = True AND MissedMonth1  = True Then
                        Response.Write("<td align='right' class='smaller-detail-line'><mark>" & FormatCurrency(Month1Sales_NoRent,0) & "</mark></td>")
                    Else
                        Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(Month1Sales_NoRent,0) & "</td>")				
                    End If
                    
                    If MissedMonth3  = True AND MissedMonth2  = True Then
                        Response.Write("<td align='right' class='smaller-detail-line'><mark>" & FormatCurrency(Month2Sales_NoRent,0) & "</mark></td>")
                        Response.Write("<td align='right' class='smaller-detail-line'><mark>" & FormatCurrency(Month3Sales_NoRent,0) & "</mark></td>")
                    Else
                        Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(Month2Sales_NoRent,0) & "</td>")
                        Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(Month3Sales_NoRent,0) & "</td>")
                    End If
    
                    
    
                    Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ThreePPAvgSales,0) & "</td>")
                    
                    If ShortageHolder < 0 Then
                        Response.Write("<td align='right' class='negative-thin smaller-detail-line'>" & FormatCurrency(ShortageHolder ,0,0,0) & "</td>")
                    Else
                        Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(ShortageHolder ,0,0,0) & "</td>")				    
                    End If
                    
                    
                    
                    ' Calc under by the current month recovered the deficit
                    If VarianceHolder < 0 Then 'Meaning they have a variance
                        If CurrentHolder >= CustMonthlyContractedSalesDollars + ABS(VarianceHolder)  Then
                            Response.Write("<td align='right' class='smaller-detail-line'><font color='blue'><b>" & FormatCurrency(CurrentHolder,0)  & "</b></foont></td>")
                        Else
                            Response.Write("<td align='right' class='smaller-detail-line'><font color='black'>" & FormatCurrency(CurrentHolder,0)  & "</foont></td>")
                        End If
                    Else
                        Response.Write("<td align='right' class='smaller-detail-line'><font color='black'>" & FormatCurrency(CurrentHolder,0) & "</foont></td>")
                    End If
    
                    
                    
                    
                    Response.Write("<td align='right' class='smaller-detail-line'>" &  FormatCurrency(Month3GP,0)  & "</td>")
                    
                    Response.Write("<td align='right' class='not-as-small-detail-line' style='border-left: 2px solid #555 !important;'>" & FormatCurrency(CustMonthlyContractedSalesDollars,0) & "</td>")
    
                    If VarianceHolder < 1 Then 
                        If ABS(VarianceHolder) < 1 Then
                            Response.Write("<td align='right' class='negative-thin not-as-small-detail-line'>" & "-$0"  & "</td>") ' handle variance less than 1 whole dollar
                        Else
                            Response.Write("<td align='right' class='negative-thin not-as-small-detail-line'>" & FormatCurrency(VarianceHolder ,0,0,0) & "</td>")
                        End If
                    Else
                        Response.Write("<td align='right' class='not-as-small-detail-line'>" & FormatCurrency(VarianceHolder ,0,0,0) & "</td>")
                    End If
    
                    If CurrentMonthVarianceHolder < 1 Then 
                        If ABS(CurrentMonthVarianceHolder) < 1 Then
                            Response.Write("<td align='right' class='negative-thin not-as-small-detail-line'>" & "-$0"  & "</td>") ' handle variance less than 1 whole dollar
                        Else
                            Response.Write("<td align='right' class='negative-thin not-as-small-detail-line'>" & FormatCurrency(CurrentMonthVarianceHolder ,0,0,0) & "</td>")
                        End If
                    Else
                        Response.Write("<td align='right' class='not-as-small-detail-line'>" & FormatCurrency(CurrentMonthVarianceHolder ,0,0,0) & "</td>")
                    End If
                    
                    
    
                    'EnrollmentDate Date
                    EnrollmentDate = cDate(EnrollmentDate) 
                    eYear = Year(EnrollmentDate)
                    If Month(EnrollmentDate) < 10 Then eMonth = "0" & Month(EnrollmentDate) else eMonth = Month(EnrollmentDate)
                    If Day(EnrollmentDate) < 10 Then eDay = "0" & Day(EnrollmentDate) else eDay = Day(EnrollmentDate)
                    EnrollmentDispayableDate = eMonth & "/" & eDay  & "/" & eYear
                    EnrollmentDispayableDate  = cDate(EnrollmentDispayableDate) 
                    Response.Write("<td align='right' class='smaller-detail-line'><span class='hidden'>" & eYear & eMonth & eDay & "</span>" & Left(EnrollmentDispayableDate,Len(EnrollmentDispayableDate)-4) & Right(EnrollmentDispayableDate,2) & "</td>")
    
                    PendingLVFHolder = rs("PendingLVF")
                    Response.Write("<td align='right' class='smaller-detail-line' style='border-right: 2px solid #555 !important;'>" &  FormatCurrency(PendingLVFHolder,2)  & "</td>")	
    
                    RentalHolder = rs("RentalHolder")
                    IF rs("Month3XSF") > 0 Then RentalHolder = RentalHolder  + rs("Month3XSF")
                    
                    If RentalHolder < 0 Then
                        Response.Write("<td align='right' class='negative-thin smaller-detail-line'>" & FormatCurrency(RentalHolder ,0) & "</td>")
                    Else
                        Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(RentalHolder ,0) & "</td>")				
                    End If
                    
                    Response.Write("<td align='right' class='smaller-detail-line'>" &  FormatCurrency(LVFHolder,2)  & "</td>")					
    
                    MaxLVFPerMachineHolder = MaxMCSCharge
                    If Not IsNumeric(MaxMCSCharge) Then 
                        Response.Write("<td align='right' class='smaller-detail-line'>" & MaxLVFPerMachineHolder & "</td>")
                    Else
                        Response.Write("<td align='right' class='smaller-detail-line'>" & FormatCurrency(MaxLVFPerMachineHolder,2) & "</td>")				
                    End If				
                    
                        
                    
                    If TotalEquipmentValue > 0 Then	
                    
                        LCPGP = 0 
                        
                        If TotalEquipmentValue <> 0 Then %>
                            <td align="right" class="smaller-detail-line">
                            <a data-toggle="modal" data-show="true" href="#" data-cust-id="<%= SelectedCustomerID %>" data-lcp-gp="<%= LCPGP %>" data-target="#modalEquipmentVPC" data-tooltip="true" data-title="View Customer Equipment"><%= FormatCurrency(TotalEquipmentValue,0) %></a>    
                            </td>
                        <% Else %>
                            <%= FormatCurrency(TotalEquipmentValue,0) %>
                        <% End If %>
                        
                    <%
                    Else
                        Response.Write("<td align='right' class='smaller-detail-line'>No Equipment</td>")
                    End If
    
                    'Action
                    Response.Write("<td align='right' class='smaller-detail-line'>")
                    btncolor = "btn-success"
                    
                    if GetMCSNotesStatus(SelectedCustomerID, MonthName(Month(DateAdd("m",-1,ReportDate)))) Then 
                        if GetMCSNotesNoActionStatus(SelectedCustomerID, MonthName(Month(DateAdd("m",-1,ReportDate)))) = 2 Then
                            btncolor = "btn-default noaction"
                        Else
                            btncolor = "btn-default"
                        End If					
                    End if 
                    Response.Write "<button type=""button"" class=""" & btncolor & """ id=""btn" & SelectedCustomerID & """ data-toggle=""modal"" data-target=""#modalGeneralNotesGroupM"" data-cust-id=""" & SelectedCustomerID & """ data-cust-name=""" &CustName & """ data-mcs-variance=""" & VarianceHolder & """ data-mcs-salespersonid1=""" & PrimarySalesMan & """ data-mcs-salespersonid2=""" & SecondarySalesMan & """  data-mcs-salesperson1=""" & PrimarySalesPerson & """ data-mcs-salesperson2=""" & SecondarySalesPerson & """ data-mcs-month=""" & MonthName(Month(DateAdd("m",-1,ReportDate))) & """ data-mcs-userno=""" & Session("userNo") & """ data-maxmcscharge=""" & MaxMCSCharge & """ data-mcsdollars=""" & CustMonthlyContractedSalesDollars & """ >Action</button>"
                    Response.Write("</td>")
                    
                    'Additional Info / Notes
                    
                    'Allow for a note here as a way to put in a note for the customer in general
                    'Use -2 as the category number for MCS notes
                    
                    If UserHasAnyUnviewedNotes(SelectedCustomerID) Then
                        'Pulsing icon
                        Response.Write("<td align='center' class='smaller-detail-line'>")
                        Response.Write("<a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='" & SelectedCustomerID & "' class='ole' rel='tooltip' style='cursor:pointer;'><i class='fa fa-file-text-o faa-pulse animated fa-2x' aria-hidden='true'></i></a>")																	
                        Response.Write("</td>")
                    Else
                        'Regular icon
                        Response.Write("<td align='center' class='smaller-detail-line'>")
                        Response.Write("<a data-toggle='modal' data-target='#modalEditCustomerNotes' data-category-id='-2' data-cust-id='" & SelectedCustomerID  & "' class='ole' rel='tooltip' style='cursor:pointer;'><i class='fa fa-file-text-o' aria-hidden='true'></i></a>")											
                        Response.Write("</td>")
                    End If
    
    
                    Response.Write("</tr>")
                    
                End If
    
                
                
                rs.movenext
                    
            Loop
            
            Response.Write("</tbody>")
    Else
    
        Response.Write("Nothing To Report")
    End If
    
    %>
                </table>
            </div>
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
    