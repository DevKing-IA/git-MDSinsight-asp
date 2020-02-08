<%
Server.ScriptTimeout = 900000 'Default value
%>
<!--#include file="../../inc/header.asp"-->
<!--#include file="../../inc/jquery_table_search.asp"-->
<!--#include file="../../inc/InsightFuncs_Orders.asp"-->

<%
dim filter, txtStartDate, txtEndDate
txtStartDate = dateCustomFormat(Date()) 
txtEndDate = dateCustomFormat(Date()) 

If Request.ServerVariables("REQUEST_METHOD") = "POST" Then 
	txtStartDate = Request.form("txtStartDate")
	txtEndDate = Request.form("txtEndDate")
End IF
WHERE_CLAUSE ="WHERE Cast(RecordCReationDateTime as date) >= '"&txtStartDate&"' AND cast(RecordCReationDateTime as date) <= '"&txtEndDate&"'"
Function dateCustomFormat(date)
	x = FormatDateTime(date, 2) 
	d = split(x, "/")
	dateCustomFormat = d(2) & "-" & d(0) & "-" & d(1)
End Function
%> 

 
<style>
	
	.element-right{
		float:right;
		margin-top: 5px;
	}
	
	.row-data{
		margin-bottom: 15px;
	}

.filter-search-width{
	max-width: 36%;
}

table.sortable th:not(.sorttable_sorted):not(.sorttable_sorted_reverse):not(.sorttable_nosort):after { 
    content: " \25B4\25BE" 
    
}

table.sortable thead {
    color:#222;
    font-weight: bold;
    cursor: pointer;
}

#PleaseWaitPanel{
position: fixed;
left: 470px;
top: 275px;
width: 975px;
height: 300px;
z-index: 9999;
background-color: #fff;
opacity:1.0;
text-align:center;
} 

#reportrange{
	float: left;
	margin-right: 5px;
}   
</style>



<%
Response.Write("<div id=""PleaseWaitPanel"">")
Response.Write("<br><br>Processing Orders Received via API Data<br><br>Please wait...<br><br>")
Response.Write("<img src=""../../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()

%>

 

<h3 class="page-header"><i class="fa fa-file-text-o"></i> Orders Received via API</h3>
 

<!-- row !-->
<div class="row row-data">	
	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
		<div class="input-group"> <span class="input-group-addon">Search</span>
		    <input id="filter" type="text" class="form-control " placeholder="Type here...">
		</div>
	</div>
	<form method="post" action="main.asp" name="frmMain">
		<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12">
			<strong class="element-right">Show orders received on the following date(s)</strong>
		</div>
		<div class="col-lg-5 col-md-5 col-sm-12 col-xs-12">
 				<input type="hidden" id="txtStartDate" name="txtStartDate" value="<%= txtStartDate%>">
				<input type="hidden" id="txtEndDate" name="txtEndDate" value="<%= txtEndDate%>">
				<div class="btn btn-default" id="reportrange">
					<i class="fa fa-calendar"></i> &nbsp;
					<span></span>
					<b class="fa fa-angle-down"></b>
				</div>
 	      	<select class="form-control hidden" name="selDtRange">
		         	<option <%If Request.form("selDtRange") = "All Dates" Then %>selected<%End If%>>All Dates</option>
		         	<option <%If Request.form("selDtRange") = "Today" Then %>selected<%End If%>>Today</option>
					<option <%If Request.form("selDtRange") = "This Week" Then %>selected<%End If%>>This Week</option>
					<option <%If Request.form("selDtRange") = "This Month" Then %>selected<%End If%>>This Month</option>
					<option <%If Request.form("selDtRange") = "This Quarter" Then %>selected<%End If%>>This Quarter</option>
					<option <%If Request.form("selDtRange") = "Last 3 Days" or Request.form("selDtRange") = "" Then %>selected<%End If%>>Last 3 Days</option>
					<option <%If Request.form("selDtRange") = "Last 10 Days" Then %>selected<%End If%>>Last 10 Days</option>
					<option <%If Request.form("selDtRange") = "Last 30 Days" Then %>selected<%End If%>>Last 30 Days</option>
					<option <%If Request.form("selDtRange") = "Last 60 Days" Then %>selected<%End If%>>Last 60 Days</option>
					<option <%If Request.form("selDtRange") = "Last 90 Days" Then %>selected<%End If%>>Last 90 Days</option>
			</select>

			<a href="#" onClick="document.frmMain.submit()"><button type="button" class="btn btn-primary">Apply Date(s)</button></a>  

			</div>
		 
	</form>
 </div>
<!-- eof row !-->

<!-- row !-->
<div class="container-fluid">
<div class="row">


<%
Set cnnOrderAPI = Server.CreateObject("ADODB.Connection")
cnnOrderAPI.open (Session("ClientCnnString"))

Set rsOrderHeaderTMP = Server.CreateObject("ADODB.Recordset")
Set rsOrderHeaderTMP2 = Server.CreateObject("ADODB.Recordset")
Set rsOrderHeader = Server.CreateObject("ADODB.Recordset")
rsOrderHeaderTMP.CursorLocation = 3 

On Error Resume Next ' In caase the table isn't there
SQLOrderHeaderTMP  = "DROP TABLE zAPI_OR_OrderHeaderTmp_" & trim(Session("Userno"))
Set rsOrderHeaderTMP = cnnOrderAPI.Execute(SQLOrderHeaderTMP)
On Error Goto 0
	
'Get All Orders That Came In On That Date
SQLOrderHeaderTMP= "SELECT * INTO  zAPI_OR_OrderHeaderTmp_" & trim(Session("Userno"))
SQLOrderHeaderTMP= SQLOrderHeaderTMP& " FROM API_OR_OrderHeader "
SQLOrderHeaderTMP= SQLOrderHeaderTMP& WHERE_CLAUSE  

Set rsOrderHeaderTMP = cnnOrderAPI.Execute(SQLOrderHeaderTMP)


'Delete All But Most Recent
SQLOrderHeaderTMP = "SELECT MAX(RecordCreationDateTime) as Expr1, OrderID, Count(OrderID) as NumOrders FROM zAPI_OR_OrderHeaderTmp_" & trim(Session("Userno")) & " GROUP BY OrderID, Day(RecordCreationDateTime)"
Set rsOrderHeaderTMP = cnnOrderAPI.Execute(SQLOrderHeaderTMP)

If Not rsOrderHeaderTMP.EOF Then
	Do While Not rsOrderHeaderTMP.EOF
	 	If rsOrderHeaderTMP("NumOrders") > 1 Then ' Dont delete if there's only 1 record
			SQLOrderHeaderTMP2 = "DELETE FROM zAPI_OR_OrderHeaderTmp_" & trim(Session("Userno")) & " WHERE OrderID = '" & rsOrderHeaderTMP("OrderID") & "' "
			SQLOrderHeaderTMP2 = SQLOrderHeaderTMP2 & " AND DateDiff(ss,RecordCreationDateTime,'" & rsOrderHeaderTMP("Expr1") & "') > 0 "
			Set rsOrderHeaderTMP2 = cnnOrderAPI.Execute(SQLOrderHeaderTMP2)
		End If
		rsOrderHeaderTMP.movenext
	Loop
End If

Set rsOrderHeaderTMP2 = Nothing
Set rsOrderHeaderTMP = Nothing
'Now only good records are left in the temp table

%>



<!-- responsive tables !-->
<div class="table-responsive">
	<br>
	<table id="tableSuperSum" class="food_planner sortable table table-striped table-condensed table-hover">
		<thead>
		    <tr>
				<th class="sorttable_numeric">Received</th>
				<th class="sorttable_numeric">Last Updated</th>
				<th class="sorttable">Order ID</th>
				<th class="sorttable">Customer ID</th> 
				<th class="sorttable">Ship To Company</th>
				<th class="sorttable">Requsted<br>Delivery Date</th>
				<th class="sorttable">Current<br>Status</th>
				<th class="sorttable_numeric"># Lines</th>
				<th class="sorttable_numeric">Total</th>
				<th class="sorttable_numeric"># Threads</th>                  
				<th class="sorttable_nosort">&nbsp;</th>
			</tr>
		</thead>
              

		<%


		Response.Write("<tbody class='searchable'>")
		
		SQLOrderHeader = "SELECT * FROM  zAPI_OR_OrderHeaderTmp_" & trim(Session("Userno")) & " ORDER BY RecordCreationDateTime DESC"
				
		'Response.Write(SQLOrderHeader & "<BR>")
		Set rsOrderHeader = cnnOrderAPI.Execute(SQLOrderHeader)
				
		If Not rsOrderHeader.EOF Then 
			Do While Not rsOrderHeader.EOF
			
				Response.Write("<tr>")
				Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rsOrderHeader("RecordCreationDateTime")) & ">" & FormatDateTime(rsOrderHeader("RecordCreationDateTime")) & "</td>")
				If GetAPIOrderLastUpdatedDateTime(rsOrderHeader("OrderID")) > rsOrderHeader("RecordCreationDateTime") Then
			    	Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(GetAPIOrderLastUpdatedDateTime(rsOrderHeader("OrderID"))) & "><strong>" & FormatDateTime(GetAPIOrderLastUpdatedDateTime(rsOrderHeader("OrderID"))) & "</strong></td>")
			    Else
				    Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(GetAPIOrderLastUpdatedDateTime(rsOrderHeader("OrderID"))) & ">" & FormatDateTime(GetAPIOrderLastUpdatedDateTime(rsOrderHeader("OrderID"))) & "</td>")
				End If					    
			    Response.write("<td>" & rsOrderHeader("OrderID") & "</td>")
			    Response.Write("<td>" & rsOrderHeader("CustID") & "</td>")
			    Response.Write("<td>" & rsOrderHeader("ShipToCompany") & "</td>")
			    Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rsOrderHeader("RequestedDeliveryDate")) & ">" & FormatDateTime(rsOrderHeader("RequestedDeliveryDate")) & "</td>")
			    If APIOrderIsVoided(rsOrderHeader("OrderID")) = True Then
				    Response.Write("<td><font color='red'>" & "Deleted" & "</font></td>")
				Else
				    Response.Write("<td>" & "Order" & "</td>")
				End If
			    Response.Write("<td>" & NumberOfAPIOrderLines(rsOrderHeader("OrderID"),GetAPIOrderHighestThread(rsOrderHeader("OrderID")))  & "</td>")
			    Response.Write("<td>" & FormatCurrency(rsOrderHeader("GrandTotal"),2) & "</td>")
			    Response.Write("<td>" & GetAPIOrderHighestThread(rsOrderHeader("OrderID")) & "</td>")
			    ' To begin with, we only allow the latest version to be reposted
			    If GetAPIOrderLastUpdatedDateTime(rsOrderHeader("OrderID")) = rsOrderHeader("RecordCreationDateTime") Then
			    'See if we allow reposting
			   		If GetAPIRepostURL() <> "" AND APIRepostOrders() = True Then 
				    	If GetUserOrderAPIPermissionLevel(Session("UserNo")) = "READ_RESEND" Or UserIsAdmin(Session("UserNo")) Then
					    	Response.Write("<td>")
							Response.Write("<button type='button' class='btn btn-success btn-sm btn-block' name='resendtoBackend" & rsOrderHeader("InternalRecordIdentifier") & "' id='" & rsOrderHeader("InternalRecordIdentifier") & "'>")
							Response.Write("<input type='hidden' id='txtInternalRecordIdentifier " & rsOrderHeader("InternalRecordIdentifier") & "' name='txtInternalRecordIdentifier" & rsOrderHeader("InternalRecordIdentifier") & "' value='" & rsOrderHeader("InternalRecordIdentifier") & "'>")
							Response.Write("<i class='fa fa-share'></i>&nbsp;Re-send to " & GetTerm("Backend") )
							Response.Write("</button>")
					    	Response.Write("</td>")
					    Else
						    Response.Write("<td>&nbsp;</td>")
					    End If
					Else
						Response.Write("<td>&nbsp;</td>")
				    End If
				Else
					Response.Write("<td>&nbsp;</td>")				
				End If	

			    Response.Write("</tr>")

				rsOrderHeader.movenext
					
			Loop
		End If
		
		Response.Write("</tbody>")
		Response.Write("</table>")		
		Response.Write("</div>")

		
%>


            </table>
          </div>
          </div>
<!-- eof responsive tables !-->



<!-- eof row !-->

<!-- row !-->
<div class="row">

<div class="col-lg-12"><hr></div>
</div>
<!-- eof row !-->

<!-- row !-->
<div class="row">

<%		

	rsOrderHeader.Close	
		
%>


</div>
<!-- eof row !-->


<!--#include file="../../inc/footer-main.asp"-->
<script src="<%= baseURL %>js/moment.min.js" type="text/javascript"></script>
<link href="<%= baseURL %>js/bootstrap-daterangepicker/daterangepicker.min.css" rel="stylesheet" type="text/css" />
<script src="<%= baseURL %>js/bootstrap-daterangepicker/daterangepicker.min.js" type="text/javascript"></script>
<script type="text/javascript">
		function setPredefined(){
			$("#optPredefined").prop("checked","checked");
		}
		startDate = moment($('#txtStartDate').val());
		endDate = moment($('#txtEndDate').val());
        $('#reportrange').daterangepicker({
                opens: 'right',
                startDate: startDate,
                endDate: endDate,
                //minDate: '01/01/2012',
                //maxDate: '12/31/2014',
                //dateLimit: {
                //    days: 60
                //},
                showDropdowns: true,
                showWeekNumbers: true,
                timePicker: false,
                timePickerIncrement: 1,
                timePicker12Hour: true,
                ranges: {
                    'Today': [moment(), moment()],
                    'Yesterday': [moment().subtract('days', 1), moment().subtract('days', 1)],
                    'Last 7 Days': [moment().subtract('days', 6), moment()],
                    'Last 10 Days': [moment().subtract('days', 9), moment()],
                    'Last 30 Days': [moment().subtract('days', 29), moment()],
                    'This Month': [moment().startOf('month'), moment().endOf('month')],
                    'Last Month': [moment().subtract('month', 1).startOf('month'), moment().subtract('month', 1).endOf('month')]
                },
                buttonClasses: ['btn'],
                applyClass: 'green',
                cancelClass: 'default',
                format: 'MM/DD/YYYY',
                separator: ' to ',
                locale: {
                    applyLabel: 'Apply',
                    fromLabel: 'From',
                    toLabel: 'To',
                    customRangeLabel: 'Custom Range',
                    daysOfWeek: ['Su', 'Mo', 'Tu', 'We', 'Th', 'Fr', 'Sa'],
                    monthNames: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
                    firstDay: 1
                }
            },
            function (start, end) {
                $('#reportrange span').html(start.format('MMMM D, YYYY') + ' - ' + end.format('MMMM D, YYYY'));
                $('#txtStartDate').val(start.format('MM/DD/YYYY'));
                $('#txtEndDate').val(end.format('MM/DD/YYYY'));
				$("#optCustom").prop("checked","checked");
            }
        );
        //Set the initial state of the picker label
        $('#reportrange span').html(startDate.format('MMMM D, YYYY') + ' - ' + endDate.format('MMMM D, YYYY'));
		$('#txtStartDate').val(startDate.format('MM/DD/YYYY'));
		$('#txtEndDate').val(endDate.format('MM/DD/YYYY'));
</script>
