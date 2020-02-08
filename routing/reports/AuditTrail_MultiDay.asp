<%
Server.ScriptTimeout = 900000 'Default value
%>
<!--#include file="../inc/header.asp"-->

<!--#include file="../inc/jquery_table_search.asp"-->

<%
CreateAuditLogEntry "Admin Report","Admin Report","Minor",0 ,MUV_Read("DisplayName") & " ran the report: Multi-Day Audit Trail"
%>


<%
dim filter, txtStartDate, txtEndDate
txtStartDate = dateCustomFormat(DateAdd("d",-3,Date())) 
txtEndDate = dateCustomFormat(Date()) 
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then 
	txtStartDate = Request.form("txtStartDate")
	txtEndDate = Request.form("txtEndDate")
End IF
WHERE_CLAUSE ="Where CONVERT(date,AuditEntryDateTime) >= '"&txtStartDate&"' AND CONVERT(date,AuditEntryDateTime) <= '"&txtEndDate&"'"
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
</style>

<script type="text/javascript">
$(document).ready(function() {
    $("#PleaseWaitPanel").hide();
});
</script>


<%
Response.Write("<div id=""PleaseWaitPanel"">")
Response.Write("<br><br>Processing Multi-Day Audit Trail Data<br><br>This may take up to a full minute, please wait...<br><br>")
Response.Write("<img src=""../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()

%>

 

<h3 class="page-header"><i class="fa fa-file-text-o"></i> Multi-Day Audit Trail</h3>
 

<!-- row !-->
<div class="row row-data">	
	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
		<div class="input-group"> <span class="input-group-addon">Narrow Results</span>
		    <input id="filter" type="text" class="form-control " placeholder="Type here...">
		</div>
	</div>
	<form method="post" action="AuditTrail_MultiDay.asp" name="frmMultiDayAuditTrail">
		<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12">
			<strong class="element-right">Date Range</strong>
		</div>
		<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12">
			<div class="form-group">
				<input type="hidden" id="txtStartDate" name="txtStartDate" value="<%= txtStartDate%>">
				<input type="hidden" id="txtEndDate" name="txtEndDate" value="<%= txtEndDate%>">
				<div class="btn btn-default" id="reportrange">
					<i class="fa fa-calendar"></i> &nbsp;
					<span></span>
					<b class="fa fa-angle-down"></b>
				</div>
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
			</div>
		<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12">
			<a href="#" onClick="document.frmMultiDayAuditTrail.submit()"><button type="button" class="btn btn-primary">Run Report</button></a>     
		</div>
	</form>
 </div>
<!-- eof row !-->

<!-- row !-->
<div class="container-fluid">
<div class="row">


<%
SQL = "SELECT * from SC_AuditLog "
SQL = SQL & WHERE_CLAUSE 
SQL = SQL & "Order By AuditEntryDateTime DESC"   

'Response.write(SQL)

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Open SQL, Session("ClientCnnString")
%>



<!-- responsive tables !-->
<div class="table-responsive">


<br>
 
            <table id="tableSuperSum" class="food_planner sortable table table-striped table-condensed table-hover">
              <thead>
                <tr>
                  <th class="sorttable_numeric">Date & Time</th>
                  <th class="sorttable">User</th>
                  <th class="sorttable">Event</th> 
                  <th class="sorttable">Description</th>
                  <th class="sorttable">IP Address</th>
                </tr>
              </thead>
              

<%		
		Response.Write("<tbody class='searchable'>")
		
		Do 
		
			Response.Write("<tr>")
		    Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rs("AuditEntryDateTime")) & ">" & FormatDateTime(rs("AuditEntryDateTime")) & "</td>")
		    Response.write("<td>" & rs("AuditUserDisplayName") & "</td>")
		    Response.Write("<td>" & rs("AuditElementOrEventName") & "</td>")
		    Response.Write("<td>" & rs("AuditDescription") & "</td>")
		    Response.Write("<td>" & rs("AuditIPAddress") & "</td>")
		    Response.Write("</tr>")
	
			rs.movenext
				
		Loop until rs.eof
		
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

	rs.Close	
		
%>


</div>
<!-- eof row !-->


<!--#include file="../inc/footer-main.asp"-->
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
