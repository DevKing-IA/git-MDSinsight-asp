<%
Server.ScriptTimeout = 900000 'Default value
%>
<!--#include file="../../../inc/header.asp"-->

<!--#include file="../../../inc/jquery_table_search.asp"-->

<%
CreateAuditLogEntry "Report","Report","Minor",0 ,MUV_Read("DisplayName") & " ran the report: " & GetTerm("Account") & " Note Activity"
%>


<%
If Request.form("selDtRange") <> "" Then
	'Construct WHERE_CLAUSE variable
	Select Case Request.form("selDtRange")
		Case "All Dates"
			WHERE_CLAUSE =""
		Case "Today"
			WHERE_CLAUSE ="Where DATEPART(dayofyear,EntryDateTime) = DATEPART(dayofyear,getdate()) AND DATEPART(year,EntryDateTime) = DATEPART(year,getdate()) "
		Case "This Week"
			WHERE_CLAUSE ="Where DATEPART(week,EntryDateTime) = DATEPART(week,getdate()) AND DATEPART(year,EntryDateTime) = DATEPART(year,getdate()) "
		Case "This Month"
			WHERE_CLAUSE ="Where DATEPART(month,EntryDateTime) = DATEPART(month,getdate()) AND DATEPART(year,EntryDateTime) = DATEPART(year,getdate()) "
		Case "This Quarter"
			WHERE_CLAUSE ="Where DATEPART(quarter,EntryDateTime) = DATEPART(quarter,getdate()) AND DATEPART(year,EntryDateTime) = DATEPART(year,getdate()) "
		Case "Last 3 Days"
			WHERE_CLAUSE ="Where EntryDateTime > DateAdd(d,-3,getdate()) "
		Case "Last 10 Days"
			WHERE_CLAUSE ="Where EntryDateTime > DateAdd(d,-10,getdate()) "
		Case "Last 30 Days"
			WHERE_CLAUSE ="Where EntryDateTime > DateAdd(d,-30,getdate()) "
		Case "Last 60 Days"
			WHERE_CLAUSE ="Where EntryDateTime > DateAdd(d,-60,getdate()) "
		Case "Last 90 Days"
			WHERE_CLAUSE ="Where EntryDateTime > DateAdd(d,-90,getdate()) "
		Case Else
			WHERE_CLAUSE=""
	End Select
End IF
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

table{
	font-size: 12px;
}

.date-time{
	width:10%;
}

.note-col{
	width: 30%;
}

.client-id-col{
	width: 5%;
}

.clients-col{
	width: 10%;
}

.username-col{
	width: 5%;
}
    
</style>

<script type="text/javascript">
$(document).ready(function() {
    $("#PleaseWaitPanel").hide();
});
</script>


<%
Response.Write("<div id=""PleaseWaitPanel"">")
Response.Write("<br><br>Processing " & GetTerm("Customer") &" Note Activity, please wait...<br><br>")
Response.Write("<img src=""../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()

%>

 

<h3 class="page-header"><i class="fa fa-file-text-o"></i><%=GetTerm("Customer")%> Note Activity</h3>
 

<!-- row !-->
<div class="row row-data">	
	<div class="col-lg-3 col-md-3 col-sm-12 col-xs-12">
		<div class="input-group"> <span class="input-group-addon">Narrow Results</span>
		    <input id="filter" type="text" class="form-control " placeholder="Type here...">
		</div>
	</div>
	<form method="post" action="Account_Note_Activity.asp" name="frmAccountNoteActivity">
		<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12">
			<strong class="element-right">Date Range</strong>
		</div>
		<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12">
	      	<select class="form-control" name="selDtRange">
		         	<option <%If Request.form("selDtRange") = "All Dates" Then %>selected<%End If%>>All Dates</option>
		         	<option <%If Request.form("selDtRange") = "Today" Then %>selected<%End If%>>Today</option>
					<option <%If Request.form("selDtRange") = "This Week" or Request.form("selDtRange") = "" Then %>selected<%End If%>>This Week</option>
					<option <%If Request.form("selDtRange") = "This Month" Then %>selected<%End If%>>This Month</option>
					<option <%If Request.form("selDtRange") = "This Quarter" Then %>selected<%End If%>>This Quarter</option>
					<option <%If Request.form("selDtRange") = "Last 3 Days" Then %>selected<%End If%>>Last 3 Days</option>
					<option <%If Request.form("selDtRange") = "Last 10 Days" Then %>selected<%End If%>>Last 10 Days</option>
					<option <%If Request.form("selDtRange") = "Last 30 Days" Then %>selected<%End If%>>Last 30 Days</option>
					<option <%If Request.form("selDtRange") = "Last 60 Days" Then %>selected<%End If%>>Last 60 Days</option>
					<option <%If Request.form("selDtRange") = "Last 90 Days" Then %>selected<%End If%>>Last 90 Days</option>
			</select>
			</div>
		<div class="col-lg-2 col-md-2 col-sm-12 col-xs-12">
			<a href="#" onClick="document.frmAccountNoteActivity.submit()"><button type="button" class="btn btn-primary">Run Report</button></a>     
		</div>
	</form>
 </div>
<!-- eof row !-->

<!-- row !-->
<div class="container-fluid">
 <div class="row">


<%
'Default to 1 week
IF WHERE_CLAUSE = "" Then WHERE_CLAUSE = " Where DATEPART(week,EntryDateTime) = DATEPART(week,getdate()) AND DATEPART(year,EntryDateTime) = DATEPART(year,getdate()) "

SQL = "SELECT EntryDateTime,CustNum,Note,UserNo,'NA' as AttachmentFilename from tblCustomerNotes "
SQL = SQL & WHERE_CLAUSE 
SQL = SQL & " UNION "
SQL = SQL & "SELECT EntryDateTime,CustNum,Note,UserNo,AttachmentFilename from tblCustomerNotesAttachments "
SQL = SQL & WHERE_CLAUSE 
SQL = SQL & "Order By EntryDateTime DESC, CustNum"   

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
                  <th class="sorttable_numeric date-time">Date & Time</th>
                  <th class="sorttable  client-id-col"><%=GetTerm("Account")%> #</th>
                  <th class="sorttable clients-col"><%=GetTerm("Account")%> Name</th>                  
                  <th class="sorttable username-col">User</th> 
                  <th class="sorttable note-col">Note</th>
                  <th class="sorttable">Attachment</th>
                </tr>
              </thead>
              

<%		
		Response.Write("<tbody class='searchable'>")
		If Not rs.Eof Then
			Do 
			
				Response.Write("<tr>")
			    Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rs("EntryDateTime")) & ">" & FormatDateTime(rs("EntryDateTime")) & "</td>")
	   		    Response.write("<td>" & rs("CustNum") & "</td>")
			    Response.write("<td>" & GetCustNameByCustNum(rs("CustNum")) & "</td>")
			    Response.Write("<td>" & GetUserDisplayNameByUserNo(rs("UserNo")) & "</td>")
			    Response.Write("<td>"  & rs("Note") &   "</td>")
			    If rs("AttachmentFilename") <> "NA" Then
				    Response.Write("<td>" & right(rs("AttachmentFilename"),len(rs("AttachmentFilename"))-Instr(rs("AttachmentFilename"),"-")) &"</td>")
				Else
				    Response.Write("<td>&nbsp;</td>")			
				End IF
			    Response.Write("</tr>")
		
				rs.movenext
					
			Loop until rs.eof
			
			Response.Write("</tbody>")
			'Response.Write("</table>")		
			'Response.Write("</div>")
		End If
		
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


<!--#include file="../../../inc/footer-main.asp"-->