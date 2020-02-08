<%
Server.ScriptTimeout = 900000 'Default value
%>
<!--#include file="../../inc/header.asp"-->

<!--#include file="../../inc/jquery_table_search.asp"-->

<%
CreateAuditLogEntry "Admin Report","Admin Report","Minor",0, MUV_Read("DisplayName") & " ran the report: Todays Audit Trail (Full)"
%>
 
<style>

 
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
Response.Write("<br><br>Processing Today's Audit Trail Data<br><br>This may take up to a full minute, please wait...<br><br>")
Response.Write("<img src=""../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()

%>

 

<h3 class="page-header"><i class="fa fa-file-text-o"></i> Today's Audit Trail (Full) for <%=FormatDateTime(Now(),1) %> </h3>
<h6 class="page-header">
<table id="table-search" class='table table-striped table-condensed table-hover display '>
</table>
</h6>


<!-- row !-->
<div class="row">


<%

'**********************************************************
'NEED A CASE WHERE CLAUSE SO WE CAN BUILD FILTERS UPON IT
'**********************************************************
WHERE_CLAUSE = " WHERE DATEPART(dayofyear,AuditEntryDateTime) = DATEPART(dayofyear,getdate()) AND DATEPART(year,AuditEntryDateTime) = DATEPART(year,getdate()) " 


IgnoreLogin = Request.Form("chkIgnoreLogin") 
IgnoreSystem = Request.Form("chkIgnoreSystem")
		
If IgnoreLogin = "on" then
	WHERE_CLAUSE = WHERE_CLAUSE & " AND AuditElementOrEventName <> 'Login' AND AuditElementOrEventName <> 'Logout' "
End If

If IgnoreSystem = "on" then
	WHERE_CLAUSE = WHERE_CLAUSE & " AND AuditUserDisplayName <> 'SYSTEM' AND AuditUserDisplayName<> 'System' "
End If


SQL = "SELECT * FROM SC_AuditLog  "
SQL = SQL & WHERE_CLAUSE 
SQL = SQL & "ORDER BY AuditEntryDateTime DESC"   

'Response.write(SQL)

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Open SQL, Session("ClientCnnString")
%>

<!-- responsive tables !-->
<div class="table-responsive">
	
<div class="input-group"> 
	
	<span class="input-group-addon">Narrow Results</span>
	
	<div class="row">
		
		<div class="col-lg-3">
     <input id="filter" type="text" class="form-control filter-search-width" placeholder="Type here...">
		</div>

<div class="col-lg-3">
	<form method="post" action="AuditTrail_Today.asp" name="frmAuditTrailToday"> 
		<% If IgnoreLogin <> "on" Then %>
			<input type='checkbox' class='check' id='chkIgnoreLogin' name='chkIgnoreLogin' onclick="document.frmAuditTrailToday.submit ();"> Ignore login/logout
		<%Else%>
			<input type='checkbox' class='check' id='chkIgnoreLogin' name='chkIgnoreLogin' onclick="document.frmAuditTrailToday.submit ();" checked> Ignore login/logout
		<%End If%>
		
		&nbsp;&nbsp;
		
		<% If IgnoreSystem <> "on" Then %>
			<input type='checkbox' class='check' id='chkIgnoreSystem' name='chkIgnoreSystem' onclick="document.frmAuditTrailToday.submit ();"> Ignore SYSTEM
		<%Else%>
			<input type='checkbox' class='check' id='chkIgnoreSystem' name='chkIgnoreSystem' onclick="document.frmAuditTrailToday.submit ();" checked> Ignore SYSTEM
		<%End If%>
		
	</form>
</div>
	
</div>

</div>

            <table id="tableSuperSum" class="food_planner sortable table table-striped table-condensed table-hover">
              <thead>
                <tr>
                  <th class="sorttable_numeric">Time</th>
                  <th class="sorttable">User</th>
                  <th class="sorttable">Event</th> 
                  <th class="sorttable">Description</th>
                  <th class="sorttable">IP Address</th>
                </tr>
              </thead>
              

<%		
		Response.Write("<tbody class='searchable'>")
		
		Do While Not rs.EOF

			Response.Write("<tr>")
		    Response.write("<td>" & FormatDateTime(rs("AuditEntryDateTime"),3) & "</td>")
		    Response.write("<td>" & rs("AuditUserDisplayName") & "</td>")
		    Response.Write("<td>" & rs("AuditElementOrEventName") & "</td>")
		    Response.Write("<td>" & rs("AuditDescription") & "</td>")
		    Response.Write("<td>" & rs("AuditIPAddress") & "</td>")
		    Response.Write("</tr>")
		    
			rs.movenext
				
		Loop
		
		Response.Write("</tbody>")
		Response.Write("</table>")		
		Response.Write("</div>")

		
%>


            </table>
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


<!--#include file="../../inc/footer-main.asp"-->