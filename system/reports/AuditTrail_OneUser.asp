<%
Server.ScriptTimeout = 900000 'Default value
%>
<!--#include file="../../inc/header.asp"-->

<!--#include file="../../inc/jquery_table_search.asp"-->

<% 

UserNameToFind = Request.QueryString("unam")
ActiveTab = Request.QueryString("tab")

CreateAuditLogEntry "Admin Report","Admin Report","Minor",0 ,MUV_Read("DisplayName") & " ran the report: Audit Trail For One User For " & UserNameToFind
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
Response.Write("<br><br>Processing Audit Trail For One User<br><br>This may take up to a full minute, please wait...<br><br>")
Response.Write("<img src=""../img/loading.gif"" />")
Response.Write("</div>")
Response.Flush()

%>

 

<h3 class="page-header"><i class="fa fa-file-text-o"></i> Audit Trail For One User For&nbsp;<%= UserNameToFind %></h3>
 


<!-- row !-->
<div class="container-fluid">
<div class="row">


<%
SQL = "SELECT * from SC_AuditLog "
SQL = SQL & "WHERE AuditUserDisplayName = '" & UserNameToFind & "' "
SQL = SQL & "Order By AuditEntryDateTime DESC"   

'Response.write(SQL)

Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Open SQL, Session("ClientCnnString")
%>



<!-- responsive tables !-->
<div class="table-responsive">


<br>
 
            <table id="tableSuperSum" class="food_planner table table-striped table-condensed table-hover sortable">
              <thead>
                <tr>
                  <th class="sorttable_numeric">Date & Time</th>
                  <th class="sorttable">Event</th> 
                  <th class="sorttable">Description</th>
                  <th class="sorttable">IP Address</th>
                </tr>
              </thead>
              

<%		
		Response.Write("<tbody class='searchable'>")
		
		Do While Not rs.EOF
		

			Response.Write("<tr>")
   		    Response.write("<td sorttable_customkey=" & FormatAsSortableDateTime(rs("AuditEntryDateTime")) & ">" & FormatDateTime(rs("AuditEntryDateTime")) & "</td>")
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