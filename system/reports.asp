<!--#include file="../inc/header.asp"-->
<h1 class="page-header"><i class="fa fa-desktop"></i> Insight System Reports</h1>

<!-- row !-->
<div class="row">

	<!-- Today's Audit Trail (Full) !-->
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">Today's Audit Trail (Full)</a></p>
		<p>This report is available to Admin users only. It shows the full system audit trail for the current day for all users.</p>
		<p align="right"><a href="<%= BaseURL %>system/reports/AuditTrail_Today.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof Today's Audit Trail (Full) !-->
	
	<!-- Multi-Day Audit Trail !-->
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">Multi-Day Audit Trail</a></p>
		<p>This report is available to Admin users only. It shows the full system audit trail for the selected time period for all users.</p>
		<p align="right"><a href="<%= BaseURL %>system/reports/AuditTrail_MultiDay.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof Multi-Day Audit Trail !-->
	
	<!-- <%=GetTerm("Customer")%> Note Activty 
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title"><%=GetTerm("Customer")%> Note Activty</a></p>
		<p>This report shows new <%=GetTerm("customer")%> notes for the time period specified. Defaults to 'This Week'.</p>
		<p align="right"><a href="<%= BaseURL %>system/reports/Account_Note_Activity.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof <%=GetTerm("Customer")%> Note Activty !-->  
	    
	<!-- Audit Trail - One Line Per User !-->    
	<div class="col-md-6 reports-box">
		<img src="<%= BaseURL %>img/general/graph.png" class="imgleft">
		<p><a href="#" class="title">Audit Trail - One Line Per User</a></p>
		<p>This report is available to Admin users only. It shows only one line for each user, displaying each user's most recent activity.</p>
		<p align="right"><a href="<%= BaseURL %>system/reports/AuditTrail_One_MostRecent.asp"><button type="button" class="btn btn-primary">Run</button></a></p>
	</div> 
	<!-- eof Audit Trail - One Line Per User !-->


</div>
<!--#include file="../inc/footer-main.asp"-->