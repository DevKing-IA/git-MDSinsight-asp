<!--#include file="../../inc/header-prospecting.asp"-->
<!--#include file="../../inc/InSightFuncs_Prospecting.asp"-->
<%'<!--#include file="dashboardBuildSQLTable.asp"-->%>
<!-- Styles -->
<style>

#chartdiv {
  width: 100%;
  height: 500px;
}

#chartdivProspectsCreatedBySalesRep {
  width: 100%;
  height: 500px;
  background-color: #fff;
}

#chartdivProspectsCreatedByLeadSource {
  width: 100%;
  height: 500px;
  background-color: #fff;
}

#chartdivAppointmentsAttendedBySalesRep {
  width: 100%;
  height: 500px;
  background-color: #fff;
}

#chartdivAppointmentsAttendedByLeadSource {
  width: 100%;
  height: 500px;
  background-color: #fff;
}


#chartdivNewClientsBySalesRep {
  width: 100%;
  height: 500px;
  background-color: #fff;
}

#chartdivNewClientsByLeadSource {
  width: 100%;
  height: 500px;
  background-color: #fff;
}


#chartdivUnqualifiedByReasonSalesRep {
  width: 100%;
  height: 500px;
  background-color: #fff;
}

#chartdivLostByReasonSalesRep {
  width: 100%;
  height: 500px;
  background-color: #fff;
}

#chartdivUnqualifiedByReasonLeadSource {
  width: 100%;
  height: 600px;
  background-color: #fff; 
  color: #fff; 
}

#chartdivLostByReasonLeadSource {
  width: 100%;
  height: 600px;
  background-color: #fff;	

}

#chartdivQualifiedClientsBySalesRep {
  width: 100%;
  height: 500px;
  background-color: #fff;
}

#chartdivQualifiedClientsByLeadSource {
  width: 100%;
  height: 500px;
  background-color: #fff;
}


.amcharts-export-menu-top-right {
  top: 10px;
  right: 0;
}
</style>

<!-- Resources -->
<script src="https://www.amcharts.com/lib/3/amcharts.js"></script>
<script src="https://www.amcharts.com/lib/3/serial.js"></script>
<script src="https://www.amcharts.com/lib/3/plugins/export/export.min.js"></script>
<link rel="stylesheet" href="https://www.amcharts.com/lib/3/plugins/export/export.css" type="text/css" media="all" />
<script src="https://www.amcharts.com/lib/3/themes/light.js"></script>
<script src="https://www.amcharts.com/lib/3/themes/none.js"></script>
<script src="https://www.amcharts.com/lib/3/themes/black.js"></script>
<script src="https://www.amcharts.com/lib/3/themes/dark.js"></script>


<%


'************************
'Read Settings_Screens
'************************
SelectedUserNumbersToDisplay = ""

Set cnnSettingsScreen = Server.CreateObject("ADODB.Connection")
cnnSettingsScreen.open Session("ClientCnnString")

SQLSettingsScreen = "SELECT * FROM Settings_Screens WHERE ScreenNumber = 1100 AND UserNo = " & Session("userNo")

Set rsInsightSettingsScreen = Server.CreateObject("ADODB.Recordset")
rsInsightSettingsScreen.CursorLocation = 3 
Set rsInsightSettingsScreen= cnnSettingsScreen.Execute(SQLSettingsScreen)

If NOT rsInsightSettingsScreen.EOF Then
	SelectedUserNumbersToDisplay = rsInsightSettingsScreen("ScreenSpecificData1")
	SelectedUserNumbersToDisplay = Left(SelectedUserNumbersToDisplay,Len(SelectedUserNumbersToDisplay)-1)
	SelectedUserNumbersToDisplay = Right(SelectedUserNumbersToDisplay,Len(SelectedUserNumbersToDisplay)-1)
Else
	Set cnnSettingsUsers = Server.CreateObject("ADODB.Connection")
	cnnSettingsUsers.open Session("ClientCnnString")
	
	SQLSettingsUsers = "SELECT * FROM tblUsers WHERE UserNo <> '' ORDER BY UserNo"
	Set rsInsightSettingsUsers = Server.CreateObject("ADODB.Recordset")
	rsInsightSettingsUsers.CursorLocation = 3 
	Set rsInsightSettingsUsers = cnnSettingsUsers.Execute(SQLSettingsUsers)
	
	If NOT rsInsightSettingsUsers.EOF Then
	
		Do While NOT rsInsightSettingsUsers.EOF
			SelectedUserNumbersToDisplay = SelectedUserNumbersToDisplay & rsInsightSettingsUsers("UserNo") & ","
			rsInsightSettingsUsers.MoveNext
		Loop
		
		SelectedUserNumbersToDisplay = Left(SelectedUserNumbersToDisplay,Len(SelectedUserNumbersToDisplay)-1)
		SelectedUserNumbersToDisplay = Right(SelectedUserNumbersToDisplay,Len(SelectedUserNumbersToDisplay)-2)

	End If

End If

Set rsInsightSettingsScreen= Nothing


'****************************
'End Read Settings_Screens
'****************************


barGraphColorArray12 = Array("#FF0F00", "#FF6600", "#FF9E01", "#FCD202", "#F8FF01", "#B0DE09", "#04D215", "#0D8ECF", "#0D52D1", "#2A0CD0", "#8A0CCF", "#CD0D74")
barGraphColorArray7 = Array("#4572a7","#aa4643","#89a54e","#80699b","#3d96ae","#db843d","#d8cc08")
barGraphColorArray22 = Array("#cc3333","#a24057","#606692","#3a85a8","#42977e","#4aaa54","#629363","#7e6e85","#9c509b","#c4625d","#eb751f","#ff9709","#ffc81d","#fff830","#e1c62f","#bf862b","#ad5a36","#cc6a6f","#eb7aa9","#bc8fa7","#999999","#787878")
stackedBarGraphColorArray6 = Array("#ff8533","#fddb35","#c0e53a","#3da5d9","#553dd9","#d73d90")
stackedBarGraphColorArray10 = Array("#5e8dcb","#d26260","#7faa39","#8d73af","#58b9d5","#58b9d5","#7b91b6","#da9392","#f2ee7a","#bcd499")

mondayOfThisWeek = DateAdd("d", -((Weekday(date()) + 7 - 2) Mod 7), date())
mondayOfLastWeek = DateAdd("ww",-1,mondayOfThisWeek)
yesterday = DateAdd("d",-1, date())


%>
<!--#include file="dashboard_inc_prospects_created.asp"-->
<!--#include file="dashboard_inc_appointments_completed.asp"-->
<!--#include file="dashboard_inc_prospects_qualified.asp"-->
<!--#include file="dashboard_inc_prospects_unqualified.asp"-->
<!--#include file="dashboard_inc_prospects_won.asp"-->
<!--#include file="dashboard_inc_prospects_lost.asp"-->


<h1 class="page-header"><i class="fa fa-fw fa-asterisk"></i> <%= GetTerm("Prospecting") %> Dashboard</h1>
<% 
'Response.write("mondayOfThisWeek : " & mondayOfThisWeek & "<br>")
'Response.write("mondayOfLastWeek : " & mondayOfLastWeek & "<br>")
'Response.write("SelectedUserNumbersToDisplay : " & SelectedUserNumbersToDisplay & "<br>")
%>
	<hr>
	<h2><i class="fa fa-plus-circle" aria-hidden="true"></i> Prospects Created</h2>
	<hr>
		
	<div class="row">
	  <div class="col-sm-12">
        <div class="well">
        	<h4 class="text-success">Prospects Created By Sales Rep </h4>
			<!-- HTML -->
			<p>View the sales reps that have created prospects for the week beginning on <%= mondayOfLastWeek %>.</p>
			
			<% If showProspectsCreatedBySalesRepChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelProspectsCreatedBySalesRep">Show/Hide Graph</button>
				<div id="panelProspectsCreatedBySalesRep" class="collapse">
					<div id="chartdivProspectsCreatedBySalesRep"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
		
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->

	<div class="row">
	  <div class="col-sm-12">
        <div class="well">
        	<h4 class="text-success">Prospects Created By Lead Source </h4>
			<!-- HTML -->
			<p>View the lead sources that have led to created prospects for the week beginning on <%= mondayOfLastWeek %>.</p>
			
			<% If showProspectsCreatedByLeadSourceChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelProspectsCreatedByLeadSource">Show/Hide Graph</button>
				<div id="panelProspectsCreatedByLeadSource" class="collapse">
					<div id="chartdivProspectsCreatedByLeadSource"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
		
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->
	





	<hr>
	<h2><i class="fa fa-calendar-check-o" aria-hidden="true"></i> Appointments Completed</h2>
	<hr>
	
	<div class="row">
	  <div class="col-sm-12">
        <div class="well" style="background-color:#ffffe0;">
        	<h4 class="text-success">Appointments Completed By Sales Rep </h4>
			<!-- HTML -->
			<p>View the sales reps that have attended appointments/meetings for the week beginning on <%= mondayOfLastWeek %>.</p>

			<% If showAppmtsCompletedBySalesRepChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelAppointmentsAttendedBySalesRep">Show/Hide Graph</button>
				<div id="panelAppointmentsAttendedBySalesRep" class="collapse">
					<div id="chartdivAppointmentsAttendedBySalesRep"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
					
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->

	<div class="row">
	  <div class="col-sm-12">
        <div class="well" style="background-color:#ffffe0;">
        	<h4 class="text-success">Appointments Completed By Lead Source </h4>
			<!-- HTML -->
			<p>View the lead sources that have led to attended appointments/meetings for the week beginning on <%= mondayOfLastWeek %>.</p>
			
			<% If showAppmtsCompletedByLeadSourceChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelAppointmentsAttendedByLeadSource">Show/Hide Graph</button>
				<div id="panelAppointmentsAttendedByLeadSource" class="collapse">
					<div id="chartdivAppointmentsAttendedByLeadSource"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
		
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->




 
	
	<hr>
	<h2><i class="fa fa-user-plus" aria-hidden="true"></i> New Clients (Converted to Customers)</h2>
	<hr>

	<div class="row">
	  <div class="col-sm-12">
        <div class="well" style="background-color:#e1ffff;">
        	<h4 class="text-success">New Clients Converted By Sales Rep </h4>
			<!-- HTML -->
			<p>View the sales reps that have converted prospects to customers for the week beginning on <%= mondayOfLastWeek %>.</p>
			
			<% If showNewClientsBySalesRepChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelNewClientsBySalesRep">Show/Hide Graph</button>
				<div id="panelNewClientsBySalesRep" class="collapse">
					<div id="chartdivNewClientsBySalesRep"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
		
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->
	
	<div class="row">
	  <div class="col-sm-12">
        <div class="well" style="background-color:#e1ffff;">
        	<h4 class="text-success">New Clients Converted By Lead Source </h4>
			<!-- HTML -->
			<p>View the lead sources that have led to prospects converted to customers for the week beginning on <%= mondayOfLastWeek %>.</p>
			
			<% If showNewClientsByLeadSourceChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelNewClientsByLeadSource">Show/Hide Graph</button>
				<div id="panelNewClientsByLeadSource" class="collapse">
					<div id="chartdivNewClientsByLeadSource"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
	
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->








	
	<hr>
	<h2><i class="fa fa-check-circle" aria-hidden="true"></i> Qualified Prospects</h2>
	<hr>
	
	<div class="row">
	  <div class="col-sm-12">
        <div class="well" style="background-color:#efffe0;">
        	<h4 class="text-success">Qualified Prospects By Sales Rep </h4>
			<!-- HTML -->
			<p>View the sales reps that have qualfied prospects for the week beginning on <%= mondayOfLastWeek %>.</p>
			
			<% If showQualifiedClientsBySalesRepChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelQualifiedBySalesRep">Show/Hide Graph</button>
				<div id="panelQualifiedBySalesRep" class="collapse">
					<div id="chartdivQualifiedClientsBySalesRep"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
			
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->
	
	<div class="row">
	  <div class="col-sm-12">
        <div class="well" style="background-color:#efffe0;">
        	<h4 class="text-success">Qualified Prospects By Lead Source </h4>
			<!-- HTML -->
			<p>View the lead sources that have generated qualfied prospects for the week beginning on <%= mondayOfLastWeek %>.</p>
			
			<% If showQualifiedClientsByLeadSourceChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelQualifiedByLeadSource">Show/Hide Graph</button>
				<div id="panelQualifiedByLeadSource" class="collapse">
					<div id="chartdivQualifiedClientsByLeadSource"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
		
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->
	



	
	<hr>
	<h2><i class="fa fa-times-circle" aria-hidden="true"></i> Unqualified Prospects By Reason</h2>
	<hr>
	
	<div class="row">
	  <div class="col-sm-12">
        <div class="well" style="background-color:#ffe6cc;">
        	<h4 class="text-success">Unqualified Prospects By Sales Rep </h4>
			<!-- HTML -->
			<p>View the sales reps that have unqualfied prospects for the week beginning on <%= mondayOfLastWeek %>, grouped by reason.</p>
			
			<% If showUnqualifiedBySalesRepChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelUnqualifiedByReasonSalesRep">Show/Hide Graph</button>
				<div id="panelUnqualifiedByReasonSalesRep" class="collapse">
					<div id="chartdivUnqualifiedByReasonSalesRep"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
					
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->
	
	<div class="row">
	  <div class="col-sm-12">
        <div class="well" style="background-color:#ffe6cc;">
        	<h4 class="text-success">Unqualified Prospects By Lead Source </h4>
			<!-- HTML -->
			<p>View the lead sources that have generated unqualfied prospects for the week beginning on <%= mondayOfLastWeek %>, grouped by reason.</p>
			
			<% If showUnqualifiedByLeadSourceChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelUnqualifiedByReasonLeadSource">Show/Hide Graph</button>
				<div id="panelUnqualifiedByReasonLeadSource" class="collapse">
					<div id="chartdivUnqualifiedByReasonLeadSource"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->
	






	
	<hr>
	<h2><i class="fa fa-user-times" aria-hidden="true"></i> Lost Prospects By Reason</h2>
	<hr>
		
	<div class="row">
	  <div class="col-sm-12">
        <div class="well" style="background-color:#fffafa;">
        	<h4 class="text-success">Prospects Lost By Sales Rep and Reason </h4>
			<!-- HTML -->
			<p>View the sales reps that have lost prospects for the week beginning on <%= mondayOfLastWeek %>, grouped by reason.</p>
			
			<% If showLostBySalesRepChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelLostByReasonSalesRep">Show/Hide Graph</button>
				<div id="panelLostByReasonSalesRep" class="collapse">
					<div id="chartdivLostByReasonSalesRep"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
		
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->
	
	<div class="row">
	  <div class="col-sm-12">
        <div class="well" style="background-color:#fffafa;">
        	<h4 class="text-success">Prospects Lost By Lead Source and Reason </h4>
			<!-- HTML -->
			<p>View the lead sources that have lost prospects  for the week beginning on <%= mondayOfLastWeek %>, grouped by reason.</p>
			
			<% If showLostByLeadSourceChart = "True" Then %>
				<button type="button" class="btn btn-info" data-toggle="collapse" data-target="#panelLostByReasonLeadSource">Show/Hide Graph</button>
				<div id="panelLostByReasonLeadSource" class="collapse">
					<div id="chartdivLostByReasonLeadSource"></div>
				</div>	
			<% Else %>
				<div><strong>No Data To Show</strong></div>
			<% End If %>
		
        </div>  
	  </div><!--/col-12-->
	</div><!--/row-->


	


	<!--

	<div class="row">
	  <div class="col-sm-12">
	    <div class="row">
	      <div class="col-md-4">
	        <div class="well">
	          <h4 class="text-danger"><span class="label label-danger pull-right">- 9%</span> New Users </h4>
	        </div>
	      </div>
	      <div class="col-md-4">
	        <div class="well">
	          <h4 class="text-success"><span class="label label-success pull-right">+ 3%</span> Returning </h4>
	        </div>
	      </div>
	      <div class="col-md-4">
	        <div class="well">
	          <h4 class="text-primary"><span class="label label-primary pull-right">201</span> Sales </h4>
	        </div>
	      </div>
	    </div><!--/row-->    
	  <!--</div><!--/col-12-->
	<!--</div><!--/row-->

	<hr>

<!--#include file="../../inc/footer-main.asp"-->


 