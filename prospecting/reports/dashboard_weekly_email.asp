<!--#include file="../../inc/header-prospecting.asp"-->
<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Prospecting.asp"-->
<%'<!--#include file="dashboardBuildSQLTable.asp"-->%>
<!-- Styles -->
<style>
h2 {
    margin-top: 40px;
    margin-bottom: 10px;
}
    .text-success {
    color: #3c763d;
    margin-top:40px;
}
</style>

<!-- Resources -->

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
End If

Set rsInsightSettingsScreen= Nothing

'Response.write(SelectedUserNumbersToDisplay )
'****************************
'End Read Settings_Screens
'****************************


mondayOfThisWeek = DateAdd("d", -((Weekday(date()) + 7 - 2) Mod 7), date())
mondayOfLastWeek = DateAdd("ww",-1,mondayOfThisWeek)
yesterday = DateAdd("d",-1, date())


%>


<h1 class="page-header"><i class="fa fa-fw fa-asterisk"></i> <%= GetTerm("Prospecting") %> Dashboard</h1>

	<hr>
	<h2><i class="fa fa-plus-circle" aria-hidden="true"></i> Prospects Created</h2>
	<hr>
	
        	<h4 class="text-success">Prospects Created By Sales Rep </h4>
			<!-- HTML -->
			<p>View the sales reps that have created prospects for the week beginning on <%= mondayOfLastWeek %>.</p>
			<table class="table">	
		        <thead>
		            <tr>
		                <th>Sales Rep</th>
		                <th>Prospects Created</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
				
				SQLProspectSalesRep = "SELECT tblUsers.userNo, COUNT(PR_Prospects.CreatedByUserNo) AS SalesRepCount"
				SQLProspectSalesRep = SQLProspectSalesRep & " FROM  tblUsers LEFT OUTER JOIN"
				SQLProspectSalesRep = SQLProspectSalesRep & " PR_Prospects ON PR_Prospects.CreatedByUserNo = tblUsers.userNo"
				SQLProspectSalesRep = SQLProspectSalesRep & " WHERE tblUsers.userNo IN (" & SelectedUserNumbersToDisplay & ") "
				SQLProspectSalesRep = SQLProspectSalesRep & " AND  PR_Prospects.CreatedDate >= '" & mondayOfLastWeek & "' AND PR_Prospects.CreatedDate < '" & mondayOfThisWeek & "' "
				SQLProspectSalesRep = SQLProspectSalesRep & " GROUP BY tblUsers.userNo"
				SQLProspectSalesRep = SQLProspectSalesRep & " ORDER BY SalesRepCount DESC"
				
				Set cnnProspectSalesRep = Server.CreateObject("ADODB.Connection")
				cnnProspectSalesRep.open(Session("ClientCnnString"))
				Set rsProspectSalesRep = Server.CreateObject("ADODB.Recordset")
				rsProspectSalesRep.CursorLocation = 3 
				Set rsProspectSalesRep = cnnProspectSalesRep.Execute(SQLProspectSalesRep)
				
				If NOT rsProspectSalesRep.EOF Then
				
					rowCount = 1
				
					Do While Not rsProspectSalesRep.EOF
				
						SalesRepCount = rsProspectSalesRep("SalesRepCount")
						CreatedByUserNo = rsProspectSalesRep("userNo")
						CreatedByUserName = GetUserDisplayNameByUserNo(CreatedByUserNo)
						%>
						<tr>
			                <td><%= CreatedByUserName %></td>
			                <td><%= SalesRepCount %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsProspectSalesRep.MoveNext
					Loop
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>
		    
		    
        	<h4 class="text-success">Prospects Created By Lead Source </h4>
			<!-- HTML -->
			<p>View the lead sources that have led to created prospects for the week beginning on <%= mondayOfLastWeek %>.</p>
			<table class="table">	
		        <thead>
		            <tr>
		                <th>Sales Rep</th>
		                <th>Prospects Created</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
				SQLProspectLeadSource = "SELECT PR_LeadSources.InternalRecordIdentifier AS LeadSourceNum, COUNT(PR_Prospects.InternalRecordIdentifier) AS LeadSourceCount "
				SQLProspectLeadSource = SQLProspectLeadSource & " FROM  PR_LeadSources LEFT OUTER JOIN"
				SQLProspectLeadSource = SQLProspectLeadSource & " PR_Prospects ON PR_Prospects.LeadSourceNumber = PR_LeadSources.InternalRecordIdentifier"
				SQLProspectLeadSource = SQLProspectLeadSource & " WHERE PR_Prospects.CreatedDate >= '" & mondayOfLastWeek & "' AND PR_Prospects.CreatedDate < '" & mondayOfThisWeek & "' "
				SQLProspectLeadSource = SQLProspectLeadSource & " GROUP BY PR_LeadSources.InternalRecordIdentifier"
				SQLProspectLeadSource = SQLProspectLeadSource & " ORDER BY LeadSourceCount DESC"
				
				Set cnnProspectLeadSource = Server.CreateObject("ADODB.Connection")
				cnnProspectLeadSource.open(Session("ClientCnnString"))
				Set rsProspectLeadSource = Server.CreateObject("ADODB.Recordset")
				rsProspectLeadSource.CursorLocation = 3 
				Set rsProspectLeadSource = cnnProspectLeadSource.Execute(SQLProspectLeadSource)
				
				If NOT rsProspectLeadSource.EOF Then
				
					rowCount = 0
				
					Do While Not rsProspectLeadSource.EOF
				
						LeadSourceCount = rsProspectLeadSource("LeadSourceCount")
						LeadSourceNumber = rsProspectLeadSource("LeadSourceNum")
						LeadSource = GetLeadSourceByNum(LeadSourceNumber)
						%>
						<tr>
			                <td><%= LeadSource %></td>
			                <td><%= LeadSourceCount %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsProspectLeadSource.MoveNext
					Loop
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>





	<hr>
	<h2><i class="fa fa-calendar-check-o" aria-hidden="true"></i> Appointments Completed</h2>
	<hr>
	
        	<h4 class="text-success">Appointments Completed By Sales Rep </h4>
			<!-- HTML -->
			<p>View the sales reps that have attended appointments/meetings for the week beginning on <%= mondayOfLastWeek %>.</p>
			<table class="table">	
		        <thead>
		            <tr>
		                <th>Sales Rep</th>
		                <th>Appmts Completed</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%

					SQLAppmtCompleteSalesRep = "SELECT tblUsers.userNo, COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
					SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " FROM  tblUsers LEFT OUTER JOIN"
					SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " PR_ProspectActivities ON PR_ProspectActivities.ActivityCreatedByUserNo = tblUsers.userNo"
					SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " WHERE tblUsers.userNo IN (" & SelectedUserNumbersToDisplay & ") AND "
					SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " PR_ProspectActivities.ActivityIsMeeting=1 AND PR_ProspectActivities.Status='Completed' "
					SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " AND PR_ProspectActivities.StatusDateTime >= '" & mondayOfLastWeek & "' AND PR_ProspectActivities.StatusDateTime < '" & mondayOfThisWeek & "' "
					SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " GROUP BY tblUsers.userNo"
					SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " ORDER BY AppmtCount DESC"
					
					Set cnnAppmtCompleteSalesRep = Server.CreateObject("ADODB.Connection")
					cnnAppmtCompleteSalesRep.open(Session("ClientCnnString"))
					Set rsAppmtCompleteSalesRep = Server.CreateObject("ADODB.Recordset")
					rsAppmtCompleteSalesRep.CursorLocation = 3 
					Set rsAppmtCompleteSalesRep = cnnAppmtCompleteSalesRep.Execute(SQLAppmtCompleteSalesRep)
					
					If NOT rsAppmtCompleteSalesRep.EOF Then
				
					rowCount = 0
				
					Do While Not rsAppmtCompleteSalesRep.EOF
					
						AppmtCount = rsAppmtCompleteSalesRep("AppmtCount")
						CreatedByUserNo = rsAppmtCompleteSalesRep("userNo")
						CreatedByUserName = GetUserDisplayNameByUserNo(CreatedByUserNo)
						%>
						<tr>
			                <td><%= CreatedByUserName %></td>
			                <td><%= AppmtCount %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsAppmtCompleteSalesRep.MoveNext
					Loop
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>



        	<h4 class="text-success">Appointments Completed By Lead Source </h4>
			<!-- HTML -->
			<p>View the lead sources that have led to attended appointments/meetings for the week beginning on <%= mondayOfLastWeek %>.</p>
			<table class="table">	
		        <thead>
		            <tr>
		                <th>Lead Source</th>
		                <th>Appmts Completed</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%

				SQLAppmtCompleteLeadSource = "SELECT PR_LeadSources.InternalRecordIdentifier AS LeadSourceNum, COUNT(PR_ProspectActivities.ProspectRecID) AS AppmtCount "
				SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " FROM  PR_LeadSources LEFT OUTER JOIN"
				SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " PR_Prospects ON PR_Prospects.LeadSourceNumber = PR_LeadSources.InternalRecordIdentifier "
				SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " INNER JOIN PR_ProspectActivities ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier "
				SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " WHERE PR_ProspectActivities.ActivityIsMeeting=1 and PR_ProspectActivities.Status='Completed'"
				SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " AND PR_ProspectActivities.StatusDateTime >='" & mondayOfLastWeek & "' AND PR_ProspectActivities.StatusDateTime <'" & mondayOfThisWeek & "' "
				SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " GROUP BY PR_LeadSources.InternalRecordIdentifier"
				SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " ORDER BY AppmtCount DESC"
				
				Set cnnAppmtCompleteLeadSource = Server.CreateObject("ADODB.Connection")
				cnnAppmtCompleteLeadSource.open(Session("ClientCnnString"))
				Set rsAppmtCompleteLeadSource = Server.CreateObject("ADODB.Recordset")
				rsAppmtCompleteLeadSource.CursorLocation = 3 
				Set rsAppmtCompleteLeadSource = cnnAppmtCompleteLeadSource.Execute(SQLAppmtCompleteLeadSource)
				
				If NOT rsAppmtCompleteLeadSource.EOF Then
				
					rowCount = 0
				
					Do While Not rsAppmtCompleteLeadSource.EOF
									
						AppmtCount = rsAppmtCompleteLeadSource("AppmtCount")
						LeadSourceNumber = rsAppmtCompleteLeadSource("LeadSourceNum")
						LeadSource = GetLeadSourceByNum(LeadSourceNumber)
						%>
						<tr>
			                <td><%= LeadSource %></td>
			                <td><%= AppmtCount %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsAppmtCompleteLeadSource.MoveNext
					Loop
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>






 
	
	<hr>
	<h2><i class="fa fa-user-plus" aria-hidden="true"></i> New Clients (Converted to Customers)</h2>
	<hr>


        	<h4 class="text-success">New Clients Converted By Sales Rep </h4>
			<!-- HTML -->
			<p>View the sales reps that have converted prospects to customers for the week beginning on <%= mondayOfLastWeek %>.</p>
			<table class="table">	
		        <thead>
		            <tr>
		                <th>Sales Rep</th>
		                <th># New Clients</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
				SQLNewClientsBySalesRep = "SELECT tblUsers.userNo, COUNT(PR_Prospects.OwnerUserNo) AS SalesRepCount, Pool"
				SQLNewClientsBySalesRep = SQLNewClientsBySalesRep & " FROM tblUsers LEFT OUTER JOIN"
				SQLNewClientsBySalesRep = SQLNewClientsBySalesRep & " PR_Prospects ON PR_Prospects.CreatedByUserNo = tblUsers.userNo"
				SQLNewClientsBySalesRep = SQLNewClientsBySalesRep & " WHERE  PR_Prospects.Pool='Won' AND tblUsers.userNo IN (" & SelectedUserNumbersToDisplay & ") "
				SQLNewClientsBySalesRep = SQLNewClientsBySalesRep & " GROUP BY tblUsers.userNo, Pool "
				SQLNewClientsBySalesRep = SQLNewClientsBySalesRep & " ORDER BY SalesRepCount DESC"
				
				Set cnnNewClientsBySalesRep = Server.CreateObject("ADODB.Connection")
				cnnNewClientsBySalesRep.open(Session("ClientCnnString"))
				Set rsNewClientsBySalesRep = Server.CreateObject("ADODB.Recordset")
				rsNewClientsBySalesRep.CursorLocation = 3 
				Set rsNewClientsBySalesRep = cnnNewClientsBySalesRep.Execute(SQLNewClientsBySalesRep)
				
				If NOT rsNewClientsBySalesRep.EOF Then
				
					rowCount = 0
				
					Do While Not rsNewClientsBySalesRep.EOF

						SalesRepCount = rsNewClientsBySalesRep("SalesRepCount")
						OwnerUserNo = rsNewClientsBySalesRep("userNo")
						OwnerUserName = GetUserDisplayNameByUserNo(OwnerUserNo)
						%>
						<tr>
			                <td><%= OwnerUserName %></td>
			                <td><%= SalesRepCount %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsNewClientsBySalesRep.MoveNext
					Loop
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>





	
        	<h4 class="text-success">New Clients Converted By Lead Source </h4>
			<!-- HTML -->
			<p>View the lead sources that have led to prospects converted to customers for the week beginning on <%= mondayOfLastWeek %>.</p>
			<table class="table">	
		        <thead>
		            <tr>
		                <th>Lead Source</th>
		                <th># New Clients</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
				SQLNewClientsByLeadSource = "SELECT COUNT(InternalRecordIdentifier) AS LeadSourceCount, LeadSourceNumber, Pool"
				SQLNewClientsByLeadSource = SQLNewClientsByLeadSource & " FROM  PR_Prospects"
				SQLNewClientsByLeadSource = SQLNewClientsByLeadSource & " WHERE  Pool='Won' "
				SQLNewClientsByLeadSource = SQLNewClientsByLeadSource & " GROUP BY LeadSourceNumber, Pool"
				SQLNewClientsByLeadSource = SQLNewClientsByLeadSource & " ORDER BY LeadSourceCount DESC"
				
				Set cnnNewClientsByLeadSource = Server.CreateObject("ADODB.Connection")
				cnnNewClientsByLeadSource.open(Session("ClientCnnString"))
				Set rsNewClientsByLeadSource = Server.CreateObject("ADODB.Recordset")
				rsNewClientsByLeadSource.CursorLocation = 3 
				Set rsNewClientsByLeadSource = cnnNewClientsByLeadSource.Execute(SQLNewClientsByLeadSource)
				
				If NOT rsNewClientsByLeadSource.EOF Then
				
					rowCount = 0
				
					Do While Not rsNewClientsByLeadSource.EOF
				
						LeadSourceCount = rsNewClientsByLeadSource("LeadSourceCount")
						LeadSourceNumber = rsNewClientsByLeadSource("LeadSourceNumber")
						LeadSource = GetLeadSourceByNum(LeadSourceNumber)
						%>
						<tr>
			                <td><%= LeadSource %></td>
			                <td><%= LeadSourceCount %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsNewClientsByLeadSource.MoveNext
					Loop
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>








	
	<hr>
	<h2><i class="fa fa-check-circle" aria-hidden="true"></i> Qualified Prospects</h2>
	<hr>
	
        	<h4 class="text-success">Qualified Prospects By Sales Rep </h4>
			<!-- HTML -->
			<p>View the sales reps that have qualfied prospects for the week beginning on <%= mondayOfLastWeek %>.</p>
			<table class="table">	
		        <thead>
		            <tr>
		                <th>Sales</th>
		                <th># Prospects Qualified</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
				SQLQualifiedClientsBySalesRep = "SELECT tblUsers.userNo, COUNT(PR_DashboardSummaryByOwnerQ_LastWeek.OwnerUserNo) AS SalesRepCount"
				SQLQualifiedClientsBySalesRep = SQLQualifiedClientsBySalesRep & " FROM tblUsers LEFT OUTER JOIN"
				SQLQualifiedClientsBySalesRep = SQLQualifiedClientsBySalesRep & " PR_DashboardSummaryByOwnerQ_LastWeek ON PR_DashboardSummaryByOwnerQ_LastWeek.OwnerUserNo = tblUsers.userNo"
				SQLQualifiedClientsBySalesRep = SQLQualifiedClientsBySalesRep & " WHERE  tblUsers.userNo IN (" & SelectedUserNumbersToDisplay & ") "
				SQLQualifiedClientsBySalesRep = SQLQualifiedClientsBySalesRep & " GROUP BY tblUsers.userNo "
				SQLQualifiedClientsBySalesRep = SQLQualifiedClientsBySalesRep & " ORDER BY SalesRepCount DESC"
				
				
				Set cnnQualifiedClientsBySalesRep = Server.CreateObject("ADODB.Connection")
				cnnQualifiedClientsBySalesRep.open(Session("ClientCnnString"))
				Set rsQualifiedClientsBySalesRep = Server.CreateObject("ADODB.Recordset")
				rsQualifiedClientsBySalesRep.CursorLocation = 3 
				Set rsQualifiedClientsBySalesRep = cnnQualifiedClientsBySalesRep.Execute(SQLQualifiedClientsBySalesRep)
				
				If NOT rsQualifiedClientsBySalesRep.EOF Then
				
					rowCount = 0
				
					Do While Not rsQualifiedClientsBySalesRep.EOF
				
				
						SalesRepCount = rsQualifiedClientsBySalesRep("SalesRepCount")
						OwnerUserNo = rsQualifiedClientsBySalesRep("userNo")
						OwnerUserName = GetUserDisplayNameByUserNo(OwnerUserNo)
						%>
						<tr>
			                <td><%= OwnerUserName %></td>
			                <td><%= SalesRepCount %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsQualifiedClientsBySalesRep.MoveNext
					Loop
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>
		    
		    
		    
		    

        	<h4 class="text-success">Qualified Prospects By Lead Source </h4>
			<!-- HTML -->
			<p>View the lead sources that have generated qualfied prospects for the week beginning on <%= mondayOfLastWeek %>.</p>
			<table class="table">	
		        <thead>
		            <tr>
		                <th>Lead Source</th>
		                <th># Prospects Qualified</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
				SQLQualifiedClientsByLeadSource = "SELECT COUNT(InternalRecordIdentifier) AS LeadSourceCount, LeadSourceNumber "
				SQLQualifiedClientsByLeadSource = SQLQualifiedClientsByLeadSource & " FROM  PR_DashboardSummaryByLSourceQ_LastWeek "
				SQLQualifiedClientsByLeadSource = SQLQualifiedClientsByLeadSource & " GROUP BY LeadSourceNumber "
				SQLQualifiedClientsByLeadSource = SQLQualifiedClientsByLeadSource & " ORDER BY LeadSourceCount DESC"
				
				Set cnnQualifiedClientsByLeadSource = Server.CreateObject("ADODB.Connection")
				cnnQualifiedClientsByLeadSource.open(Session("ClientCnnString"))
				Set rsQualifiedClientsByLeadSource = Server.CreateObject("ADODB.Recordset")
				rsQualifiedClientsByLeadSource.CursorLocation = 3 
				Set rsQualifiedClientsByLeadSource = cnnQualifiedClientsByLeadSource.Execute(SQLQualifiedClientsByLeadSource)
				
				If NOT rsQualifiedClientsByLeadSource.EOF Then
				
					rowCount = 0
				
					Do While Not rsQualifiedClientsByLeadSource.EOF
				
						LeadSourceCount = rsQualifiedClientsByLeadSource("LeadSourceCount")
						LeadSourceNumber = rsQualifiedClientsByLeadSource("LeadSourceNumber")
						LeadSource = GetLeadSourceByNum(LeadSourceNumber)
						%>
						<tr>
			                <td><%= LeadSource %></td>
			                <td><%= LeadSourceCount %></td>
			            </tr>

						<%
						rowCount = rowCount + 1
						rsQualifiedClientsByLeadSource.MoveNext
					Loop
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>



	
	<hr>
	<h2><i class="fa fa-times-circle" aria-hidden="true"></i> Unqualified Prospects By Reason</h2>
	<hr>
	
        	<h4 class="text-success">Unqualified Prospects By Sales Rep </h4>
			<!-- HTML -->
			<p>View the sales reps that have unqualfied prospects for the week beginning on <%= mondayOfLastWeek %>, grouped by reason.</p>
			<table class="table">	
		        <thead>
		            <tr>
		                <th><%= GetTerm("Salesperson") %></th>
		                <th>Reason / # Prospects Unqualified</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
				SQLUnqualifiedBySalesRepReasons = "SELECT * "
				SQLUnqualifiedBySalesRepReasons = SQLUnqualifiedBySalesRepReasons & " FROM  PR_DashboardSummaryByOwnerUQ_LastWeek "
				SQLUnqualifiedBySalesRepReasons = SQLUnqualifiedBySalesRepReasons & " WHERE OwnerUserNo IN (" & SelectedUserNumbersToDisplay & ") "
				SQLUnqualifiedBySalesRepReasons = SQLUnqualifiedBySalesRepReasons & " AND LastStageNumber = 0 "
				SQLUnqualifiedBySalesRepReasons = SQLUnqualifiedBySalesRepReasons & " ORDER BY OwnerUserNo, ReasonNo "
				
				
				Set cnnUnqualifiedBySalesRep = Server.CreateObject("ADODB.Connection")
				cnnUnqualifiedBySalesRep.open(Session("ClientCnnString"))
				Set rsUnqualifiedBySalesRep = Server.CreateObject("ADODB.Recordset")
				rsUnqualifiedBySalesRep.CursorLocation = 3 
				Set rsUnqualifiedBySalesRep = cnnUnqualifiedBySalesRep.Execute(SQLUnqualifiedBySalesRepReasons)
				
				If NOT rsUnqualifiedBySalesRep.EOF Then
				
					CurrentOwnerUserNo = ""
					repChange = 0
					
					Do While Not rsUnqualifiedBySalesRep.EOF
					
				
						If CurrentOwnerUserNo <> "" Then
							If CurrentOwnerUserNo <> rsUnqualifiedBySalesRep("OwnerUserNo") Then
								jChartDataSalesRepUnqualified = jChartDataSalesRepUnqualified & "},"
								CurrentOwnerUserNo = rsUnqualifiedBySalesRep("OwnerUserNo")
								repChange = 1
							End If
						Else
							repChange = 1
							CurrentOwnerUserNo = rsUnqualifiedBySalesRep("OwnerUserNo")
						End If
					
						OwnerUserNo = rsUnqualifiedBySalesRep("OwnerUserNo")
						OwnerUserName = GetUserDisplayNameByUserNo(OwnerUserNo)
					
						If rsUnqualifiedBySalesRep("NumberOfProspects") > 0 Then
							
							If repChange = 1 Then
								%>
								<tr>
					                <td><%= OwnerUserName %></td>
					                <td><%= GetReasonByNum(rsUnqualifiedBySalesRep("ReasonNo")) %> : <%= rsUnqualifiedBySalesRep("NumberOfProspects") %></td>
					            </tr>
								<%
								repChange = 0
							Else
								%>
								<tr>
					                <td>&nbsp;</td>
					                <td><%= GetReasonByNum(rsUnqualifiedBySalesRep("ReasonNo")) %> : <%= rsUnqualifiedBySalesRep("NumberOfProspects") %></td>
					            </tr>
								<%						
							End If
						End If
					 			
						rsUnqualifiedBySalesRep.MoveNext
					Loop
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>
		    


        	<h4 class="text-success">Unqualified Prospects By Lead Source </h4>
			<!-- HTML -->
			<p>View the lead sources that have generated unqualfied prospects for the week beginning on <%= mondayOfLastWeek %>, grouped by reason.</p>
			<table class="table">	
		        <thead>
		            <tr>
		                <th>Lead Source</th>
		                <th>Reason / # Prospects Unqualified</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
								
				SQLUnqualifiedByLeadSourceReasons = "SELECT * "
				SQLUnqualifiedByLeadSourceReasons = SQLUnqualifiedByLeadSourceReasons & " FROM  PR_DashboardSummaryByLSourceUQ_LastWeek "
				SQLUnqualifiedByLeadSourceReasons = SQLUnqualifiedByLeadSourceReasons & " WHERE LastStageNumber = 0 "
				SQLUnqualifiedByLeadSourceReasons = SQLUnqualifiedByLeadSourceReasons & " ORDER BY LeadSourceNumber "
				
				
				Set cnnUnqualifiedByLeadSource = Server.CreateObject("ADODB.Connection")
				cnnUnqualifiedByLeadSource.open(Session("ClientCnnString"))
				Set rsUnqualifiedByLeadSource = Server.CreateObject("ADODB.Recordset")
				rsUnqualifiedByLeadSource.CursorLocation = 3 
				Set rsUnqualifiedByLeadSource = cnnUnqualifiedByLeadSource.Execute(SQLUnqualifiedByLeadSourceReasons)
								
				If NOT rsUnqualifiedByLeadSource.EOF Then
				
					showUnqualifiedByLeadSourceChart = "True"
					jChartDataLeadSourceUnqualified = ""
					CurrentLeadSourceNumber = ""
					leadsourceChange = 0
					
					Do While Not rsUnqualifiedByLeadSource.EOF
					
						If CurrentLeadSourceNumber <> "" Then
							If CurrentLeadSourceNumber <> rsUnqualifiedByLeadSource("LeadSourceNumber") Then
								CurrentLeadSourceNumber = rsUnqualifiedByLeadSource("LeadSourceNumber")
								leadsourceChange = 1
							Else
								jChartDataLeadSourceUnqualified = jChartDataLeadSourceUnqualified & ","
							End If
						Else
							leadsourceChange = 1
							CurrentLeadSourceNumber = rsUnqualifiedByLeadSource("LeadSourceNumber")
						End If
					
						LeadSourceNumber = rsUnqualifiedByLeadSource("LeadSourceNumber")
						LeadSource = GetLeadSourceByNum(LeadSourceNumber)
						
						If LeadSource = "" Then LeadSource = "Blank"
						
					
						If rsUnqualifiedByLeadSource("NumberOfProspects") > 0 Then
							
							If leadsourceChange = 1 Then
								%>
								<tr>
					                <td><%= LeadSource %></td>
					                <td><%= GetReasonByNum(rsUnqualifiedByLeadSource("ReasonNo")) %> : <%= rsUnqualifiedByLeadSource("NumberOfProspects") %></td>
					            </tr>
								<%
								leadsourceChange = 0
							Else
								%>
								<tr>
					                <td>&nbsp;</td>
					                <td><%= GetReasonByNum(rsUnqualifiedByLeadSource("ReasonNo")) %> : <%= rsUnqualifiedByLeadSource("NumberOfProspects") %></td>
					            </tr>
								<%						
							End If
						End If
						

						rsUnqualifiedByLeadSource.MoveNext
					Loop
					
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>




	
	<hr>
	<h2><i class="fa fa-user-times" aria-hidden="true"></i> Lost Prospects By Reason</h2>
	<hr>
		

        	<h4 class="text-success">Prospects Lost By Sales Rep and Reason </h4>
			<!-- HTML -->
			<p>View the sales reps that have lost prospects for the week beginning on <%= mondayOfLastWeek %>, grouped by reason.</p>

			<table class="table">	
		        <thead>
		            <tr>
		                <th><%= GetTerm("Salesperson") %></th>
		                <th>Reason / # Prospects Lost</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
				
				SQLLostBySalesRepReasons = "SELECT * "
				SQLLostBySalesRepReasons = SQLLostBySalesRepReasons & " FROM  PR_DashboardSummaryByOwnerUQ_LastWeek "
				SQLLostBySalesRepReasons = SQLLostBySalesRepReasons & " WHERE OwnerUserNo IN (" & SelectedUserNumbersToDisplay & ") AND LastStageNumber = 1"
				SQLLostBySalesRepReasons = SQLLostBySalesRepReasons & " ORDER BY OwnerUserNo, ReasonNo "
				
				
				Set cnnLostBySalesRep = Server.CreateObject("ADODB.Connection")
				cnnLostBySalesRep.open(Session("ClientCnnString"))
				Set rsLostBySalesRep = Server.CreateObject("ADODB.Recordset")
				rsLostBySalesRep.CursorLocation = 3 
				Set rsLostBySalesRep = cnnLostBySalesRep.Execute(SQLLostBySalesRepReasons)
				
				If NOT rsLostBySalesRep.EOF Then
				
					showLostBySalesRepChart = "True"
					jChartDataSalesRepLost = ""
					CurrentOwnerUserNo = ""
					repChange = 0
					
					Do While Not rsLostBySalesRep.EOF
					
						If CurrentOwnerUserNo <> "" Then
							If CurrentOwnerUserNo <> rsLostBySalesRep("OwnerUserNo") Then
								CurrentOwnerUserNo = rsLostBySalesRep("OwnerUserNo")
								repChange = 1
							End If
						Else
							repChange = 1
							CurrentOwnerUserNo = rsLostBySalesRep("OwnerUserNo")
						End If
					
						OwnerUserNo = rsLostBySalesRep("OwnerUserNo")
						OwnerUserName = GetUserDisplayNameByUserNo(OwnerUserNo)

					
						If rsLostBySalesRep("NumberOfProspects") > 0 Then
							
							If repChange = 1 Then
								%>
								<tr>
					                <td><%= OwnerUserName %></td>
					                <td><%= GetReasonByNum(rsLostBySalesRep("ReasonNo")) %> : <%= rsLostBySalesRep("NumberOfProspects") %></td>
					            </tr>
								<%
								repChange = 0
							Else
								%>
								<tr>
					                <td>&nbsp;</td>
					                <td><%= GetReasonByNum(rsLostBySalesRep("ReasonNo")) %> : <%= rsLostBySalesRep("NumberOfProspects") %></td>
					            </tr>
								<%						
							End If
						End If

						rsLostBySalesRep.MoveNext
					Loop
					
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>


        	<h4 class="text-success">Prospects Lost By Lead Source and Reason </h4>
			<!-- HTML -->
			<p>View the lead sources that have lost prospects  for the week beginning on <%= mondayOfLastWeek %>, grouped by reason.</p>

			<table class="table">	
		        <thead>
		            <tr>
		                <th>Lead Source</th>
		                <th>Reason / # Prospects Lost</th>
		            </tr>
		        </thead>
		        <tbody>
		        
				<%
								
				SQLLostByLeadSourceReasons = "SELECT * "
				SQLLostByLeadSourceReasons = SQLLostByLeadSourceReasons & " FROM  PR_DashboardSummaryByLSourceUQ_LastWeek "
				SQLLostByLeadSourceReasons = SQLLostByLeadSourceReasons & " WHERE LastStageNumber = 1"
				SQLLostByLeadSourceReasons = SQLLostByLeadSourceReasons & " ORDER BY LeadSourceNumber "
				
				
				Set cnnLostByLeadSource = Server.CreateObject("ADODB.Connection")
				cnnLostByLeadSource.open(Session("ClientCnnString"))
				Set rsLostByLeadSource = Server.CreateObject("ADODB.Recordset")
				rsLostByLeadSource.CursorLocation = 3 
				Set rsLostByLeadSource = cnnLostByLeadSource.Execute(SQLLostByLeadSourceReasons)
				
				If NOT rsLostByLeadSource.EOF Then
				
					showLostByLeadSourceChart = "True"
					CurrentLeadSourceNumber = ""
					leadSourceChange = 0
					
					Do While Not rsLostByLeadSource.EOF
					
						If CurrentLeadSourceNumber <> "" Then
							If CurrentLeadSourceNumber <> rsLostByLeadSource("LeadSourceNumber") Then
								CurrentLeadSourceNumber = rsLostByLeadSource("LeadSourceNumber")
								leadSourceChange = 1
							End If
						Else
							leadSourceChange = 1
							CurrentLeadSourceNumber = rsLostByLeadSource("LeadSourceNumber")
						End If
					
						LeadSourceNumber = rsLostByLeadSource("LeadSourceNumber")
						LeadSource = GetLeadSourceByNum(LeadSourceNumber)
					
						If rsLostByLeadSource("NumberOfProspects") > 0 Then
							
							If leadsourceChange = 1 Then
								%>
								<tr>
					                <td><%= LeadSource %></td>
					                <td><%= GetReasonByNum(rsLostByLeadSource("ReasonNo")) %> : <%= rsLostByLeadSource("NumberOfProspects") %></td>
					            </tr>
								<%
								leadsourceChange = 0
							Else
								%>
								<tr>
					                <td>&nbsp;</td>
					                <td><%= GetReasonByNum(rsLostByLeadSource("ReasonNo")) %> : <%= rsLostByLeadSource("NumberOfProspects") %></td>
					            </tr>
								<%						
							End If
						End If

					rsLostByLeadSource.MoveNext
				Loop
				
				Else
					%><tr><td colspan="2">No Data This Past Week</td></tr><%
				End If
				%>
		        </tbody>
		    </table>

	<hr>

<!--#include file="../../inc/footer-main.asp"-->


 