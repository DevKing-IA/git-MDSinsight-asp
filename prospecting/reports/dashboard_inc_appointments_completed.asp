

<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For Appointments Completed By Sales Rep -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%

SQLAppmtCompleteSalesRep = "SELECT tblUsers.userNo, COUNT(PR_ProspectActivities.ActivityCreatedByUserNo) AS AppmtCount"
SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " FROM  tblUsers LEFT OUTER JOIN"
SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " PR_ProspectActivities ON PR_ProspectActivities.ActivityCreatedByUserNo = tblUsers.userNo"
SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " WHERE tblUsers.userNo IN (" & SelectedUserNumbersToDisplay & ") AND "
SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " PR_ProspectActivities.ActivityIsMeeting=1 AND PR_ProspectActivities.Status='Completed' "
SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " AND PR_ProspectActivities.StatusDateTime >= '" & mondayOfLastWeek & "' AND PR_ProspectActivities.StatusDateTime < '" & mondayOfThisWeek & "' "
SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " GROUP BY tblUsers.userNo"
SQLAppmtCompleteSalesRep = SQLAppmtCompleteSalesRep & " ORDER BY AppmtCount DESC"

showAppmtsCompletedBySalesRepChart = "False"

Set cnnAppmtCompleteSalesRep = Server.CreateObject("ADODB.Connection")
cnnAppmtCompleteSalesRep.open(Session("ClientCnnString"))
Set rsAppmtCompleteSalesRep = Server.CreateObject("ADODB.Recordset")
rsAppmtCompleteSalesRep.CursorLocation = 3 
Set rsAppmtCompleteSalesRep = cnnAppmtCompleteSalesRep.Execute(SQLAppmtCompleteSalesRep)

If NOT rsAppmtCompleteSalesRep.EOF Then

	colorCount = 0

	jChartDataAppmtCompleteSalesRep = ""
	showAppmtsCompletedBySalesRepChart = "True"
	
	Do While Not rsAppmtCompleteSalesRep.EOF

		AppmtCount = rsAppmtCompleteSalesRep("AppmtCount")
		CreatedByUserNo = rsAppmtCompleteSalesRep("userNo")
		CreatedByUserName = GetUserDisplayNameByUserNo(CreatedByUserNo)
		
		jChartDataAppmtCompleteSalesRep = jChartDataAppmtCompleteSalesRep & "{'salesrep':'" & CreatedByUserName & "','numappmts':" & AppmtCount & ",'color':'" & barGraphColorArray12(colorCount) & "'},"
		
		colorCount = colorCount + 1
		
		If colorCount = 12 Then colorCount = 0
	
		rsAppmtCompleteSalesRep.MoveNext
	Loop
	
	jChartDataAppmtCompleteSalesRep = Left(jChartDataAppmtCompleteSalesRep,Len(jChartDataAppmtCompleteSalesRep)-1)
		
End If

%>


<!-- Chart code -->
<script>


var chartSalesRepAppointmentsAttended = AmCharts.makeChart("chartdivAppointmentsAttendedBySalesRep", {
  "type": "serial",
  "theme": "none",
  "marginRight": 70,
  "dataProvider": [<%= jChartDataAppmtCompleteSalesRep %>],
  "valueAxes": [{
    "axisAlpha": 0,
    "position": "left",
    "title": "# Meetings Attended",
    "unit": " appmts",
    "integersOnly": true
  }],
  "startDuration": 1,
  "graphs": [{
    "balloonText": "<b>[[category]]: [[value]]</b>",
    "fillColorsField": "color",
    "fillAlphas": 0.9,
    "lineAlpha": 0.2,
    "type": "column",
    "valueField": "numappmts"
  }],
  "chartCursor": {
    "categoryBalloonEnabled": false,
    "cursorAlpha": 0,
    "zoomable": false
  },
  "categoryField": "salesrep",
  "categoryAxis": {
    "gridPosition": "start",
    "labelRotation": 45
  },
  "export": {
    "enabled": true
  }

});
</script>
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- END Chart Code For Appointments Completed By Sales Rep -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->




<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For Appointments Completed By Lead Source -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%

SQLAppmtCompleteLeadSource = "SELECT PR_LeadSources.InternalRecordIdentifier AS LeadSourceNum, COUNT(PR_ProspectActivities.ProspectRecID) AS AppmtCount "
SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " FROM  PR_LeadSources LEFT OUTER JOIN"
SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " PR_Prospects ON PR_Prospects.LeadSourceNumber = PR_LeadSources.InternalRecordIdentifier "
SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " INNER JOIN PR_ProspectActivities ON PR_ProspectActivities.ProspectRecID = PR_Prospects.InternalRecordIdentifier "
SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " WHERE PR_ProspectActivities.ActivityIsMeeting=1 and PR_ProspectActivities.Status='Completed'"
SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " AND PR_ProspectActivities.StatusDateTime >='" & mondayOfLastWeek & "' AND PR_ProspectActivities.StatusDateTime <'" & mondayOfThisWeek & "' "
SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " GROUP BY PR_LeadSources.InternalRecordIdentifier"
SQLAppmtCompleteLeadSource = SQLAppmtCompleteLeadSource & " ORDER BY AppmtCount DESC"

showAppmtsCompletedByLeadSourceChart = "False"

Set cnnAppmtCompleteLeadSource = Server.CreateObject("ADODB.Connection")
cnnAppmtCompleteLeadSource.open(Session("ClientCnnString"))
Set rsAppmtCompleteLeadSource = Server.CreateObject("ADODB.Recordset")
rsAppmtCompleteLeadSource.CursorLocation = 3 
Set rsAppmtCompleteLeadSource = cnnAppmtCompleteLeadSource.Execute(SQLAppmtCompleteLeadSource)

If NOT rsAppmtCompleteLeadSource.EOF Then

	colorCount = 0

	jChartDataAppmtCompleteLeadSource = ""
	showAppmtsCompletedByLeadSourceChart = "True"
	
	Do While Not rsAppmtCompleteLeadSource.EOF

		AppmtCount = rsAppmtCompleteLeadSource("AppmtCount")
		LeadSourceNumber = rsAppmtCompleteLeadSource("LeadSourceNum")
		LeadSource = GetLeadSourceByNum(LeadSourceNumber)
		
		jChartDataAppmtCompleteLeadSource = jChartDataAppmtCompleteLeadSource & "{'leadsource':'" & LeadSource & "','numappmts':" & AppmtCount & ",'color':'" & barGraphColorArray22(colorCount) & "'},"
		
		colorCount = colorCount + 1
		
		If colorCount = 22 Then colorCount = 0
	
		rsAppmtCompleteLeadSource.MoveNext
	Loop
	
	jChartDataAppmtCompleteLeadSource = Left(jChartDataAppmtCompleteLeadSource,Len(jChartDataAppmtCompleteLeadSource)-1)
		
End If


%>
<!-- Chart code -->
<script>


var chartLeadSourceAppointmentsAttended = AmCharts.makeChart("chartdivAppointmentsAttendedByLeadSource", {
  "type": "serial",
  "theme": "none",
  "marginRight": 70,
  "dataProvider": [<%= jChartDataAppmtCompleteLeadSource  %>],
  "valueAxes": [{
    "axisAlpha": 0,
    "position": "left",
    "title": "# Meetings Attended",
    "unit": " appmts",
    "integersOnly": true
  }],
  "startDuration": 1,
  "graphs": [{
    "balloonText": "<b>[[category]]: [[value]]</b>",
    "fillColorsField": "color",
    "fillAlphas": 0.9,
    "lineAlpha": 0.2,
    "type": "column",
    "valueField": "numappmts"
  }],
  "chartCursor": {
    "categoryBalloonEnabled": false,
    "cursorAlpha": 0,
    "zoomable": false
  },
  "categoryField": "leadsource",
  "categoryAxis": {
    "gridPosition": "start",
    "labelRotation": 45
  },
  "export": {
    "enabled": true
  }

});
</script>
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- END Chart Code For Appointments Completed By Lead Source -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->






