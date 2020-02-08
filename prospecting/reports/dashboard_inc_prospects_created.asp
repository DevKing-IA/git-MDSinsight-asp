<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For Prospects Created By Sales Rep -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%

SQLProspectSalesRep = "SELECT tblUsers.userNo, COUNT(PR_Prospects.CreatedByUserNo) AS SalesRepCount"
SQLProspectSalesRep = SQLProspectSalesRep & " FROM  tblUsers LEFT OUTER JOIN"
SQLProspectSalesRep = SQLProspectSalesRep & " PR_Prospects ON PR_Prospects.CreatedByUserNo = tblUsers.userNo"
SQLProspectSalesRep = SQLProspectSalesRep & " WHERE tblUsers.userNo IN (" & SelectedUserNumbersToDisplay & ") "
SQLProspectSalesRep = SQLProspectSalesRep & " AND  PR_Prospects.CreatedDate >= '" & mondayOfLastWeek & "' AND PR_Prospects.CreatedDate < '" & mondayOfThisWeek & "' "
SQLProspectSalesRep = SQLProspectSalesRep & " GROUP BY tblUsers.userNo"
SQLProspectSalesRep = SQLProspectSalesRep & " ORDER BY SalesRepCount DESC"

showProspectsCreatedBySalesRepChart = "False"

Set cnnProspectSalesRep = Server.CreateObject("ADODB.Connection")
cnnProspectSalesRep.open(Session("ClientCnnString"))
Set rsProspectSalesRep = Server.CreateObject("ADODB.Recordset")
rsProspectSalesRep.CursorLocation = 3 
Set rsProspectSalesRep = cnnProspectSalesRep.Execute(SQLProspectSalesRep)

If NOT rsProspectSalesRep.EOF Then

	colorCount = 0

	jChartDataSalesRep = ""
	showProspectsCreatedBySalesRepChart = "True"
	
	Do While Not rsProspectSalesRep.EOF

		SalesRepCount = rsProspectSalesRep("SalesRepCount")
		CreatedByUserNo = rsProspectSalesRep("userNo")
		CreatedByUserName = GetUserDisplayNameByUserNo(CreatedByUserNo)
		
		jChartDataSalesRep = jChartDataSalesRep & "{'salesrep':'" & CreatedByUserName & "','numprospects':" & SalesRepCount & ",'color':'" & barGraphColorArray12(colorCount) & "'},"
		
		colorCount = colorCount + 1
		
		If colorCount = 12 Then colorCount = 0
	
		rsProspectSalesRep.MoveNext
	Loop
	
	jChartDataSalesRep = Left(jChartDataSalesRep,Len(jChartDataSalesRep)-1)
		
End If

%>


<!-- Chart code -->
<script>


var chartSalesRep = AmCharts.makeChart("chartdivProspectsCreatedBySalesRep", {
  "type": "serial",
  "theme": "none",
  "marginRight": 70,
  "dataProvider": [<%= jChartDataSalesRep %>],
  "valueAxes": [{
    "axisAlpha": 0,
    "position": "left",
    "title": "# Prospects Created",
    "unit": " prospects",
    "integersOnly": true
  }],
  "startDuration": 1,
  "graphs": [{
    "balloonText": "<b>[[category]]: [[value]]</b>",
    "fillColorsField": "color",
    "fillAlphas": 0.9,
    "lineAlpha": 0.2,
    "type": "column",
    "valueField": "numprospects"
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
<!-- END Chart Code For Prospects Created By Sales Rep -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->




<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For Prospects Created By Lead Source -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%

SQLProspectLeadSource = "SELECT PR_LeadSources.InternalRecordIdentifier AS LeadSourceNum, COUNT(PR_Prospects.InternalRecordIdentifier) AS LeadSourceCount "
SQLProspectLeadSource = SQLProspectLeadSource & " FROM  PR_LeadSources LEFT OUTER JOIN"
SQLProspectLeadSource = SQLProspectLeadSource & " PR_Prospects ON PR_Prospects.LeadSourceNumber = PR_LeadSources.InternalRecordIdentifier"
SQLProspectLeadSource = SQLProspectLeadSource & " WHERE PR_Prospects.CreatedDate >= '" & mondayOfLastWeek & "' AND PR_Prospects.CreatedDate < '" & mondayOfThisWeek & "' "
SQLProspectLeadSource = SQLProspectLeadSource & " GROUP BY PR_LeadSources.InternalRecordIdentifier"
SQLProspectLeadSource = SQLProspectLeadSource & " ORDER BY LeadSourceCount DESC"

showProspectsCreatedByLeadSourceChart = "False"

Set cnnProspectLeadSource = Server.CreateObject("ADODB.Connection")
cnnProspectLeadSource.open(Session("ClientCnnString"))
Set rsProspectLeadSource = Server.CreateObject("ADODB.Recordset")
rsProspectLeadSource.CursorLocation = 3 
Set rsProspectLeadSource = cnnProspectLeadSource.Execute(SQLProspectLeadSource)

If NOT rsProspectLeadSource.EOF Then

	colorCount = 0

	jChartDataLeadSource = ""
	showProspectsCreatedByLeadSourceChart = "True"
	
	Do While Not rsProspectLeadSource.EOF

		LeadSourceCount = rsProspectLeadSource("LeadSourceCount")
		LeadSourceNumber = rsProspectLeadSource("LeadSourceNum")
		LeadSource = GetLeadSourceByNum(LeadSourceNumber)
		
		jChartDataLeadSource = jChartDataLeadSource & "{'leadsource':'" & LeadSource & "','numleads':" & LeadSourceCount & ",'color':'" & barGraphColorArray22(colorCount) & "'},"
		
		colorCount = colorCount + 1
		
		If colorCount = 22 Then colorCount = 0
	
		rsProspectLeadSource.MoveNext
	Loop
	
	jChartDataLeadSource = Left(jChartDataLeadSource,Len(jChartDataLeadSource)-1)
		
End If


%>
<!-- Chart code -->
<script>


var chartLeadSource = AmCharts.makeChart("chartdivProspectsCreatedByLeadSource", {
  "type": "serial",
  "theme": "none",
  "marginRight": 70,
  "dataProvider": [<%= jChartDataLeadSource  %>],
  "valueAxes": [{
    "axisAlpha": 0,
    "position": "left",
    "title": "# Prospects Created",
    "unit": " prospects",
    "integersOnly": true
  }],
  "startDuration": 1,
  "graphs": [{
    "balloonText": "<b>[[category]]: [[value]]</b>",
    "fillColorsField": "color",
    "fillAlphas": 0.9,
    "lineAlpha": 0.2,
    "type": "column",
    "valueField": "numleads"
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
<!-- END Chart Code For Prospects Created By Lead Source -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->



