


<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For Prospects Qualified By Sales Rep -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%
showQualifiedClientsBySalesRepChart = "False"

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

	colorCount = 0
	showQualifiedClientsBySalesRepChart = "True"

	jChartDataSalesRepQualified = ""
	Do While Not rsQualifiedClientsBySalesRep.EOF

		SalesRepCount = rsQualifiedClientsBySalesRep("SalesRepCount")
		OwnerUserNo = rsQualifiedClientsBySalesRep("userNo")
		OwnerUserName = GetUserDisplayNameByUserNo(OwnerUserNo)
		
		jChartDataSalesRepQualified = jChartDataSalesRepQualified & "{'salesrep':'" & OwnerUserName & "','numprospects':" & SalesRepCount & ",'color':'" & barGraphColorArray12(colorCount) & "'},"
		
		colorCount = colorCount + 1
		
		If colorCount = 12 Then colorCount = 0
	
		rsQualifiedClientsBySalesRep.MoveNext
	Loop
	
	jChartDataSalesRepQualified = Left(jChartDataSalesRepQualified,Len(jChartDataSalesRepQualified)-1)
		
End If

%>


<!-- Chart code -->
<script>


var chartSalesRepQualified = AmCharts.makeChart("chartdivQualifiedClientsBySalesRep", {
  "type": "serial",
  "theme": "none",
  "marginRight": 70,
  "dataProvider": [<%= jChartDataSalesRepQualified %>],
  "valueAxes": [{
    "axisAlpha": 0,
    "position": "left",
    "title": "# Prospects Qualified",
    "unit": " clients",
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
<!-- END Chart Code For Prospects Qualified By Sales Rep -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->



<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For Prospects Qualified By Lead Source -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%

showQualifiedClientsByLeadSourceChart = "False"

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

	colorCount = 0
	showQualifiedClientsByLeadSourceChart = "True"
	jChartDataLeadSourceQualified = ""
	
	Do While Not rsQualifiedClientsByLeadSource.EOF

		LeadSourceCount = rsQualifiedClientsByLeadSource("LeadSourceCount")
		LeadSourceNumber = rsQualifiedClientsByLeadSource("LeadSourceNumber")
		LeadSource = GetLeadSourceByNum(LeadSourceNumber)
		
		jChartDataLeadSourceQualified = jChartDataLeadSourceQualified & "{'leadsource':'" & LeadSource & "','numleads':" & LeadSourceCount & ",'color':'" & barGraphColorArray22(colorCount) & "'},"
		
		colorCount = colorCount + 1
		
		If colorCount = 22 Then colorCount = 0
	
		rsQualifiedClientsByLeadSource.MoveNext
	Loop
	
	jChartDataLeadSourceQualified  = Left(jChartDataLeadSourceQualified,Len(jChartDataLeadSourceQualified)-1)
		
End If


%>
<!-- Chart code -->
<script>


var chartLeadSourceQualified = AmCharts.makeChart("chartdivQualifiedClientsByLeadSource", {
  "type": "serial",
  "theme": "none",
  "marginRight": 70,
  "dataProvider": [<%= jChartDataLeadSourceQualified  %>],
  "valueAxes": [{
    "axisAlpha": 0,
    "position": "left",
    "title": "# Prospects Qualified",
    "unit": " clients",
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
<!-- END Chart Code For Qualified Prospects Created By Lead Source -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->

