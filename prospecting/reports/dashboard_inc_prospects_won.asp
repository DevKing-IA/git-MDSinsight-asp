



<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For New Clients Created By Sales Rep -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%
showNewClientsBySalesRepChart = "False"

SQLNewClientsBySalesRep = "SELECT tblUsers.userNo, COUNT(PR_Prospects.OwnerUserNo) AS SalesRepCount, Pool"
SQLNewClientsBySalesRep = SQLNewClientsBySalesRep & " FROM tblUsers LEFT OUTER JOIN"
SQLNewClientsBySalesRep = SQLNewClientsBySalesRep & " PR_Prospects ON PR_Prospects.CreatedByUserNo = tblUsers.userNo"
SQLNewClientsBySalesRep = SQLNewClientsBySalesRep & " WHERE  PR_Prospects.Pool='Won' AND tblUsers.userNo IN (" & SelectedUserNumbersToDisplay & ") "
SQLNewClientsBySalesRep = SQLNewClientsBySalesRep & " GROUP BY tblUsers.userNo, Pool "
SQLNewClientsBySalesRep = SQLNewClientsBySalesRep & " ORDER BY SalesRepCount DESC"

'Response.write("<br><br><br><br>" & SQLNewClientsBySalesRep)


Set cnnNewClientsBySalesRep = Server.CreateObject("ADODB.Connection")
cnnNewClientsBySalesRep.open(Session("ClientCnnString"))
Set rsNewClientsBySalesRep = Server.CreateObject("ADODB.Recordset")
rsNewClientsBySalesRep.CursorLocation = 3 
Set rsNewClientsBySalesRep = cnnNewClientsBySalesRep.Execute(SQLNewClientsBySalesRep)

If NOT rsNewClientsBySalesRep.EOF Then

	colorCount = 0
	showNewClientsBySalesRepChart = "True"

	jChartDataSalesRepWon = ""
	Do While Not rsNewClientsBySalesRep.EOF

		SalesRepCount = rsNewClientsBySalesRep("SalesRepCount")
		OwnerUserNo = rsNewClientsBySalesRep("userNo")
		OwnerUserName = GetUserDisplayNameByUserNo(OwnerUserNo)
		
		jChartDataSalesRepWon = jChartDataSalesRepWon & "{'salesrep':'" & OwnerUserName & "','numprospects':" & SalesRepCount & ",'color':'" & barGraphColorArray12(colorCount) & "'},"
		
		colorCount = colorCount + 1
		
		If colorCount = 12 Then colorCount = 0
	
		rsNewClientsBySalesRep.MoveNext
	Loop
	
	jChartDataSalesRepWon = Left(jChartDataSalesRepWon,Len(jChartDataSalesRepWon)-1)
		
End If

%>


<!-- Chart code -->
<script>


var chartSalesRep = AmCharts.makeChart("chartdivNewClientsBySalesRep", {
  "type": "serial",
  "theme": "none",
  "marginRight": 70,
  "dataProvider": [<%= jChartDataSalesRepWon %>],
  "valueAxes": [{
    "axisAlpha": 0,
    "position": "left",
    "title": "# Clients Converted to Customers",
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
<!-- END Chart Code For New Clients Created By Sales Rep -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->



<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For New Clients Created By Lead Source -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%

showNewClientsByLeadSourceChart = "False"

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

	colorCount = 0
	showNewClientsByLeadSourceChart = "True"
	jChartDataLeadSourceNew = ""
	
	Do While Not rsNewClientsByLeadSource.EOF

		LeadSourceCount = rsNewClientsByLeadSource("LeadSourceCount")
		LeadSourceNumber = rsNewClientsByLeadSource("LeadSourceNumber")
		LeadSource = GetLeadSourceByNum(LeadSourceNumber)
		
		jChartDataLeadSourceNew = jChartDataLeadSourceNew & "{'leadsource':'" & LeadSource & "','numleads':" & LeadSourceCount & ",'color':'" & barGraphColorArray22(colorCount) & "'},"
		
		colorCount = colorCount + 1
		
		If colorCount = 22 Then colorCount = 0
	
		rsNewClientsByLeadSource.MoveNext
	Loop
	
	jChartDataLeadSourceNew  = Left(jChartDataLeadSourceNew,Len(jChartDataLeadSourceNew)-1)
		
End If


%>
<!-- Chart code -->
<script>


var chartLeadSource = AmCharts.makeChart("chartdivNewClientsByLeadSource", {
  "type": "serial",
  "theme": "none",
  "marginRight": 70,
  "dataProvider": [<%= jChartDataLeadSourceNew  %>],
  "valueAxes": [{
    "axisAlpha": 0,
    "position": "left",
    "title": "# Clients Converted to Customers",
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
<!-- END Chart Code For New Clients Created By Lead Source -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->



