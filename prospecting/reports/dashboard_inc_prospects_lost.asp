

<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For Lost Prospects By Sales Rep And Reason -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%
showLostBySalesRepChart = "False"


SQLLostBySalesRepLegend = "SELECT DISTINCT LastReasonNumber, Reason FROM PR_DashboardDetailsUQ_LastWeek INNER JOIN"
SQLLostBySalesRepLegend = SQLLostBySalesRepLegend & " PR_Reasons ON PR_DashboardDetailsUQ_LastWeek.LastReasonNumber = PR_Reasons.InternalRecordIdentifier "
SQLLostBySalesRepLegend = SQLLostBySalesRepLegend & " WHERE LastStageNumber = 1 "
SQLLostBySalesRepLegend = SQLLostBySalesRepLegend & " ORDER BY Reason "


Set cnnLostBySalesRepLegend = Server.CreateObject("ADODB.Connection")
cnnLostBySalesRepLegend.open(Session("ClientCnnString"))
Set rsLostBySalesRepLegend = Server.CreateObject("ADODB.Recordset")
rsLostBySalesRepLegend.CursorLocation = 3 
Set rsLostBySalesRepLegend = cnnLostBySalesRepLegend.Execute(SQLLostBySalesRepLegend)

If NOT rsLostBySalesRepLegend.EOF Then

	jChartLegendSalesRepLost = ""
	Do While Not rsLostBySalesRepLegend.EOF
	
		jChartLegendSalesRepLost = jChartLegendSalesRepLost & "{'balloonText':'<b>[[title]]</b><br><span style=font-size:14px>[[category]]: <b>[[value]]</b></span>',"
		jChartLegendSalesRepLost = jChartLegendSalesRepLost & "'fillAlphas': 0.8,"
		jChartLegendSalesRepLost = jChartLegendSalesRepLost & "'labelText': '[[value]]',"
		jChartLegendSalesRepLost = jChartLegendSalesRepLost & "'lineAlpha': 0.3,"
		jChartLegendSalesRepLost = jChartLegendSalesRepLost & "'title': '" & rsLostBySalesRepLegend("Reason") & "',"
		jChartLegendSalesRepLost = jChartLegendSalesRepLost & "'type': 'column',"
		jChartLegendSalesRepLost = jChartLegendSalesRepLost & "'color': '#000000',"
		jChartLegendSalesRepLost = jChartLegendSalesRepLost & "'valueField': '" & rsLostBySalesRepLegend("LastReasonNumber") & "'},"
		
		rsLostBySalesRepLegend.MoveNext
	Loop
	
	jChartLegendSalesRepLost = Left(jChartLegendSalesRepLost,Len(jChartLegendSalesRepLost)-1)
	
End If




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
				jChartDataSalesRepLost = jChartDataSalesRepLost & "},"
				CurrentOwnerUserNo = rsLostBySalesRep("OwnerUserNo")
				repChange = 1
			Else
				jChartDataSalesRepLost = jChartDataSalesRepLost & ","
			End If
		Else
			repChange = 1
			CurrentOwnerUserNo = rsLostBySalesRep("OwnerUserNo")
		End If
	
		OwnerUserNo = rsLostBySalesRep("OwnerUserNo")
		OwnerUserName = GetUserDisplayNameByUserNo(OwnerUserNo)
	
		If repChange = 1 Then
			jChartDataSalesRepLost = jChartDataSalesRepLost & "{'salesrep':'" & OwnerUserName & "',"
			repChange = 0
		End If
		jChartDataSalesRepLost = jChartDataSalesRepLost & "'" & rsLostBySalesRep("ReasonNo") & "': " & rsLostBySalesRep("NumberOfProspects") 
		 
	 			
		rsLostBySalesRep.MoveNext
	Loop
	
	jChartDataSalesRepLost = jChartDataSalesRepLost & "},"
	jChartDataSalesRepLost = Left(jChartDataSalesRepLost,Len(jChartDataSalesRepLost)-1)
		
End If


%>


<!-- Chart code -->
<script>
var chartLostSalesRepReasons = AmCharts.makeChart("chartdivLostByReasonSalesRep", {
    "type": "serial",
	"theme": "none",
    "legend": {
        "horizontalGap": 10,
        "maxColumns": 1,
        "position": "right",
		"useGraphSettings": true,
		"markerSize": 10,
		"switchable": true,
    },
    "dataProvider": [<%= jChartDataSalesRepLost %>],
    "valueAxes": [{
        "stackType": "regular",
        "axisAlpha": 0.3,
        "gridAlpha": 0,
        "unit": " prospects",
        "integersOnly": true,
		"autoGridCount": false,
    	"gridCount": 5        
    }],    
    "graphs": [<%= jChartLegendSalesRepLost %>],
    "rotate": false,
    "categoryField": "salesrep",
    "categoryAxis": {
        "gridPosition": "start",
        "axisAlpha": 0,
        "gridAlpha": 0,
        "position": "left"
    },
    "export": {
    	"enabled": true
     }

});
</script>
	

<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- END Chart Code For Lost Prospects By Sales Rep and Reason -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->



<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For Lost Prospects By Lead Source And Reason -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%
showLostByLeadSourceChart = "False"


SQLLostByLeadSourceLegend = "SELECT DISTINCT LastReasonNumber, Reason FROM PR_DashboardDetailsUQ_LastWeek INNER JOIN"
SQLLostByLeadSourceLegend = SQLLostByLeadSourceLegend & " PR_Reasons ON PR_DashboardDetailsUQ_LastWeek.LastReasonNumber = PR_Reasons.InternalRecordIdentifier "
SQLLostByLeadSourceLegend = SQLLostByLeadSourceLegend & " WHERE LastStageNumber = 1 "
SQLLostByLeadSourceLegend = SQLLostByLeadSourceLegend & " ORDER BY Reason "


Set cnnLostByLeadSourceLegend = Server.CreateObject("ADODB.Connection")
cnnLostByLeadSourceLegend.open(Session("ClientCnnString"))
Set rsLostByLeadSourceLegend = Server.CreateObject("ADODB.Recordset")
rsLostByLeadSourceLegend.CursorLocation = 3 
Set rsLostByLeadSourceLegend = cnnLostByLeadSourceLegend.Execute(SQLLostByLeadSourceLegend)

If NOT rsLostByLeadSourceLegend.EOF Then

	jChartLegendLeadSourceLost = ""
	colorCount = 0

	Do While Not rsLostByLeadSourceLegend.EOF
	
		jChartLegendLeadSourceLost = jChartLegendLeadSourceLost & "{'balloonText':'<b>[[title]]</b><br><span style=font-size:14px>[[category]]: <b>[[value]]</b></span>',"
		jChartLegendLeadSourceLost = jChartLegendLeadSourceLost & "'fillAlphas': 0.8,"
		jChartLegendLeadSourceLost = jChartLegendLeadSourceLost & "'labelText': '[[value]]',"
		jChartLegendLeadSourceLost = jChartLegendLeadSourceLost & "'lineAlpha': 0.3,"
		jChartLegendLeadSourceLost = jChartLegendLeadSourceLost & "'title': '" & rsLostByLeadSourceLegend("Reason") & "',"
		jChartLegendLeadSourceLost = jChartLegendLeadSourceLost & "'type': 'column',"
		jChartLegendLeadSourceLost = jChartLegendLeadSourceLost & "'color': '#000000',"
		'jChartLegendLeadSourceLost = jChartLegendLeadSourceLost & "'lineColor':'" & stackedBarGraphColorArray10(colorCount) & "',"
		jChartLegendLeadSourceLost = jChartLegendLeadSourceLost & "'valueField': '" & rsLostByLeadSourceLegend("LastReasonNumber") & "'},"
	
 		colorCount = colorCount + 1
		If colorCount = 10 Then colorCount = 0
	
		rsLostByLeadSourceLegend.MoveNext
	Loop
	
	jChartLegendLeadSourceLost = Left(jChartLegendLeadSourceLost,Len(jChartLegendLeadSourceLost)-1)
	
End If






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
	jChartDataLeadSourceLost = ""
	CurrentLeadSourceNumber = ""
	repChange = 0
	
	Do While Not rsLostByLeadSource.EOF
	
		If CurrentLeadSourceNumber <> "" Then
			If CurrentLeadSourceNumber <> rsLostByLeadSource("LeadSourceNumber") Then
				jChartDataLeadSourceLost = jChartDataLeadSourceLost & "},"
				CurrentLeadSourceNumber = rsLostByLeadSource("LeadSourceNumber")
				repChange = 1
			Else
				jChartDataLeadSourceLost = jChartDataLeadSourceLost & ","
			End If
		Else
			repChange = 1
			CurrentLeadSourceNumber = rsLostByLeadSource("LeadSourceNumber")
		End If
	
		LeadSourceNumber = rsLostByLeadSource("LeadSourceNumber")
		LeadSource = GetLeadSourceByNum(LeadSourceNumber)
	
		If repChange = 1 Then
			jChartDataLeadSourceLost = jChartDataLeadSourceLost & "{'LeadSource':'" & LeadSource & "',"
			repChange = 0
		End If
		jChartDataLeadSourceLost = jChartDataLeadSourceLost & "'" & rsLostByLeadSource("ReasonNo") & "': " & rsLostByLeadSource("NumberOfProspects") 
		 
	 			
		rsLostByLeadSource.MoveNext
	Loop
	
	jChartDataLeadSourceLost = jChartDataLeadSourceLost & "},"
	jChartDataLeadSourceLost = Left(jChartDataLeadSourceLost,Len(jChartDataLeadSourceLost)-1)
		
End If


%>


<!-- Chart code -->
<script>
var chartLostLeadSourceReasons = AmCharts.makeChart("chartdivLostByReasonLeadSource", {
    "type": "serial",
	"theme": "none",
    "legend": {
        "horizontalGap": 10,
        "maxColumns": 7,
        "position": "top",
		"useGraphSettings": true,
		"markerSize": 10,
		"switchable": true,
    },
    "dataProvider": [<%= jChartDataLeadSourceLost %>],
    "valueAxes": [{
        "stackType": "regular",
        "axisAlpha": 0.3,
        "gridAlpha": 0,
        "unit": " prospects",
        "integersOnly": true    
    }],    
    "graphs": [<%= jChartLegendLeadSourceLost %>],
    "rotate": false,
    "categoryField": "LeadSource",
    "categoryAxis": {
        "gridPosition": "start",
        "axisAlpha": 0,
        "gridAlpha": 0,
        "position": "left",
        "labelRotation": 45,
    },
    "export": {
    	"enabled": true
     }

});
</script>
	

<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- END Chart Code For Lost Prospects By Lead Source and Reason -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->


