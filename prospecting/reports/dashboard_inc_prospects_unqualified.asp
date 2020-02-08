

<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For Unqualified Prospects By Sales Rep And Reason -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%
showUnqualifiedBySalesRepChart = "False"

'Format For jQuery Chart
'-----------------------------------------------------
'SalesRepUserNo/Name
'Unqualified Reason 1 - Total Count For Reason 1
'Unqualified Reason 2 - Total Count For Reason 2
'Unqualified Reason 3 - Total Count For Reason 3
'ETC....
'-----------------------------------------------------
'Reasons MUST Be The Same For Every Sales Rep
'-----------------------------------------------------

SQLUnqualifiedBySalesRepLegend = "SELECT DISTINCT LastReasonNumber, Reason FROM PR_DashboardDetailsUQ_LastWeek INNER JOIN"
SQLUnqualifiedBySalesRepLegend = SQLUnqualifiedBySalesRepLegend & " PR_Reasons ON PR_DashboardDetailsUQ_LastWeek.LastReasonNumber = PR_Reasons.InternalRecordIdentifier "
SQLUnqualifiedBySalesRepLegend = SQLUnqualifiedBySalesRepLegend & " WHERE LastStageNumber = 0 "
SQLUnqualifiedBySalesRepLegend = SQLUnqualifiedBySalesRepLegend & " ORDER BY Reason "


Set cnnUnqualifiedBySalesRepLegend = Server.CreateObject("ADODB.Connection")
cnnUnqualifiedBySalesRepLegend.open(Session("ClientCnnString"))
Set rsUnqualifiedBySalesRepLegend = Server.CreateObject("ADODB.Recordset")
rsUnqualifiedBySalesRepLegend.CursorLocation = 3 
Set rsUnqualifiedBySalesRepLegend = cnnUnqualifiedBySalesRepLegend.Execute(SQLUnqualifiedBySalesRepLegend)

If NOT rsUnqualifiedBySalesRepLegend.EOF Then

	jChartLegendSalesRepUnqualified = ""

	Do While Not rsUnqualifiedBySalesRepLegend.EOF
	
		jChartLegendSalesRepUnqualified = jChartLegendSalesRepUnqualified & "{'balloonText':'<b>[[title]]</b><br><span style=font-size:14px>[[category]]: <b>[[value]]</b></span>',"
		jChartLegendSalesRepUnqualified = jChartLegendSalesRepUnqualified & "'fillAlphas': 0.8,"
		jChartLegendSalesRepUnqualified = jChartLegendSalesRepUnqualified & "'labelText': '[[value]]',"
		jChartLegendSalesRepUnqualified = jChartLegendSalesRepUnqualified & "'lineAlpha': 0.3,"
		jChartLegendSalesRepUnqualified = jChartLegendSalesRepUnqualified & "'title': '" & rsUnqualifiedBySalesRepLegend("Reason") & "',"
		jChartLegendSalesRepUnqualified = jChartLegendSalesRepUnqualified & "'type': 'column',"
		jChartLegendSalesRepUnqualified = jChartLegendSalesRepUnqualified & "'color': '#000000',"
		jChartLegendSalesRepUnqualified = jChartLegendSalesRepUnqualified & "'valueField': '" & rsUnqualifiedBySalesRepLegend("LastReasonNumber") & "'},"
		
		rsUnqualifiedBySalesRepLegend.MoveNext
	Loop
	
	jChartLegendSalesRepUnqualified = Left(jChartLegendSalesRepUnqualified,Len(jChartLegendSalesRepUnqualified)-1)
	
End If





SQLUnqualifiedBySalesRepReasons = "SELECT * "
SQLUnqualifiedBySalesRepReasons = SQLUnqualifiedBySalesRepReasons & " FROM  PR_DashboardSummaryByOwnerUQ_LastWeek "
SQLUnqualifiedBySalesRepReasons = SQLUnqualifiedBySalesRepReasons & " WHERE OwnerUserNo IN (" & SelectedUserNumbersToDisplay & ") "
SQLUnqualifiedBySalesRepReasons = SQLUnqualifiedBySalesRepReasons & " AND LastStageNumber = 0 "
SQLUnqualifiedBySalesRepReasons = SQLUnqualifiedBySalesRepReasons & " ORDER BY OwnerUserNo, ReasonNo "

showUnqualifiedByLeadSourceChart = "False"

Set cnnUnqualifiedBySalesRep = Server.CreateObject("ADODB.Connection")
cnnUnqualifiedBySalesRep.open(Session("ClientCnnString"))
Set rsUnqualifiedBySalesRep = Server.CreateObject("ADODB.Recordset")
rsUnqualifiedBySalesRep.CursorLocation = 3 
Set rsUnqualifiedBySalesRep = cnnUnqualifiedBySalesRep.Execute(SQLUnqualifiedBySalesRepReasons)

If NOT rsUnqualifiedBySalesRep.EOF Then

	showUnqualifiedByLeadSourceChart = "True"
	jChartDataSalesRepUnqualified = ""
	CurrentOwnerUserNo = ""
	repChange = 0
	
	Do While Not rsUnqualifiedBySalesRep.EOF
	
		If CurrentOwnerUserNo <> "" Then
			If CurrentOwnerUserNo <> rsUnqualifiedBySalesRep("OwnerUserNo") Then
				jChartDataSalesRepUnqualified = jChartDataSalesRepUnqualified & "},"
				CurrentOwnerUserNo = rsUnqualifiedBySalesRep("OwnerUserNo")
				repChange = 1
			Else
				jChartDataSalesRepUnqualified = jChartDataSalesRepUnqualified & ","
			End If
		Else
			repChange = 1
			CurrentOwnerUserNo = rsUnqualifiedBySalesRep("OwnerUserNo")
		End If
	
		OwnerUserNo = rsUnqualifiedBySalesRep("OwnerUserNo")
		OwnerUserName = GetUserDisplayNameByUserNo(OwnerUserNo)
	
		If repChange = 1 Then
			jChartDataSalesRepUnqualified = jChartDataSalesRepUnqualified & "{'salesrep':'" & OwnerUserName & "',"
			repChange = 0
		End If
		jChartDataSalesRepUnqualified = jChartDataSalesRepUnqualified & "'" & rsUnqualifiedBySalesRep("ReasonNo") & "': " & rsUnqualifiedBySalesRep("NumberOfProspects") 
		 
	 			
		rsUnqualifiedBySalesRep.MoveNext
	Loop
	
	jChartDataSalesRepUnqualified = jChartDataSalesRepUnqualified & "},"
	jChartDataSalesRepUnqualified = Left(jChartDataSalesRepUnqualified,Len(jChartDataSalesRepUnqualified)-1)
		
End If


%>


<!-- Chart code -->
<script>
var chartUnqualifiedSalesRepReasons = AmCharts.makeChart("chartdivUnqualifiedByReasonSalesRep", {
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
    "dataProvider": [<%= jChartDataSalesRepUnqualified %>],
    "valueAxes": [{
        "stackType": "regular",
        "axisAlpha": 0.3,
        "gridAlpha": 0,
        "unit": " prospects",
        "integersOnly": true,
		"autoGridCount": false,
    	"gridCount": 20        
    }],    
    "graphs": [<%= jChartLegendSalesRepUnqualified %>],
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
<!-- END Chart Code For Unqualified Prospects By Sales Rep and Reason -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->






<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- Chart Code For Unqualified Prospects By Lead Source And Reason -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->
<%
showUnqualifiedByLeadSourceChart = "False"


SQLUnqualifiedByLeadSourceLegend = "SELECT DISTINCT LastReasonNumber, Reason FROM PR_DashboardDetailsUQ_LastWeek INNER JOIN"
SQLUnqualifiedByLeadSourceLegend = SQLUnqualifiedByLeadSourceLegend & " PR_Reasons ON PR_DashboardDetailsUQ_LastWeek.LastReasonNumber = PR_Reasons.InternalRecordIdentifier "
SQLUnqualifiedByLeadSourceLegend = SQLUnqualifiedByLeadSourceLegend & " WHERE LastStageNumber = 0 ORDER BY Reason "


Set cnnUnqualifiedByLeadSourceLegend = Server.CreateObject("ADODB.Connection")
cnnUnqualifiedByLeadSourceLegend.open(Session("ClientCnnString"))
Set rsUnqualifiedByLeadSourceLegend = Server.CreateObject("ADODB.Recordset")
rsUnqualifiedByLeadSourceLegend.CursorLocation = 3 
Set rsUnqualifiedByLeadSourceLegend = cnnUnqualifiedByLeadSourceLegend.Execute(SQLUnqualifiedByLeadSourceLegend)

If NOT rsUnqualifiedByLeadSourceLegend.EOF Then

	jChartLegendLeadSourceUnqualified = ""
	colorCount = 0

	Do While Not rsUnqualifiedByLeadSourceLegend.EOF
	
		jChartLegendLeadSourceUnqualified = jChartLegendLeadSourceUnqualified & "{'balloonText':'<b>[[title]]</b><br><span style=font-size:14px>[[category]]: <b>[[value]]</b></span>',"
		jChartLegendLeadSourceUnqualified = jChartLegendLeadSourceUnqualified & "'fillAlphas': 0.8,"
		jChartLegendLeadSourceUnqualified = jChartLegendLeadSourceUnqualified & "'labelText': '[[value]]',"
		jChartLegendLeadSourceUnqualified = jChartLegendLeadSourceUnqualified & "'lineAlpha': 0.3,"
		jChartLegendLeadSourceUnqualified = jChartLegendLeadSourceUnqualified & "'title': '" & rsUnqualifiedByLeadSourceLegend("Reason") & "',"
		jChartLegendLeadSourceUnqualified = jChartLegendLeadSourceUnqualified & "'type': 'column',"
		jChartLegendLeadSourceUnqualified = jChartLegendLeadSourceUnqualified & "'color': '#FFFFFF',"
		'jChartLegendLeadSourceUnqualified = jChartLegendLeadSourceUnqualified & "'lineColor':'" & barGraphColorArray22(colorCount) & "',"
		jChartLegendLeadSourceUnqualified = jChartLegendLeadSourceUnqualified & "'valueField': '" & rsUnqualifiedByLeadSourceLegend("LastReasonNumber") & "'},"

 		colorCount = colorCount + 1
		If colorCount = 22 Then colorCount = 0
		
		rsUnqualifiedByLeadSourceLegend.MoveNext
	Loop
	
	jChartLegendLeadSourceUnqualified = Left(jChartLegendLeadSourceUnqualified,Len(jChartLegendLeadSourceUnqualified)-1)
	
End If



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
				jChartDataLeadSourceUnqualified = jChartDataLeadSourceUnqualified & "},"
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
		
	
		If leadsourceChange = 1 Then
			jChartDataLeadSourceUnqualified = jChartDataLeadSourceUnqualified & "{'leadsource':'" & LeadSource & "',"
			leadsourceChange = 0
		End If
		jChartDataLeadSourceUnqualified = jChartDataLeadSourceUnqualified & "'" & rsUnqualifiedByLeadSource("ReasonNo") & "': " & rsUnqualifiedByLeadSource("NumberOfProspects")	 
	 			
		rsUnqualifiedByLeadSource.MoveNext
	Loop
	
	jChartDataLeadSourceUnqualified = jChartDataLeadSourceUnqualified & "},"
	jChartDataLeadSourceUnqualified = Left(jChartDataLeadSourceUnqualified,Len(jChartDataLeadSourceUnqualified)-1)
		
End If


%>


<!-- Chart code -->
<script>
var chartUnqualifiedLeadSourceReasons = AmCharts.makeChart("chartdivUnqualifiedByReasonLeadSource", {
    "type": "serial",
	"theme": "none",
    "legend": {
        "horizontalGap": 10,
        "maxColumns": 7,
        "position": "right",
		"useGraphSettings": true,
		"markerSize": 10,
		"switchable": true,
    },
    "dataProvider": [<%= jChartDataLeadSourceUnqualified %>],
    "valueAxes": [{
        "stackType": "regular",
        "axisAlpha": 0,
        "gridAlpha": 0,
        "unit": " prospects",
        "integersOnly": true    
    }],    
    "graphs": [<%= jChartLegendLeadSourceUnqualified %>],
    "rotate": false,
    "categoryField": "leadsource",
    "categoryAxis": {
        "gridPosition": "start",
        "axisAlpha": 0,
        "gridAlpha": 0,
        "position": "left",   
        "labelRotation": 45
    },
    "export": {
    	"enabled": true
     }

});
</script>
	

<!------------------------------------------------------------------------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------->
<!-- END Chart Code For Unqualified Prospects By Lead Source and Reason -->
<!------------------------------------------------------------------------------->
<!------------------------------------------------------------------------------------------------------------------------------------------------->

