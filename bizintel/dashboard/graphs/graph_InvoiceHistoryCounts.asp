<%


Set cnnInvHistoryCountGraph = Server.CreateObject("ADODB.Connection")
cnnInvHistoryCountGraph.open (Session("ClientCnnString"))
Set rsInvHistoryCountGraph = Server.CreateObject("ADODB.Recordset")
rsInvHistoryCountGraph.CursorLocation = 3

%>
<div id="chartdivInvoiceHistoryCounts" style="width: 100%; height: 350px; margin: 0 auto"></div>	
<%

'Now Build the Chart Data Variable

amChartDataInvHistCounts = ""

currentMonthInYear = Month(Now())
currentYear = Year(Now())

If cInt(currentMonthInYear) = 12 Then
	
	For i = 1 to 12 'Loop through all months of this year
	
		loopMonth = i
		loopYear = currentYear

		%><!--#include file="graph_InvoiceHistoryCountsBuildData.asp"--><%
	Next
	
Else
	
	
	For i = cInt(currentMonthInYear) to 12 'Loop through the months in the previous year
	
		loopMonth = i
		loopYear = currentYear - 1
		
		%><!--#include file="graph_InvoiceHistoryCountsBuildData.asp"--><%		        
	Next

	For i = 1 to cInt(currentMonthInYear) 'Loop through all the remaining months of this year

		loopMonth = i
		loopYear = currentYear
		
		%><!--#include file="graph_InvoiceHistoryCountsBuildData.asp"--><% 
		        
	Next
	

End If

'Strip last comma

If amChartDataInvHistCounts <> "" Then
	amChartDataInvHistCounts = Left(amChartDataInvHistCounts, len(amChartDataInvHistCounts)-1)
End If


cnnInvHistoryCountGraph.Close
Set cnnInHistoryCountGraph = Nothing
Set rsInvHistoryCountGraph = Nothing
%>