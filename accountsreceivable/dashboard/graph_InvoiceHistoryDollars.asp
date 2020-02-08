<%


Set cnnInvHistoryDollarsGraph = Server.CreateObject("ADODB.Connection")
cnnInvHistoryDollarsGraph.open (Session("ClientCnnString"))
Set rsInvHistoryDollarsGraph = Server.CreateObject("ADODB.Recordset")
rsInvHistoryDollarsGraph.CursorLocation = 3


%>
<div id="chartdivInvoiceHistoryDollars" style="width: 100%; height: 400px; margin: 0 auto"></div>	
<%

'Now Build the Chart Data Variable

amChartDataInvHistDollars = ""

currentMonthInYear = Month(Now())
currentYear = Year(Now())

If cInt(currentMonthInYear) = 12 Then
	
	For i = 1 to 12 'Loop through all months of this year
	
		loopMonth = i
		loopYear = currentYear

		%><!--#include file="graph_InvoiceHistoryDollarsBuildData.asp"--><%
	Next
	
Else
	
	
	For i = cInt(currentMonthInYear) to 12 'Loop through the months in the previous year
	
		loopMonth = i
		loopYear = currentYear - 1
		
		%><!--#include file="graph_InvoiceHistoryDollarsBuildData.asp"--><%		        
	Next

	For i = 1 to cInt(currentMonthInYear) 'Loop through all the remaining months of this year

		loopMonth = i
		loopYear = currentYear
		
		%><!--#include file="graph_InvoiceHistoryDollarsBuildData.asp"--><% 
		        
	Next
	

End If

'Strip last comma

If amChartDataInvHistDollars <> "" Then
	amChartDataInvHistDollars = Left(amChartDataInvHistDollars, len(amChartDataInvHistDollars)-1)
End If


cnnInvHistoryDollarsGraph.Close
Set cnnInHistoryDollarsGraph = Nothing
Set rsInvHistoryDollarsGraph = Nothing

%>