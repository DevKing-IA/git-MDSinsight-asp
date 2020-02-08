<!--#include file="../inc/header-prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->

<%

function mmddyy(input)
    dim m: m = month(input)
    dim d: d = day(input)
    if (m < 10) then m = "0" & m
    if (d < 10) then d = "0" & d

    mmddyy = m & "/" & d & "/" & right(year(input), 2)
end function

Function dateCustomFormat(date)
	x = FormatDateTime(date, 2) 
	d = split(x, "/")
	dateCustomFormat = d(2) & "-" & d(0) & "-" & d(1)
End Function


	ReportSpecificData2 = ""
	ReportSpecificData2a = ""
	ReportSpecificData2b = ""
	ReportSpecificData3 = ""
	ReportSpecificData4 = ""
	ReportSpecificData5 = ""	
	ReportSpecificData6 = ""
	ReportSpecificData7 = ""
	ReportSpecificData8 = ""
	ReportSpecificData10 = ""
	ReportSpecificData10a = ""
	ReportSpecificData11 = ""
	ReportSpecificData12 = ""	
	ReportSpecificData13 = ""
	ReportSpecificData14 = ""
	ReportSpecificData15 = ""
	ReportSpecificData16 = ""
	ReportSpecificData17 = ""
	ReportSpecificData18 = ""
	ReportSpecificData19 = ""
	
'**************************************************************************************
'To create an "All Prospects" view, set the owner based on permission type, and set 
'Next Activity date range to all activities from 1/1/14 to today
'**************************************************************************************

	If userIsAdmin(Session("userNo")) = True OR userIsInsideSalesManager(Session("userNo")) = True OR userIsOutsideSalesManager(Session("userNo")) = True Then
		ReportSpecificData9 = ""
	Else
		ReportSpecificData9 = Session("userNo")
	End If
	
	'ReportSpecificData20 = "NextActivityScheduledDateRange"
	'ReportSpecificData21 = dateCustomFormat("1/1/2014")
	'ReportSpecificData22 = dateCustomFormat(date())
	
	ReportSpecificData20 = ""
	ReportSpecificData21 = ""
	ReportSpecificData22 = ""
	ReportSpecificData22a = ""
	
	dummy = MUV_WRITE("CRMSTARTDATE",ReportSpecificData21)
	dummy = MUV_WRITE("CRMENDDATE",ReportSpecificData22)
	
'***********************************************************************************************************************

	



SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = 'All Prospects'"

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)

'Rec does not exist yet, make it quick but empty, update it later
If rs.EOF Then
	SQL = "INSERT INTO Settings_Reports (ReportNumber, UserNo, PoolForProspecting, UserReportName) Values (1400, " & Session("userNo") & ",'Live','All Prospects')"
	rs.Close
	Set rs= cnn8.Execute(SQL)
End If


'Now update the table with the values
SQL = "Update Settings_Reports Set "
SQL = SQL & "ReportSpecificData2 = '" & ReportSpecificData2 & "', " 
SQL = SQL & "ReportSpecificData2a = '" & ReportSpecificData2a & "', "
SQL = SQL & "ReportSpecificData2b = '" & ReportSpecificData2b & "', "
SQL = SQL & "ReportSpecificData3 = '" & ReportSpecificData3 & "', " 
SQL = SQL & "ReportSpecificData4 = '" & ReportSpecificData4 & "', " 
SQL = SQL & "ReportSpecificData5 = '" & ReportSpecificData5 & "', " 
SQL = SQL & "ReportSpecificData6 = '" & ReportSpecificData6 & "', " 
SQL = SQL & "ReportSpecificData7 = '" & ReportSpecificData7 & "', " 
SQL = SQL & "ReportSpecificData8 = '" & ReportSpecificData8 & "', " 
SQL = SQL & "ReportSpecificData9 = '" & ReportSpecificData9 & "', " 
SQL = SQL & "ReportSpecificData10 = '" & ReportSpecificData10 & "', "
SQL = SQL & "ReportSpecificData10a = '" & ReportSpecificData10a & "', "
SQL = SQL & "ReportSpecificData11 = '" & ReportSpecificData11 & "', "
SQL = SQL & "ReportSpecificData12 = '" & ReportSpecificData12 & "', "
SQL = SQL & "ReportSpecificData13 = '" & ReportSpecificData13 & "', "
SQL = SQL & "ReportSpecificData14 = '" & ReportSpecificData14 & "', "
SQL = SQL & "ReportSpecificData15 = '" & ReportSpecificData15 & "', "
SQL = SQL & "ReportSpecificData16 = '" & ReportSpecificData16 & "', "
SQL = SQL & "ReportSpecificData17 = '" & ReportSpecificData17 & "', "
SQL = SQL & "ReportSpecificData18 = '" & ReportSpecificData18 & "', "
SQL = SQL & "ReportSpecificData19 = '" & ReportSpecificData19 & "', "
SQL = SQL & "ReportSpecificData20 = '" & ReportSpecificData20 & "', "
SQL = SQL & "ReportSpecificData21 = '" & ReportSpecificData21 & "', "
SQL = SQL & "ReportSpecificData22 = '" & ReportSpecificData22 & "', "
SQL = SQL & "ReportSpecificData22a = '" & ReportSpecificData22a & "' "
	 
SQL = SQL & " WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = 'All Prospects'"
Set rs= cnn8.Execute(SQL)

cnn8.Close

Set rs = Nothing
Set cnn8 = Nothing
dummy = MUV_WRITE("CRMVIEWSTATE","All Prospects")
%>

<form id="frmCreateAllProspectsView" name="frmCreateAllProspectsView" method="POST" action="main.asp">
	<input type="hidden" name="selectFilteredView" id="selectFilteredView" value="All Prospects">
</form>

<script type="text/javascript">
  document.forms['frmCreateAllProspectsView'].submit();
</script>	
	
<%
'***************************************************************************************************************
'***************************************************************************************************************
'                                                                                                              *
'                                                                                                              *
' REPORT FIELDS AND THE CUSTOM FILTER FIELDS THAT THEY REFER TO                                                *
'                                                                                                              *
'                                                                                                              *
' 1. ReportSpecificData1 = Columns Filter from other modal on prospecting screen                               *
'                                                                                                              *
' 2. ReportSpecificData2 = Filter Type for Stage - HasNotChangedInXDays, HasChangedInXDays,                    *
'                                                  HasNotChangedInDateRange, HasChangedInDateRange             *
'                                                                                                              *
'                                                                                                              *
'                                                                                                              *
' 2a. ReportSpecificData2a = Stage HAS NOT Changed Date Quick Pick Range                                       *
'                                                                                                              *
'										Value comes from selStageNotChangedDateRangeCustom                     *
'                                                                                                              *
'                                                                                                              *
'                                                                                                              *
'                                                                                                              *
' 2b. ReportSpecificData2b = Stage HAS Changed Date Quick Pick Range                                           *
'                                                                                                              *
'										Value comes from selStageChangedDateRangeCustom                        *
'                                                                                                              *
'                                                                                                              *
' 3. ReportSpecificData3 = # Days For HasNotChangedInXDays, HasChangedInXDays                                  *
'                                     Values Come From: txtStageNotChangedDays,                                *
'                                                       txtStageChangedDays                                    *
'                                                                                                              *
' 4. ReportSpecificData4 = Stage Date Range Start Date For HasNotChangedInDateRange, HasChangedInDateRange     *
'                                                                                                              *
'                                     Values Come From: txtStageNotChangedDateRangeStartDate,                  *
'                                                       txtStageChangedDateRangeStartDate                      *
'                                                                                                              *
' 5. ReportSpecificData5 = End Date Range Date For HasNotChangedInDateRange, HasChangedInDateRange             *
'                                                                                                              *
'                                     Values Come From: txtStageNotChangedDateRangeEndDate,                    *
'                                                       txtStageChangedDateRangeEndDate                        *
'                                                                                                              *
' 6. ReportSpecificData6 = Filter by Lead Source from field selLeadSourceNumber                                *
'                                                                                                              *
' 7. ReportSpecificData7 = Filter by Industry from field selIndustryNumber                                     *
'                                                                                                              *
' 8. ReportSpecificData8 = Filter by Telemarketer from field selTelemarketerUserNo                             *
'                                                                                                              *
' 9. ReportSpecificData9 = Filter by Prospect Owner from field selProspectOwnerUserNo                          *
'                                                                                                              *
'                                                                                                              *
'                                                                                                              *
' 10. ReportSpecificData10 = Filter by Prospect Creator from field selProspectCreatedByUserNo                  *
'                                                                                                              *
'                                                                                                              *
' 10a. ReportSpecificData10a = Prospect Creation Date Quick Pick Range                                         *
'                                                                                                              *
'										Value comes from selProspectCreatedDateRangeCustom                     *
'                                                                                                              *
'                                                                                                              *
' 11. ReportSpecificData11 = Created By Date Range Start Date For Prospect Creation Date                       *
'                                                                                                              *
'                                     Value Comes From: txtProspectCreatedRangeStartDate                       *
'                                                                                                              *
' 12. ReportSpecificData12 = Created By Range End Date For Prospect Creation Date                              *
'                                                                                                              *
'                                     Value Comes From: txtProspectCreatedRangeEndDate                         *
'                                                                                                              *
'                                                                                                              *
'                                                                                                              *
' 13. ReportSpecificData13 = Filter Type for Number of Employees, optNumEmployeesRangeCompare                  *
'                            Values: ByPredefinedRange, ByCustomNumber, ByCustomRange                          *
'                                                                                                              *
'  *. ReportSpecificData13 = Filter Value for Number of Employees                                              *
'                            Values ByPredefinedRange: selEmployeeRangeNo                                      *
'                            Values ByCustomNumber: selEmployeeRangeComparisonOperator,                        *
'                                                   txtEmployeeRangeComparisonNumberSingle                     *
'                            Values ByCustomRange : txtEmployeeCustomRangeNumber1,                             *
'                                                   txtEmployeeCustomRangeNumber2                              *
'                                                                                                              * 
'                                                                                                              *
' 14. ReportSpecificData14 = Filter Type for Number of Pantries, optNumPantriesCompare                         *
'                            Values: ByCustomNumber, ByCustomRange                                             *
'                                                                                                              *   
'  *. ReportSpecificData14 = Filter Value for Number of Pantries                                               *
'                            Values ByCustomNumber: selNumPantriesComparisonOperator,                          *
'                                                   txtNumPantriesComparisonNumberSingle                       *
'                            Values ByCustomRange : txtNumPantriesCustomRangeNumber1,                          *
'                                                   txtNumPantriesCustomRangeNumber2                           *
'                                                                                                              *
'                                                                                                              *
' 15. ReportSpecificData15 = City Autocomplete Filter                                                          *
'                                                                                                              *
' 16. ReportSpecificData16 = State Autocomplete Filter                                                         *
'                                                                                                              *
' 17. ReportSpecificData17 = Zip Code Autocomplete Filter                                                      *
'                                                                                                              *
' 18. ReportSpecificData18 = Selected Stages To Filter                                                         *
'                                                                                                              *
'                                                                                                              *
' 19. ReportSpecificData19 = Selected Next Activities To Filter                                                *
'                                                                                                              *
' 20. ReportSpecificData20 = Filter Selected for Next Activity - NextActivityScheduledDateRange                *
'                                                                                                              *
'                                                                                                              *
' 21. ReportSpecificData21 = Next Activity Date Range Start Date For NextActivityScheduledDateRange            *
'                                                                                                              *
'                                     Value Comes From: txtNextActivityScheduledDateRangeStartDate,            *
'                                                                                                              *
' 22. ReportSpecificData22 = Next Activity End Date Range Date For NextActivityScheduledDateRange              *
'                                                                                                              *
'                                     Value Comes From: txtNextActivityScheduledDateRangeStartDate,            *
'                                                                                                              *
' 22a. ReportSpecificData22a = Next Activity Quick Pick Range For NextActivityScheduledDateRange               *
'                                                                                                              *
'										Value comes from selNextActivityScheduledDateRangeCustom               *
'                                                                                                              *
' 23. ReportSpecificData23 = Selected LOST Stages To Filter                                                    *
'                                                                                                              *
' 23a. ReportSpecificData23a = Stage Quick Pick Range For WasLostInDateRange                                   *
'                                                                                                              *
'										Value comes from selLostStageDateRangeCustom                           *
'                                                                                                              *
'                                                                                                              *
' 24. ReportSpecificData24 = Stage Date Range Start Date For WasLostInDateRange                                *
'                                                                                                              *
'                                     Value Comes From: txtStageLostDateRangeStartDate                         *
'                                                                                                              *
' 25. ReportSpecificData25 = End Date Range Date For WasLostInDateRange                                        *
'                                                                                                              *
'                                     Value Comes From: txtStageLostDateRangeEndDate                           *
'                                                                                                              * 
'                                                                                                              *
'                                                                                                              *
'                                                                                                              *
'                                                                                                              *
' 26. ReportSpecificData26 = Selected UNQUALIFIED Stages To Filter                                             *
'                                                                                                              *
'                                                                                                              *
' 26a. ReportSpecificData26a = Stage Quick Pick Range For WasUnqualifiedInDateRange                            *
'                                                                                                              *
'										Value comes from selUnqualifiedStageDateRangeCustom                    *
'                                                                                                              *
'                                                                                                              *
' 27. ReportSpecificData27 = Stage Date Range Start Date For WasUnqualifiedInDateRange                         *
'                                                                                                              *
'                                     Value Comes From: txtStageUnqualifiedRangeStartDate                      *
'                                                                                                              *
' 28. ReportSpecificData28 = End Date Range Date For WasLostInDateRange                                        *
'                                                                                                              *
'                                     Value Comes From: txtStageUnqualifiedRangeEndDate                        *
'                                                                                                              * 
'***************************************************************************************************************
'***************************************************************************************************************


 %>