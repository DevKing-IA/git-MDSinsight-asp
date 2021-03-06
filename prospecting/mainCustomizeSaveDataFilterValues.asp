﻿<!--#include file="../inc/header-prospecting.asp"-->
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




'******************************************************************************************************************
'This will give us an array of the internal record identifiers of the selected stage checkboxes
'******************************************************************************************************************
selectedStagesToFilter = Request.Form("chkStage")

'******************************************************************************************************************
'We need to see if the user selected any stages
'******************************************************************************************************************
selectedStagesToFilterArray = Split(selectedStagesToFilter,",")
upperBound = ubound(selectedStagesToFilterArray)
stagesToFilter = "" 

For i=0 to upperBound
   'Response.write("The value of array in " & i & " is :"  & selectedStagesToFilterArray(i) & "<br />")
   If selectedStagesToFilterArray(i) <> "" Then
   		stagesToFilter = stagesToFilter & selectedStagesToFilterArray(i) & ","
   End If
Next

If stagesToFilter <> "" Then
	stagesToFilter = Left(stagesToFilter, Len(stagesToFilter)-1)
End If


'******************************************************************************************************************
'If the user did select stages to filter from, then we need to see if they selected any date filtering criteria
'******************************************************************************************************************

If stagesToFilter <> "" Then
	ReportSpecificData18 = stagesToFilter
Else
	ReportSpecificData18 = ""
End If

'*****************************************************
'Did the user select the filter, "Where stage HAS NOT changed in X days"'
'*****************************************************
optStageChangeDatesNotChangedDays = Request.Form("optStageChangeDatesNotChangedDays")

If optStageChangeDatesNotChangedDays = "HasNotChangedInXDays" Then

	txtStageNotChangedDays = Request.Form("txtStageNotChangedDays")
	ReportSpecificData2 = "HasNotChangedInXDays"
	ReportSpecificData3 = txtStageNotChangedDays
	ReportSpecificData4 = ""
	ReportSpecificData5 = ""	
	
End If
	
'*****************************************************
'Did the user select the filter, "Where stage HAS changed in X days"'
'*****************************************************
optStageChangeDatesChangedDays = Request.Form("optStageChangeDatesChangedDays")

If optStageChangeDatesChangedDays = "HasChangedInXDays" Then

	txtStageChangedDays = Request.Form("txtStageChangedDays")
	ReportSpecificData2 = "HasChangedInXDays"
	ReportSpecificData3 = txtStageChangedDays
	ReportSpecificData4 = ""
	ReportSpecificData5 = ""
	
End If

'*****************************************************
'Did the user select the filter, "Where stage HAS NOT changed in a DATE range"'
'*****************************************************
optStageNotChangedDateRange = Request.Form("optStageNotChangedDateRange")

If optStageNotChangedDateRange = "HasNotChangedInDateRange" Then

	selStageNotChangedDateRangeCustom = Request.Form("selStageNotChangedDateRangeCustom")
	txtStageNotChangedDateRangeStartDate = Request.Form("txtStageNotChangedDateRangeStartDate")
	txtStageNotChangedDateRangeEndDate= Request.Form("txtStageNotChangedDateRangeEndDate")
	
	If selStageNotChangedDateRangeCustom <> "" Then
		ReportSpecificData2 = "HasNotChangedInDateRange"
		ReportSpecificDat2a = selStageNotChangedDateRangeCustom
		ReportSpecificData3 = ""
		ReportSpecificData4 = "" 
		ReportSpecificData5 = ""
	Else
		ReportSpecificData2 = "HasNotChangedInDateRange"
		ReportSpecificDat2a = ""
		ReportSpecificData3 = ""
		ReportSpecificData4 = txtStageNotChangedDateRangeStartDate 
		ReportSpecificData5 = txtStageNotChangedDateRangeEndDate
	End If		
End If
	
'*****************************************************
'Did the user select the filter, "Where stage HAS changed in a DATE range"'
'*****************************************************

optStageChangeDateRange = Request.Form("optStageChangeDateRange")

If optStageChangeDateRange = "HasChangedInDateRange" Then

	selStageChangedDateRangeCustom = Request.Form("selStageChangedDateRangeCustom")
	txtStageChangedDateRangeStartDate = Request.Form("txtStageChangedDateRangeStartDate")
	txtStageChangedDateRangeEndDate = Request.Form("txtStageChangedDateRangeEndDate")
	
	If selStageChangedDateRangeCustom <> "" Then
		ReportSpecificData2 = "HasChangedInDateRange"
		ReportSpecificData2b = selStageChangedDateRangeCustom
		ReportSpecificData3 = ""
		ReportSpecificData4 = ""
		ReportSpecificData5 = ""
	Else
		ReportSpecificData2 = "HasChangedInDateRange"
		ReportSpecificData2b = ""
		ReportSpecificData3 = ""
		ReportSpecificData4 = txtStageChangedDateRangeStartDate
		ReportSpecificData5 = txtStageChangedDateRangeEndDate
	End If
End If



'******************************************************************************************************************
'Miscellaneous drop down fields for filtering criteria
'******************************************************************************************************************

	selLeadSourceNumber = Request.Form("selLeadSourceNumber")
	ReportSpecificData6 = selLeadSourceNumber
	
	selIndustryNumber = Request.Form("selIndustryNumber")
	ReportSpecificData7 = selIndustryNumber
	
	selTelemarketerUserNo = Request.Form("selTelemarketerUserNo")
	ReportSpecificData8 = selTelemarketerUserNo
	
	selProspectOwnerUserNo = Request.Form("selProspectOwnerUserNo")
	ReportSpecificData9 = selProspectOwnerUserNo
	
	selProspectCreatedByUserNo = Request.Form("selProspectCreatedByUserNo")
	ReportSpecificData10 = selProspectCreatedByUserNo


'******************************************************************************************************************
'Prospect Creation daterange picker filter
'******************************************************************************************************************

	'*****************************************************
	'Did the user select the filter, "Where prospect created within in a DATE range"'
	'*****************************************************
	
	optProspectCreatedDateRange = Request.Form("optProspectCreatedDateRange")

	If optProspectCreatedDateRange = "setrange" Then
	
		txtProspectCreatedRangeStartDate = Request.Form("txtProspectCreatedRangeStartDate")
		txtProspectCreatedRangeEndDate = Request.Form("txtProspectCreatedRangeEndDate")
		selProspectCreatedDateRangeCustom = Request.Form("selProspectCreatedDateRangeCustom")

		If selProspectCreatedDateRangeCustom <> "" Then
			ReportSpecificData10a = selProspectCreatedDateRangeCustom
			ReportSpecificData11 = ""
			ReportSpecificData12 = ""
		Else
			ReportSpecificData10a = ""
			ReportSpecificData11 = txtProspectCreatedRangeStartDate
			ReportSpecificData12 = txtProspectCreatedRangeEndDate		
		End If
	Else
		ReportSpecificData10a = ""
		ReportSpecificData11 = ""
		ReportSpecificData12 = ""	
	End If



'******************************************************************************************************************
'Number of Employees filter
'******************************************************************************************************************

	'*****************************************************
	'Did the user select the filter, "select predefined employee range"'
	'*****************************************************
	optNumEmployeesRangeCompare = Request.Form("optNumEmployeesRangeCompare")


	'******************************************************************************************************************
	'If the user did select number of employees to filter from, then we need to check which criteria they selected
	'******************************************************************************************************************
		
	If optNumEmployeesRangeCompare <> "" Then

		If optNumEmployeesRangeCompare = "ByPredefinedRange" Then
	
			selEmployeeRangeNo = Request.Form("selEmployeeRangeNo")
			ReportSpecificData13 = "ByPredefinedRange," & selEmployeeRangeNo & ",X"
			
		ElseIf optNumEmployeesRangeCompare = "ByCustomNumber" Then
		
			selEmployeeRangeComparisonOperator = Request.Form("selEmployeeRangeComparisonOperator")
			txtEmployeeRangeComparisonNumberSingle = Request.Form("txtEmployeeRangeComparisonNumberSingle")
			
			ReportSpecificData13 = "ByCustomNumber," & selEmployeeRangeComparisonOperator & "," & txtEmployeeRangeComparisonNumberSingle
			
		ElseIf optNumEmployeesRangeCompare = "ByCustomRange" Then 
		
			txtEmployeeCustomRangeNumber1 = Request.Form("txtEmployeeCustomRangeNumber1")
			txtEmployeeCustomRangeNumber2 = Request.Form("txtEmployeeCustomRangeNumber2")
			
			ReportSpecificData13 = "ByCustomRange," & txtEmployeeCustomRangeNumber1 & "," & txtEmployeeCustomRangeNumber2
		
		End If
		
	Else
	
		ReportSpecificData13 = ""
		
	End If


'******************************************************************************************************************
'Number of Pantries filter
'******************************************************************************************************************

	'*****************************************************
	'Did the user select the filter, "select predefined employee range"'
	'*****************************************************
	optNumPantriesCompare = Request.Form("optNumPantriesCompare")


	'******************************************************************************************************************
	'If the user did select number of employees to filter from, then we need to check which criteria they selected
	'******************************************************************************************************************
		
	If optNumPantriesCompare <> "" Then

		If optNumPantriesCompare = "ByCustomNumber" Then
		
			selNumPantriesComparisonOperator = Request.Form("selNumPantriesComparisonOperator")
			txtNumPantriesComparisonNumberSingle = Request.Form("txtNumPantriesComparisonNumberSingle")
			
			ReportSpecificData14 = "ByCustomNumber," & selNumPantriesComparisonOperator & "," & txtNumPantriesComparisonNumberSingle
			
		ElseIf optNumPantriesCompare= "ByCustomRange" Then 
		
			txtNumPantriesCustomRangeNumber1 = Request.Form("txtNumPantriesCustomRangeNumber1")
			txtNumPantriesCustomRangeNumber2 = Request.Form("txtNumPantriesCustomRangeNumber2")
			
			ReportSpecificData14 = "ByCustomRange," & txtNumPantriesCustomRangeNumber1 & "," & txtNumPantriesCustomRangeNumber2
		
		End If
	Else
		ReportSpecificData14 = ""		
	End If

'******************************************************************************************************************
'Location - City, State, Zip filter
'******************************************************************************************************************

	ReportSpecificData15 = Request.Form("txtCityFilter")
	ReportSpecificData16 = Request.Form("txtStateFilter")
	ReportSpecificData17 = Request.Form("txtZipFilter")

'******************************************************************************************************************


'******************************************************************************************************************
'This will give us an array of the internal record identifiers of the selected next activity checkboxes
'******************************************************************************************************************
selectedNextActivitiesToFilter = Request.Form("chkNextActivity")

'******************************************************************************************************************
'We need to see if the user selected any next activities
'******************************************************************************************************************
selectedNextActivitiesToFilterArray = Split(selectedNextActivitiesToFilter,",")
upperBoundNextActivity = ubound(selectedNextActivitiesToFilterArray)
nextActivitiesToFilter = "" 

For i=0 to upperBoundNextActivity
   If selectedNextActivitiesToFilterArray(i) <> "" Then
   		nextActivitiesToFilter = nextActivitiesToFilter & selectedNextActivitiesToFilterArray(i) & ","
   End If
Next

If nextActivitiesToFilter <> "" Then
	nextActivitiesToFilter = Left(nextActivitiesToFilter, Len(nextActivitiesToFilter)-1)
End If


'***********************************************************************************************************************
'If the user did select next activities to filter from, then we need to see if they selected any date filtering criteria
'***********************************************************************************************************************

If nextActivitiesToFilter <> "" Then
	ReportSpecificData19 = nextActivitiesToFilter
Else
	ReportSpecificData19 = ""
End If

	
'**************************************************************************************
'Did the user select the filter, "Where a next activity is schedule in a date range"'
'**************************************************************************************

optNextActivityDateRange = Request.Form("optNextActivityScheduledDateRange")

If optNextActivityDateRange = "NextActivityScheduledDateRange" Then

	txtNextActivityScheduledDateRangeStartDate = Request.Form("txtNextActivityScheduledDateRangeStartDate")
	txtNextActivityScheduledDateRangeEndDate = Request.Form("txtNextActivityScheduledDateRangeEndDate")
	selNextActivityScheduledDateRangeCustom = Request.Form("selNextActivityScheduledDateRangeCustom")
	
	If selNextActivityScheduledDateRangeCustom <> "" Then
		ReportSpecificData20 = "NextActivityScheduledDateRange"
		ReportSpecificData21 = ""
		ReportSpecificData22 = ""
		ReportSpecificData22a = selNextActivityScheduledDateRangeCustom
	Else
		ReportSpecificData20 = "NextActivityScheduledDateRange"
		ReportSpecificData21 = txtNextActivityScheduledDateRangeStartDate
		ReportSpecificData22 = txtNextActivityScheduledDateRangeEndDate
		ReportSpecificData22a = ""
		dummy = MUV_WRITE("CRMSTARTDATE",ReportSpecificData21)
		dummy = MUV_WRITE("CRMENDDATE",ReportSpecificData22)
	End If
	
Else
	ReportSpecificData20 = ""
	ReportSpecificData21 = ""
	ReportSpecificData22 = ""
	ReportSpecificData22a = ""
End If
'***********************************************************************************************************************



SQL = "SELECT * from Settings_Reports where ReportNumber = 1400 AND PoolForProspecting = 'Live' AND UserNo = " & Session("userNo")

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)

'Rec does not exist yet, make it quick but empty, update it later
If rs.EOF Then
	SQL = "Insert into Settings_Reports (ReportNumber, UserNo, PoolForProspecting) Values (1400, " & Session("userNo") & ", 'Live')"
	rs.Close
	Set rs= cnn8.Execute(SQL)
End If


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
'                                     Values Come From: txtNextActivityScheduledDateRangeStartDate,            *
'                                                       txtNextActivityScheduledDateRangeEndDate               *
'                                                                                                              *
' 22. ReportSpecificData22 = Next Activity End Date Range Date For NextActivityScheduledDateRange              *
'                                                                                                              *
'                                     Values Come From: txtNextActivityScheduledDateRangeStartDate,            *
'                                                       txtNextActivityScheduledDateRangeEndDate               *
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

'Response.Write("ReportSpecificData2: " & ReportSpecificData2 & "<br>")


'**********************************************************************************************************************************
'If all filters have been cleared, delete the user record in Settings_Reports and recreate the default user view record
'**********************************************************************************************************************************
If ReportSpecificData2 = "" AND ReportSpecificData2a = "" AND ReportSpecificData2b = "" AND ReportSpecificData3 = "" AND ReportSpecificData4 = "" AND _
	ReportSpecificData5 = "" AND ReportSpecificData6 = "" AND ReportSpecificData7 = "" AND ReportSpecificData8 = "" AND ReportSpecificData9 = "" AND _
	ReportSpecificData10 = "" AND ReportSpecificData10a = "" AND ReportSpecificData11 = "" AND ReportSpecificData12 = "" AND ReportSpecificData13 = "" AND _
	ReportSpecificData14 = "" AND ReportSpecificData15 = "" AND ReportSpecificData16 = "" AND ReportSpecificData17 = "" AND _
	ReportSpecificData18 = "" AND ReportSpecificData19 = "" AND ReportSpecificData20 = "" AND ReportSpecificData21 = "" AND _
	ReportSpecificData22 = "" AND ReportSpecificData22a = "" Then	
	
	dummy = MUV_WRITE("CRMVIEWSTATE","Default")
	%>
	<form id="frmClearFilterView" name="frmClearFilterView" method="POST" action="main.asp">
		<input type="hidden" name="selectFilteredView" id="selectFilteredView" value="Default">
	</form>
	
	<script type="text/javascript">
	  document.forms['frmClearFilterView'].submit();
	</script>	
<%		

Else

	SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = 'Current'"
	
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs= cnn8.Execute(SQL)
	
	'Rec does not exist yet, make it quick but empty, update it later
	If rs.EOF Then
		SQL = "INSERT INTO Settings_Reports (ReportNumber, UserNo, PoolForProspecting, UserReportName) Values (1400, " & Session("userNo") & ",'Live','Current')"
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
	 
	SQL = SQL & " WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = 'Current'"
	Set rs= cnn8.Execute(SQL)
	
	cnn8.Close
	
	Set rs = Nothing
	Set cnn8 = Nothing
	dummy = MUV_WRITE("CRMVIEWSTATE","Current")
	%>
	
	<form id="frmClearFilterView2" name="frmClearFilterView2" method="POST" action="main.asp">
		<input type="hidden" name="selectFilteredView" id="selectFilteredView" value="Current">
	</form>
	
	<script type="text/javascript">
	  document.forms['frmClearFilterView2'].submit();
	</script>	
	
	<%

End If
%>

 
