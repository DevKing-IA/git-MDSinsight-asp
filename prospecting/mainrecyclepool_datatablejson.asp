<% If Session("Userno") = "" Then Response.Redirect("../default.asp") %>

<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/protect.asp"-->
<!--#include file="../inc/InsightFuncs.asp"-->

<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<% Server.ScriptTimeout = 9000 %>

<%

Response.AddHeader "Content-Type", "application/json"

'Quick rebuild of the PR_ProspectContactSearch table

Set DataConn = Server.CreateObject("ADODB.Connection")
DataConn.CursorLocation = adUseClient
DataConn.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")

SQL = "DELETE FROM PR_ProspectContactSearch"
Set rs= DataConn.Execute(SQL)

SQL = "INSERT INTO PR_ProspectContactSearch (ProspectIntRecID, Company, City, State, FirstName, LastName) "
SQL = SQL & "SELECT PR_Prospects.InternalRecordIdentifier, PR_Prospects.Company, PR_Prospects.City, PR_Prospects.State, "
SQL = SQL & "PR_ProspectContacts.FirstName, PR_ProspectContacts.LastName "
SQL = SQL & "FROM PR_Prospects LEFT OUTER JOIN "
SQL = SQL & "PR_ProspectContacts ON PR_ProspectContacts.ProspectIntRecID = PR_Prospects.InternalRecordIdentifier "
SQL = SQL & "WHERE PR_Prospects.Pool = 'Dead'"
Set rs= DataConn.Execute(SQL)

Set rs = Nothing
'cnn8.Close
'Set cnn8 = Nothing


' show or hide autocomplete searchbox & get # days to Show

ShowLivePoolProspectSearchBox = False

'Set cnn9 = Server.CreateObject("ADODB.Connection")
'cnn9.open (Session("ClientCnnString"))

Set rs9 = DataConn.Execute("SELECT * FROM Settings_Prospecting")
If Not rs9.EOF Then
	If rs9("ShowLivePoolProspectSearchBox")=1 Then
		ShowLivePoolProspectSearchBox = True
	End If
	If IsNumeric(rs9("ProspectActivityDefaultDaysToShow")) Then
		ProspectActivityDefaultDaysToShow = rs9("ProspectActivityDefaultDaysToShow")
	Else
		ProspectActivityDefaultDaysToShow = 10
	End If
Else
	ProspectActivityDefaultDaysToShow = 10
End If
rs9.Close		
set rs9 = Nothing

'cnn9.Close
'Set cnn9 = Nothing

'****************************************************************************************************
'Read Settings_Reports To See If We Are Loading A Saved Custom Report
'****************************************************************************************************

customFilterReportName = Request.Form("selectFilteredView")
customFilterReportNameQuotes = Replace(Request.Form("selectFilteredView"),"''","'")

If customFilterReportName = "" Then 
	customFilterReportName = MUV_READ("CRMVIEWSTATERECPOOL")
Else
	dummy = MUV_WRITE("CRMVIEWSTATERECPOOL",customFilterReportNameQuotes)
End If

If customFilterReportName = "" Then 
	customFilterReportName = "Default"
	dummy = MUV_WRITE("CRMVIEWSTATERECPOOL","Default")
End If

customFilterReportNameForSQL = Replace(customFilterReportName,"'","''")


If MUV_READ("CRMVIEWSTATERECPOOL") = "Default" Then

	dateTenDaysFromNow = DateAdd("d",10, Now())
	
	'Now we have a setting for this, so possibly oover-ride the 10 days
	If IsNumeric(ProspectActivityDefaultDaysToShow) Then
		dateTenDaysFromNow = DateAdd("d",ProspectActivityDefaultDaysToShow, Now())			
	End If
	
	nextActivityStartDate = dateCustomFormat("01/01/2014")
	nextActivityEndDate = dateCustomFormat(dateTenDaysFromNow)

	SQL = "SELECT * from Settings_Reports where ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Dead' AND UserReportName = '" & customFilterReportNameForSQL & "'"
	'Set cnn8 = Server.CreateObject("ADODB.Connection")
	'cnn8.CursorLocation = adUseClient
	'cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs= DataConn.Execute(SQL)
	
	If Not rs.EOF Then
		SQL = "UPDATE Settings_Reports SET  ReportSpecificData21 = '" & nextActivityStartDate & "' "
		SQL = SQL & ", ReportSpecificData22 = '" & nextActivityEndDate & "' "
		SQL = SQL & "WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Dead' AND UserReportName = 'Default'"
		'Response.Write("<br><br><br><br>" & SQL)
		Set rs= DataConn.Execute(SQL)
	End If
	
	Set rs = Nothing
	'cnn8.Close
	'Set cnn8 = Nothing

End If


'****************************************************************************************************
'Read Settings_Reports To Obtain Filters For Prospecting Grid Data
'****************************************************************************************************
SQL = "SELECT * from Settings_Reports where ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Dead' AND UserReportName = '" & customFilterReportNameForSQL & "'"
'Set cnn8 = Server.CreateObject("ADODB.Connection")
'cnn8.CursorLocation = adUseClient
'cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= DataConn.Execute(SQL)

UseSettings_Reports = False
showHideColumns = ""


If NOT rs.EOF Then
	UseSettings_Reports = True
	showHideColumns 	 = rs("ReportSpecificData1")
	ReportSpecificData2  = rs("ReportSpecificData2")
	ReportSpecificData3  = rs("ReportSpecificData3")
	ReportSpecificData4  = rs("ReportSpecificData4")
	ReportSpecificData5  = rs("ReportSpecificData5")
	ReportSpecificData6  = rs("ReportSpecificData6")
	ReportSpecificData7  = rs("ReportSpecificData7")
	ReportSpecificData8  = rs("ReportSpecificData8")
	ReportSpecificData9  = rs("ReportSpecificData9")
	ReportSpecificData10  = rs("ReportSpecificData10")
	ReportSpecificData11  = rs("ReportSpecificData11")
	ReportSpecificData12  = rs("ReportSpecificData12")
	ReportSpecificData13  = rs("ReportSpecificData13")
	ReportSpecificData14  = rs("ReportSpecificData14")
	ReportSpecificData15  = rs("ReportSpecificData15")
	ReportSpecificData16  = rs("ReportSpecificData16")
	ReportSpecificData17  = rs("ReportSpecificData17")
	ReportSpecificData18  = rs("ReportSpecificData18")
	ReportSpecificData19  = rs("ReportSpecificData19")
	ReportSpecificData20  = rs("ReportSpecificData20")
	ReportSpecificData21  = rs("ReportSpecificData21")
	ReportSpecificData22  = rs("ReportSpecificData22")
	
		
Else
    '****************************************************************
	'Create Default Filter View User Record in Settings_Reports
	'****************************************************************
	
	'dateTenDaysFromNow = DateAdd("d",10, DateSerial(Year(Now()), Month(Now()), Day(Now())))
	dateTenDaysFromNow = DateAdd("d",10, Now())
	
	'Now we have a setting for this, so possibly oover-ride the 10 days
	If IsNumeric(ProspectActivityDefaultDaysToShow) Then
		'dateTenDaysFromNow = DateAdd("d",ProspectActivityDefaultDaysToShow, DateSerial(Year(Now()), Month(Now()), Day(Now())))	
		dateTenDaysFromNow = DateAdd("d",ProspectActivityDefaultDaysToShow, Now())			
	End If
	
	nextActivityStartDate = dateCustomFormat("01/01/2014")
	nextActivityEndDate = dateCustomFormat(dateTenDaysFromNow)
	
  	SQLCreateUserFilter = "INSERT INTO Settings_Reports(ReportNumber,UserNo,PoolForProspecting,ReportSpecificData9,ReportSpecificData20,ReportSpecificData21,ReportSpecificData22,UserReportName) "
  	SQLCreateUserFilter = SQLCreateUserFilter & " VALUES (1400," & Session("userNo") & ",'Dead'," & Session("userNo") & ",'NextActivityScheduledDateRange','" & nextActivityStartDate & "','" & nextActivityEndDate & "','" & customFilterReportNameForSQL & "')"

	'Set cnnCreateUserFilter = Server.CreateObject("ADODB.Connection")
	'cnnCreateUserFilter.CursorLocation = adUseClient
	'cnnCreateUserFilter.open (Session("ClientCnnString"))
	Set rsCreateUserFilter = Server.CreateObject("ADODB.Recordset")
	rsCreateUserFilter.CursorLocation = 3 
	
	Set rsCreateUserFilter = DataConn.Execute(SQLCreateUserFilter)
		
	set rsCreateUserFilter = Nothing
	'cnnCreateUserFilter.close
	'set cnnCreateUserFilter = Nothing
	
	ReportSpecificData9 = Session("userNo")
	ReportSpecificData20 = "NextActivityScheduledDateRange"
	ReportSpecificData21 = nextActivityStartDate
	ReportSpecificData22 = nextActivityEndDate
	dummy = MUV_WRITE("CRMSTARTDATE",ReportSpecificData21)
	dummy = MUV_WRITE("CRMENDDATE",ReportSpecificData22)
	
End If
'****************************************************************************************************
'End Read Settings_Reports
'****************************************************************************************************


'****************************************************************************************************
'Build Master SQL STMT With All Filters Set
'****************************************************************************************************

If ReportSpecificData18 <> "" Then
	selectedStagesToFilterArray = Split(ReportSpecificData18,",")
	upperBound = ubound(selectedStagesToFilterArray)
Else
	upperBound = -1
	selectedStagesToFilterArray = ""
End If


If ReportSpecificData2 <> "" Then

	If ReportSpecificData2 = "HasNotChangedInXDays" Then
	
		stageHasNotChangedInXDays = ReportSpecificData3
	
	ElseIf ReportSpecificData2 = "HasChangedInXDays" Then
	
		stageHasChangedInXDays = ReportSpecificData3
	
	ElseIf ReportSpecificData2 = "HasNotChangedInDateRange" Then
	
		If ReportSpecificData4 <> "" AND ReportSpecificData5 <> "" Then
			startDateStageNotChangedRange = dateCustomFormat(ReportSpecificData4)
			endDateStageNotChangedRange = dateCustomFormat(ReportSpecificData5) 
		End If
		
	ElseIf ReportSpecificData2 = "HasChangedInDateRange" Then
	
		If ReportSpecificData4 <> "" AND ReportSpecificData5 <> "" Then
			startDateStageChangedRange = dateCustomFormat(ReportSpecificData4)
			endDateStageChangedRange = dateCustomFormat(ReportSpecificData5) 
		End If
					
	End If
End If


If ReportSpecificData6 <> "" Then
	selectedLeadSourcesToFilterArray = Split(ReportSpecificData6,",")
	upperBoundLeadSource = ubound(selectedLeadSourcesToFilterArray)
Else
	upperBoundLeadSource = -1
	selectedLeadSourcesToFilterArray = ""
End If
For i=0 to upperBoundLeadSource
   '''''cInt(selectedLeadSourcesToFilterArray(i)) will be each selected lead source
Next
  		

If ReportSpecificData7 <> "" Then
	selectedIndustriesToFilterArray = Split(ReportSpecificData7,",")
	upperBoundIndustries = ubound(selectedIndustriesToFilterArray)
Else
	upperBoundIndustries = -1
	selectedIndustriesToFilterArray = ""
End If
For i=0 to upperBoundIndustries
   '''''cInt(selectedIndustriesToFilterArray(i)) will be each selected industry
Next



If ReportSpecificData8 <> "" Then
	selectedTelemarketersToFilterArray = Split(ReportSpecificData8,",")
	upperBoundTelemarketers = ubound(selectedTelemarketersToFilterArray )
Else
	upperBoundTelemarketers = -1
	selectedIndustriesToFilterArray = ""
End If
For i=0 to upperBoundTelemarketers
   '''''cInt(selectedTelemarketersToFilterArray(i)) will be each selected telemarketer
Next



If ReportSpecificData9 <> "" Then
	selectedOwnersToFilterArray = Split(ReportSpecificData9,",")
	upperBoundOwners = ubound(selectedOwnersToFilterArray)
Else
	upperBoundOwners = -1
	selectedOwnersToFilterArray = ""
End If
For i=0 to upperBoundOwners
   '''''cInt(selectedOwnersToFilterArray(i)) will be each selected owner
Next



If ReportSpecificData10 <> "" Then
	selectedCreatedByUsersToFilterArray = Split(ReportSpecificData10,",")
	upperBoundCreatedByUsers = ubound(selectedCreatedByUsersToFilterArray)
Else
	upperBoundCreatedByUsers = -1
	selectedCreatedByUsersToFilterArray = ""
End If
For i=0 to upperBoundCreatedByUsers
   '''''cInt(selectedCreatedByUsersToFilterArray(i)) will be each selected created by userno
Next



If ReportSpecificData11 <> "" AND ReportSpecificData12 <> "" Then
	startDateProspectCreatedRange = dateCustomFormat(ReportSpecificData11)
	endDateProspectCreatedRange = dateCustomFormat(ReportSpecificData12) 
End If



If ReportSpecificData13 <> "" Then
	employeeFilterArray = Split(ReportSpecificData13,",")
	uBoundEmployees = uBound(employeeFilterArray)
Else
	employeeFilterArray = ""
	uBoundEmployees = -1
End If

If uBoundEmployees > 0 Then

	selectedEmployeeFilterType = employeeFilterArray(0)
	
	If selectedEmployeeFilterType = "ByPredefinedRange" Then
		selectedEmployeeRange = employeeFilterArray(1)
		selectedEmployeeCompOperator = ""
		selectedEmployeeCompNumber = ""
		selectedEmployeeRangeNumber1 = ""
		selectedEmployeeRangeNumber2 = ""
	End If
	
	If selectedEmployeeFilterType = "ByCustomNumber" Then
		selectedEmployeeCompOperator = employeeFilterArray(1)
		selectedEmployeeCompNumber = employeeFilterArray(2)
		selectedEmployeeRange = ""
		selectedEmployeeRangeNumber1 = ""
		selectedEmployeeRangeNumber2 = ""				
	End If
	
	If selectedEmployeeFilterType = "ByCustomRange" Then
		selectedEmployeeRangeNumber1 = employeeFilterArray(1)
		selectedEmployeeRangeNumber2 = employeeFilterArray(2)
		selectedEmployeeRange = ""
		selectedEmployeeCompOperator = ""
		selectedEmployeeCompNumber = ""				
	End If
Else
	selectedEmployeeFilterType = ""
	selectedEmployeeRangeNumber1 = ""
	selectedEmployeeRangeNumber2 = ""
	selectedEmployeeRange = ""
	selectedEmployeeCompOperator = ""
	selectedEmployeeCompNumber = ""			     	
End If


If ReportSpecificData14 <> "" Then
	PantryFilterArray = Split(ReportSpecificData14,",")
	uBoundPantries = uBound(PantryFilterArray)
Else
	PantryFilterArray = ""
	uBoundPantries = -1
End If

If uBoundPantries > 0 Then

	selectedPantryFilterType = PantryFilterArray(0)
	     		
	If selectedPantryFilterType = "ByCustomNumber" Then
		selectedPantryCompOperator = PantryFilterArray(1)
		selectedPantryCompNumber = PantryFilterArray(2)
		selectedPantryRangeNumber1 = ""
		selectedPantryRangeNumber2 = ""				
	End If
	
	If selectedPantryFilterType = "ByCustomRange" Then
		selectedPantryRangeNumber1 = PantryFilterArray(1)
		selectedPantryRangeNumber2 = PantryFilterArray(2)
		selectedPantryCompOperator = ""
		selectedPantryCompNumber = ""				
	End If
Else
	selectedPantryFilterType = ""
	selectedPantryRangeNumber1 = ""
	selectedPantryRangeNumber2 = ""
	selectedPantryCompOperator = ""
	selectedPantryCompNumber = ""		     	
End If

'****************************************************************************************************
'End Build SQL Criteria
'****************************************************************************************************
'****************************************************************************************************
'End Read Settings_Reports
'****************************************************************************************************


%>


<%

	MaxActivityDaysWarningInit = GetCRMMaxActivityDaysWarning()
	MaxActivityDaysPermittedInit = GetCRMMaxActivityDaysPermitted()
	

%>



<%

function mmddyy(input)
    dim m: m = month(input)
    dim d: d = day(input)
    if (m < 10) then m = "0" & m
    if (d < 10) then d = "0" & d

    mmddyy = m & "/" & d & "/" & right(year(input), 2)
end function

function mmddyyyy(input)
    dim m: m = month(input)
    dim d: d = day(input)
    if (m < 10) then m = "0" & m
    if (d < 10) then d = "0" & d

    mmddyyyy = m & "/" & d & "/" & year(input)
end function




Function dateCustomFormat(passeDdate)
	x = FormatDateTime(passeDdate, 2) 
	d = split(x, "/")
	dateCustomFormat = d(2) & "-" & d(0) & "-" & d(1)
End Function



%>	


		
		<!--#include file="mainRecyclePoolCustomizeBuildSQLTable.asp"-->


		<%

	searchValue=Request.QueryString("search[value]")
	orderValue=Request.QueryString("order[0][column]")
	orderType=Request.QueryString("order[0][dir]")
	
	'orderValue =5
	
	If Request.QueryString("length")="" THEN
		PageSize=10
	ELSEIF Request.QueryString("length")="-1" THEN
		PageSize=0
	ELSE
		PageSize=CLng(Request.QueryString("length"))
	END IF
	
	IF Request.QueryString("start")="NaN" THEN
		rowStart=1
	ELSE
		rowStart=CLng(Request.QueryString("start"))
	END IF
	
	IF PageSize=0 THEN
		nPage=1
	ELSE
		nPage=1+rowStart/PageSize			
	END IF
		
	
		
		'SQL8 = "SELECT *, PR_Prospects.InternalRecordIdentifier AS Expr1  FROM PR_Prospects "
		
		SQL8 = "SELECT DISTINCT PR_Prospects.InternalRecordIdentifier AS Expr1,"
		SQL8 = SQL8 & "PR_Prospects.Company,PR_Prospects.Street,PR_Prospects.City,PR_Prospects.State,PR_Prospects.PostalCode,PR_Prospects.Country"
		SQL8 = SQL8 & ",PR_Prospects.LeadSourceNumber,PR_Prospects.IndustryNumber,PR_Prospects.EmployeeRangeNumber,PR_Prospects.OwnerUserNo,PR_Prospects.CreatedDate,PR_Prospects.CreatedByUserNo,PR_Prospects.TelemarketerUserNo,PR_Prospects.NumberOfPantries"
		SQL8 = SQL8 & ",PR_LeadSources.LeadSource"
		'SQL8 = SQL8 & ",PR_ProspectStages.RecordCreationDateTime"		
		SQL8 = SQL8 & ",PR_Industries.Industry"
		'SQL8 = SQL8 & ",PR_Stages.Stage"
		SQL8 = SQL8 & "  FROM PR_Prospects "
		
		SQL8 = SQL8 & " Inner Join zProspectFilter_" & Session("UserNo") & " ON PR_Prospects.InternalRecordIdentifier = "
		SQL8 = SQL8 & " zProspectFilter_" & Session("UserNo") & ".InternalRecordIdentifier"
		
		SQL8 = SQL8 & " INNER JOIN PR_ProspectStages ON PR_Prospects.InternalRecordIdentifier=PR_ProspectStages.ProspectRecID"
		SQL8 = SQL8 & " INNER JOIN PR_Stages ON PR_ProspectStages.StageRecID = PR_Stages.InternalRecordIdentifier"
		SQL8 = SQL8 & " INNER JOIN PR_Industries ON PR_Prospects.IndustryNumber=PR_Industries.InternalRecordIdentifier"
		SQL8 = SQL8 & " INNER JOIN PR_LeadSources ON PR_Prospects.LeadSourceNumber=PR_LeadSources.InternalRecordIdentifier"
		
	SQL8 = SQL8 & " WHERE PR_Prospects.Pool='Dead'"
	
	IF LEN(searchValue)>0 THEN
		SQL8=SQL8 & " AND (Company LIKE '%" & searchValue & "%')"
	END IF
	
	
	
		
		
	
	SELECT CASE orderValue
		CASE "0"
			SQL8=SQL8 & " ORDER BY Company " & orderType
		CASE "1"
			SQL8=SQL8 & " ORDER BY Company " & orderType
		CASE "2"
			SQL8=SQL8 & " ORDER BY PR_Prospects.Street  " & orderType
		CASE "3"
			SQL8=SQL8 & " ORDER BY PR_Prospects.City  " & orderType
		CASE "4"
			SQL8=SQL8 & " ORDER BY PR_Prospects.State  " & orderType
		CASE "5"
			SQL8=SQL8 & " ORDER BY PR_Prospects.PostalCode  " & orderType
		CASE "6"
			SQL8=SQL8 & " ORDER BY PR_LeadSources.LeadSource  " & orderType
		CASE "7"			
			'SQL8=SQL8 & " ORDER BY PR_Stages.Stage " & orderType
		CASE "8"			
			'SQL8=SQL8 & " ORDER BY PR_ProspectStages.RecordCreationDateTime " & orderType	
		CASE "9"			
			'SQL8=SQL8 & " ORDER BY PR_ProspectStages.RecordCreationDateTime " & orderType
		CASE "10"			
			SQL8=SQL8 & " ORDER BY PR_Industries.Industry " & orderType
		CASE "11"
			SQL8=SQL8 & " ORDER BY PR_Prospects.EmployeeRangeNumber " & orderType
		CASE "12"
			SQL8=SQL8 & " ORDER BY PR_Prospects.OwnerUserNo " & orderType	
		CASE "13"			
			SQL8=SQL8 & " ORDER BY PR_Prospects.CreatedDate " & orderType
		CASE "14"			
			SQL8=SQL8 & " ORDER BY PR_Prospects.CreatedByUserNo " & orderType
		CASE "15"			
			SQL8=SQL8 & " ORDER BY PR_Prospects.TelemarketerUserNo " & orderType
		CASE "16"			
			SQL8=SQL8 & " ORDER BY PR_Prospects.NumberOfPantries " & orderType
		CASE "17"			
			SQL8=SQL8 & " ORDER BY Expr1 " & orderType				
		CASE ELSE
			
			SQL8=SQL8 & " ORDER BY PR_Prospects.Company " & orderType
	
			
	END SELECT		

		
		
		'Set cnn8 = Server.CreateObject("ADODB.Connection")
		'cnn8.CursorLocation = adUseClient ' Wierd but must do this
		'cnn8.open (Session("ClientCnnString"))
		Set rs8 = Server.CreateObject("ADODB.Recordset")
		rs8.CursorLocation = 3'adUseClient
	


		'Set rs8 = DataConn.Execute(SQL8)
		rs8.Open SQL8, DataConn, 1
		
		IF PageSize=0 THEN
			rs8.PageSize=rs8.recordCount			
		ELSE
			rs8.PageSize = PageSize
		END IF
		nPageCount = rs8.PageCount
		nRecordCount=rs8.recordCount
			
			
		If not rs8.EOF Then
			rs8.AbsolutePage = nPage


			Do While Not ( rs8.Eof Or rs8.AbsolutePage <> nPage )
			
				'RecCounter = RecCounter + 1
			
				InternalRecordIdentifier = rs8("Expr1")
				Company = rs8("Company")
				'ActivityRecID = rs8("ActivityRecID")
				'ActivityDueDate = rs8("ActivityDueDate")
				Street = rs8("Street")
				City = rs8("City")
				State = rs8("State")
				PostalCode = rs8("PostalCode")
				Country = rs8("Country")
				LeadSourceNumber = rs8("LeadSourceNumber")
				IndustryNumber= rs8("IndustryNumber")
				EmployeeRangeNumber = rs8("EmployeeRangeNumber")
				OwnerUserNo = rs8("OwnerUserNo")
				CreatedDate = rs8("CreatedDate")
				CreatedByUserNo = rs8("CreatedByUserNo")
				TelemarketerUserNo = rs8("TelemarketerUserNo")
				NumberOfPantries = rs8("NumberOfPantries")
				'ProspectWatch = rs8("ProspectWatch")
				
				ActivityRecID = GetCurrentProspectActivityNumberByProspectNumber(InternalRecordIdentifier)
				ActivityDueDate = GetCurrentProspectActivityDueDateByProspectNumber(InternalRecordIdentifier)
				
 				tmpCreatedDate = CreatedDate
				tmpCreatedDate = cdate(tmpCreatedDate )
				
				tmpStageDate = GetProspectLastStageChangeDateByProspectNumber(InternalRecordIdentifier)
				tmpStageDate = cdate(tmpStageDate)
				
				
				col_nextactivity = "No Next Activity"
				col_nextactivityduedate = "No Next Activity"
				
			    If ActivityRecID <> "" Then 
			    
				    If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then
							col_nextactivity = "<a class='getProspectInfo' href='#'>"&GetActivityByNum(ActivityRecID)&"</a>"  
					ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then
							col_nextactivity = "<a class='getProspectInfo' data-toggle='modal' data-show='true' href='#' data-activity-id='"&ActivityRecID&"' data-prospect-id='"&InternalRecordIdentifier&"' data-target='#myProspectingModalEditActivity' data-tooltip='true' data-title='Edit Prospect Activity'>"&GetActivityByNum(ActivityRecID) &"&nbsp;&nbsp;<i class='fa fa-pencil-square fa-lg' aria-hidden='true'></i></a>" 
					ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then 
							col_nextactivity = "<a class='getProspectInfo' href='#'>"&GetActivityByNum(ActivityRecID)&"</a>"  
					ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then 
							col_nextactivity = "<a class='getProspectInfo' data-toggle='modal' data-show='true' href='#' data-activity-id='"&ActivityRecID&"' data-prospect-id='"& InternalRecordIdentifier &"' data-target='#myProspectingModalEditActivity' data-tooltip='true' data-title='Edit Prospect Activity'>"&GetActivityByNum(ActivityRecID)&"&nbsp;&nbsp;<i class='fa fa-pencil-square fa-lg' aria-hidden='true'></i></a>" 
					End If
					
					
						unformattedActivityTime = timevalue(hour(ActivityDueDate) & ":" & minute(ActivityDueDate))
						
						If hour(ActivityDueDate) > 12 Then
							activityTime = hour(ActivityDueDate) - 12  & ":" & minute(ActivityDueDate) & " " & right(unformattedActivityTime, 2)
						Else
							activityTime = hour(ActivityDueDate) & ":" & minute(ActivityDueDate) & " " & right(unformattedActivityTime, 2)
						End If
						
						
					If DateDiff("d",ActivityDueDate,Date()) > 0 Then 
						col_nextactivityduedate = "<span class='activitydateoverdue'>"&Month(ActivityDueDate) & "/" &  Day(ActivityDueDate) & "/" &  Right(Year(ActivityDueDate),2)&"</span><span class='activitytime'>"&activityTime&"</span>"
					ElseIf Abs(DateDiff("d",ActivityDueDate,Date())) = 0 Then 
						col_nextactivityduedate = "<span class='activitydatetoday'>"& Month(ActivityDueDate) & "/" &  Day(ActivityDueDate) & "/" &  Right(Year(ActivityDueDate),2) &"</span><span class='activitytime'>"&activityTime &"</span>"
					ElseIf DateDiff("d",ActivityDueDate,Date()) < 0 Then
						col_nextactivityduedate = "<span class='activitydate'>"& Month(ActivityDueDate) & "/" & Day(ActivityDueDate) & "/" & Right(Year(ActivityDueDate),2) &"</span><span class='activitytime'>"&activityTime&"</span>"
					End If 

				 End If 		
			
			col_telemarketer = ""
			If TelemarketerUserNo <> 0 Then 
				col_telemarketer = GetUserDisplayNameByUserNo(TelemarketerUserNo)
			End If	 
			

			

				
			tmpStageDate = GetProspectLastStageChangeDateByProspectNumber(InternalRecordIdentifier)
			tmpStageDate = cdate(tmpStageDate)
				
			col_stagedate = "<span class='activitydatetoday'>"& Month(tmpStageDate) & "/" & Day(tmpStageDate) & "/" & Right(Year(tmpStageDate),2) & "</span>"
			
			col_recycle = ""
			If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then
				col_recycle = "NA"
			ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then
				col_recycle = "<a class='btn btn-success btn-xs' href='recycleInactiveProspect.asp?i="& InternalRecordIdentifier &"' role='button'><i class='fa fa-recycle' aria-hidden='true'></i></a>"	
			ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then
				col_recycle = "<button type='button' class='btn btn-warning btn-xs'><i class='fa fa-recycle' aria-hidden='true'></i></button>"						
			ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then
				col_recycle = "<a class='btn btn-success btn-xs' href='recycleInactiveProspect.asp?i="&InternalRecordIdentifier &"' role='button'><i class='fa fa-recycle' aria-hidden='true'></i></a>"						
			End If	
			
			col_watch = ""
			If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then
				col_watch = "NA"  
			ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then
				col_watch = "<button type='button' class='btn btn-primary btn-xs' data-toggle='modal' data-target='#myProspectingWatchModal' data-tooltip='true' data-title='Watch This Prospect' data-show='true'><i class='fa fa-eye' aria-hidden='true'></i></button>"  
			ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then
				col_watch = "<button type='button' class='btn btn-warning btn-xs' data-tooltip='true' data-title='You Do Not Own This Prospect'><i class='fa fa-times' aria-hidden='true'></i></button>"  					
			ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then 
				col_watch = "<button type='button' class='btn btn-primary btn-xs' data-toggle='modal' data-target='#myProspectingWatchModal' data-tooltip='true' data-title='Watch This Prospect' data-show='true'><i class='fa fa-eye' aria-hidden='true'></i></button>"  
			End If
			
			col_delete = ""
			If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then 
				col_delete = "NA"  
			ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then 
				col_delete ="<button type='button' class='btn btn-danger btn-xs' data-toggle='modal' data-target='#myProspectingDeleteModal' data-tooltip='true' data-title='Delete This Prospect' data-show='true'><i class='fa fa-trash' aria-hidden='true'></i></button>"  
			ElseIf (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then
				col_delete ="<button type='button' class='btn btn-warning btn-xs' data-tooltip='true' data-title='You Do Not Own This Prospect'><i class='fa fa-times' aria-hidden='true'></i></button>"  					
			ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then 
				col_delete ="<button type='button' class='btn btn-danger btn-xs' data-toggle='modal' data-target='#myProspectingDeleteModal' data-tooltip='true' data-title='Delete This Prospect' data-show='true'><i class='fa fa-trash' aria-hidden='true'></i></button>"  
			End If 
				
			IF LEN(JSONdata)>0 THEN
				JSONdata=JSONdata & ","
			END IF
			JSONdata=JSONdata & "{"
			JSONdata=JSONdata & """col_checkbox"":""<input type='checkbox' class='checksingle' name='checksingle' id='"&InternalRecordIdentifier&"' />"""
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_company"":""<a href='viewProspectDetailRecyclePool.asp?i="& InternalRecordIdentifier& "'>" & removeUnusualForJSON(Company) & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_address"":""" & removeUnusualForJSON(Street) & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_city"":""" & removeUnusualForJSON(City) &""""
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_state"":""" & State  & """" '4
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_zip"":""" & PostalCode & """" '5
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_leadsource"":""" & removeUnusualForJSON(GetLeadSourceByNum(LeadSourceNumber))  & """" '6
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_stage"":""" & GetStageByNum(GetProspectCurrentStageByProspectNumber(InternalRecordIdentifier)) & """" '7
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_stagereason"":""" & removeUnusualForJSON(GetStageReasonByStageIntRecID(GetProspectCurrentStageIntRecIDByProspectNumber(InternalRecordIdentifier))) & """" '8
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_stagedate"":""" & mmddyy(GetProspectLastStageChangeDateByProspectNumber(InternalRecordIdentifier))  & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_industry"":""" & removeUnusualForJSON(GetIndustryByNum(IndustryNumber))  & """" '10
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_numemployees"":""" & GetEmployeeRangeByNum(EmployeeRangeNumber)  & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_owner"":""" & GetUserDisplayNameByUserNo(OwnerUserNo)  & """" '12
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_createddate"":""" & tmpCreatedDate & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_createdby"":""" & GetUserDisplayNameByUserNo(CreatedByUserNo) & """" '14
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_telemarketer"":""" & col_telemarketer & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_numpantries"":""" & NumberOfPantries & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_prospectid"":""" & InternalRecordIdentifier & """"
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_recycle"":""" & col_recycle & """" '18
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_watch"":""" & col_watch & """" '19
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_edit"":""<a href='viewProspectDetailRecyclePool.asp?i="&InternalRecordIdentifier&"'><p data-placement='middle' data-toggle='tooltip' title='Edit'><button class='btn btn-success btn-xs' data-title='Edit' data-toggle='modal' data-target='#edit' ><span class='glyphicon glyphicon-pencil'></span></button></p></a>"""
			JSONdata=JSONdata & ","
			JSONdata=JSONdata & """col_delete"":""" & col_delete & """"
						
			
			JSONdata=JSONdata & "}"
							

		%>
	    

		<%	    
			'Response.Flush()
	  		rs8.MoveNext		
		  	Loop
		  
		End If
		
		Set rs8 = Nothing
		'cnn8.Close
		'Set cnn8 = Nothing
		  
		%>	

	    

	                
<%

retData="{""orderby"":""" & orderValue & """,""draw"": " & CLng(Request.QueryString("draw")) & ",""recordsTotal"": " & nRecordCount & ",""recordsFiltered"": " & nRecordCount & ",""data"": [" & JSONdata & "]}"

  

Response.Write retData

function removeUnusualForJSON(value)

'value = Replace(value,"/","") 'slash &#47;
value = Replace(value,"\","") 'backslash &#92;
removeUnusualForJSON=REPLACE(value,"""","&quot;")
	
END FUNCTION


DataConn.Close
Set DataConn = Nothing
%>

 