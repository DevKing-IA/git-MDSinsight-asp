<!--#include file="../inc/header-prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<style>

#loadingmodal {
    display:    none;
    position:   fixed;
    z-index:    1000;
    top:        0;
    left:       0;
    height:     100%;
    width:      100%;
    background: rgba( 255, 255, 255, .8 )  
                url('../img/preloader.gif') 
                50% 50% 
                no-repeat;
}

#loadingmodal {
    overflow: hidden;   
}
#loadingmodal {
    display: block;
}
</style>


<div id="loadingmodal"><h1>Loading <%= GetTerm("New Customer Pool") %></h1></div>

<script>
 
	$(window).on('load', function (e) {
	    $('#loadingmodal').fadeOut(1000);
	})
	
</script>

<%
'****************************************************************************************************
'Read Settings_Reports To See If We Are Loading A Saved Custom Report
'****************************************************************************************************

customFilterReportName = Request.Form("selectFilteredView")
customFilterReportNameQuotes = Replace(Request.Form("selectFilteredView"),"''","'")

If customFilterReportName = "" Then 
	customFilterReportName = MUV_READ("CRMVIEWSTATEWONPOOL")
Else
	dummy = MUV_WRITE("CRMVIEWSTATEWONPOOL",customFilterReportNameQuotes)
End If

If customFilterReportName = "" Then 
	customFilterReportName = "Default"
	dummy = MUV_WRITE("CRMVIEWSTATEWONPOOL","Default")
End If

customFilterReportNameForSQL = Replace(customFilterReportName,"'","''")

'****************************************************************************************************
'Read Settings_Reports To Obtain Filters For Prospecting Grid Data
'****************************************************************************************************
SQL = "SELECT * from Settings_Reports where ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Won' AND UserReportName = '" & customFilterReportNameForSQL & "'"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs= cnn8.Execute(SQL)

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
	ReportSpecificData23  = rs("ReportSpecificData23")
	ReportSpecificData24  = rs("ReportSpecificData24")
	ReportSpecificData25  = rs("ReportSpecificData25")
	ReportSpecificData26  = rs("ReportSpecificData26")
	ReportSpecificData27  = rs("ReportSpecificData27")
	ReportSpecificData28  = rs("ReportSpecificData28")
		
Else
    '****************************************************************
	'Create Default Filter View User Record in Settings_Reports
	'****************************************************************
	
	nextActivityStartDate = dateCustomFormat("01/01/2014")
	nextActivityEndDate = DateSerial(Year(Now()), Month(Now()), Day(Now()))
	
  	SQLCreateUserFilter = "INSERT INTO Settings_Reports(ReportNumber,UserNo,PoolForProspecting,ReportSpecificData9,ReportSpecificData20,ReportSpecificData21,ReportSpecificData22,UserReportName) "
  	SQLCreateUserFilter = SQLCreateUserFilter & " VALUES (1400," & Session("userNo") & ",'Won'," & Session("userNo") & ",'NextActivityScheduledDateRange','" & nextActivityStartDate & "','" & nextActivityEndDate & "','" & customFilterReportNameForSQL & "')"

	Set cnnCreateUserFilter = Server.CreateObject("ADODB.Connection")
	cnnCreateUserFilter.open (Session("ClientCnnString"))
	Set rsCreateUserFilter = Server.CreateObject("ADODB.Recordset")
	rsCreateUserFilter.CursorLocation = 3 
	
	Set rsCreateUserFilter = cnnCreateUserFilter.Execute(SQLCreateUserFilter)
		
	set rsCreateUserFilter = Nothing
	cnnCreateUserFilter.close
	set cnnCreateUserFilter = Nothing
	
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
	
	ElseIf ReportSpecificData2 = "WasUnqualifiedInDateRange" Then
	
		If ReportSpecificData4 <> "" AND ReportSpecificData5 <> "" Then
			startDateStageNotChangedRange = dateCustomFormat(ReportSpecificData4)
			endDateStageNotChangedRange = dateCustomFormat(ReportSpecificData5) 
		End If
		
	ElseIf ReportSpecificData2 = "WasLostInDateRange" Then
	
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

<style type="text/css">
	.col_address {display: none;}
	.col_city {display: none;}
	.col_state {display: none;}
	.col_zip {display: none;}
	.col_owner {display: none;}
	.col_createddate {display: none;}
	.col_createdby {display: none;}
	.col_numemployees {display: none;}
	.col_telemarketer {display: none;}
	.col_numpantries {display: none;}
	.col_prospectid {display: none;}
	.col_leadsource {display: none;}
	.col_industry {display: none;}

</style>

<%

	MaxActivityDaysWarningInit = GetCRMMaxActivityDaysWarning()
	MaxActivityDaysPermittedInit = GetCRMMaxActivityDaysPermitted()
	
	If showHideColumns <> "" Then
		Response.write("<style type='text/css'>")
		If InStr(showHideColumns,"col_address") Then Response.Write(".col_address {display: table-cell;}")
		If InStr(showHideColumns,"col_city") Then Response.Write(".col_city {display: table-cell;}")
		If InStr(showHideColumns,"col_state") Then Response.Write(".col_state {display: table-cell;}")
		If InStr(showHideColumns,"col_zip") Then Response.Write(".col_zip {display: table-cell;}")
		If InStr(showHideColumns,"col_owner") Then Response.Write(".col_owner {display: table-cell;}")
		If InStr(showHideColumns,"col_createddate") Then Response.Write(".col_createddate {display: table-cell; width:120px;}")
		If InStr(showHideColumns,"col_createdby") Then Response.Write(".col_createdby {display: table-cell;}")
		If InStr(showHideColumns,"col_numemployees") Then Response.Write(".col_numemployees {display: table-cell;}")
		If InStr(showHideColumns,"col_telemarketer") Then Response.Write(".col_telemarketer {display: table-cell; width:120px;}")
		If InStr(showHideColumns,"col_numpantries") Then Response.Write(".col_numpantries {display: table-cell; width:100px;}")
		If InStr(showHideColumns,"col_prospectid") Then Response.Write(".col_prospectid {display: table-cell;}")
		If InStr(showHideColumns,"col_leadsource") Then Response.Write(".col_leadsource {display: table-cell;}")
		If InStr(showHideColumns,"col_industry") Then Response.Write(".col_industry {display: table-cell; width:120px;}")
		'If InStr(showHideColumns,"col_stage") Then Response.Write(".col_stage {display: table-cell;}")
		Response.write("</style>")
	Else	
		Response.write("<style type='text/css'>")
		Response.Write(".col_stage {display: table-cell; width:60px;}")
		Response.Write(".col_stagedate {display: table-cell; width:120px;}")
		Response.write("</style>")
	End If

	If UseSettings_Reports = True Then
		Response.write("<style type='text/css'>")
		If upperBoundOwners >= 0 Then Response.Write(".col_owner {display: table-cell;}")
		If ReportSpecificData11 <> "" AND ReportSpecificData12 <> "" Then Response.Write(".col_createddate {display: table-cell; width:120px;}")
		If upperBoundCreatedByUsers >= 0 Then Response.Write(".col_createdby {display: table-cell;}")
		If selectedEmployeeFilterType <> "" Then Response.Write(".col_numemployees {display: table-cell;}")
		If upperBoundTelemarketers >= 0 Then Response.Write(".col_telemarketer {display: table-cell; width:120px;}")
		If selectedPantryFilterType <> "" Then Response.Write(".col_numpantries {display: table-cell; width:100px;}")
		If upperBoundLeadSource >= 0 Then Response.Write(".col_leadsource {display: table-cell;}")
		If upperBoundIndustries >= 0 Then Response.Write(".col_industry {display: table-cell; width:120px;}")
		'If upperBoundStages >= 0 Then Response.Write(".col_stage {display: table-cell;}")
		If ReportSpecificData15 <> "" Then Response.Write(".col_city {display: table-cell;}")
		If ReportSpecificData16 <> "" Then Response.Write(".col_state {display: table-cell;}")
		If ReportSpecificData17 <> "" Then Response.Write(".col_zip {display: table-cell;}")
		
		Response.write("</style>")	
	End If

%>

<!-- Prospecting Custom JS -->
	<!--#include file="mainJQuery.asp"-->
<!-- End Prospecting Custom JS -->

<!-- Prospecting Custom CSS -->
	<link href="mainCSS.css" rel="stylesheet" type="text/css">
<!-- End Prospecting Custom CSS -->

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


Function dateCustomFormat(date)
	x = FormatDateTime(date, 2) 
	d = split(x, "/")
	dateCustomFormat = d(2) & "-" & d(0) & "-" & d(1)
End Function


'Response.Write("<div id=""PleaseWaitPanel"">")
'Response.Write(" <br><br>This may take up to a full minute, please wait...<br><br>")
'Response.Write("<img src=""../img/loading.gif"" />")
'Response.Write("</div>")
'Response.Flush()

%>	
<h1 class="page-header"><i class="fa fa-fw fa-asterisk"></i> <%= GetTerm("New Customer Pool") %>
	<!-- customize !-->
	
	<div class="col pull-right"><h3 style="margin-top: 5px;">Viewing: 
	<% If MUV_READ("CRMVIEWSTATEWONPOOL")= "Default" Then %>
		Default View (My Leads, All Dates)
	<% ElseIf MUV_READ("CRMVIEWSTATEWONPOOL")= "Current" Then %>
		Unsaved Filter View
	<% Else %>
		<%= MUV_READ("CRMVIEWSTATEWONPOOL") %>
	<% End If %>
	</h3></div>
</h1>

<div class="container">
  <div class="row">
    <div class="col-lg-9">
    
	
	<div class="form-group pull-left">
	    <input type="text" class="search form-control" placeholder="What are you looking for?">
	</div>
	<span class="counter pull-right"></span><br clear="all">

      <div class="panel panel-primary">
        <div class="panel-heading"><h4>Viewing:
			<% If MUV_READ("CRMVIEWSTATEWONPOOL")= "Default" Then %>
				Default View (My Leads, All Dates)
			<% ElseIf MUV_READ("CRMVIEWSTATEWONPOOL")= "Current" Then %>
				Unsaved Filter View
			<% Else %>
				<%= MUV_READ("CRMVIEWSTATEWONPOOL") %>
			<% End If %>
        </h4></div> 		
        <div class="panel-body fixed-panel">
			<table class="table table-striped sortable results table-fixed" id="prospectTable">
		
			<thead>
				<th data-defaultsort="disabled" class="col_checkbox" scope="row"><input type="checkbox" id="checkall" name="checkall" /></th>
				<th class="col_company" class="az" data-defaultsign="nospan" scope="row">Company</th>
				<!--
				<th class="col_nextactivity" scope="row">Last Activity</th>
				<th class="col_nextactivitydate" data-defaultsort="desc" data-dateformat="MM/DD/YYYY" data-firstsort="desc" data-defaultsign="month" scope="row">Last Activity Date</th>
				-->
				<th class="col_address" scope="row">Address</th>
				<th class="col_city" scope="row">City</th>
				<th class="col_state" scope="row">State</th>
				<th class="col_zip" scope="row">Zip</th>
				<th class="col_leadsource" scope="row">Lead Source</th>
				<th class="col_stage" scope="row">Stage</th>
				<th class="col_stagedate" scope="row" data-dateformat="MM/DD/YYYY" data-firstsort="desc" data-defaultsign="month" scope="row">Won Date</th>
				<th class="col_industry" scope="row">Industry</th>
				<th class="col_numemployees" scope="row">Num Emp</th>
				<th class="col_owner" scope="row">Owner</th>
				<th class="col_createddate" data-dateformat="MM/DD/YYYY" data-defaultsign="month" scope="row">Created Date</th>
				<th class="col_createdby" scope="row">Created By</th>
				<th class="col_telemarketer" scope="row">Telemarketer</th>
				<th class="col_numpantries" scope="row">Pantries</th>
				<th class="col_prospectid" scope="row">Prospect ID</th>
				<th class="col_edit" data-defaultsort="disabled" scope="row">Details</th>
			</thead>
			
			<tbody>
		
		<!------------------------------------------------------------------------------>	
		<!-- include file that builds a temp SQL table based on selected filters     !-->
		<!------------------------------------------------------------------------------>
		
		<!--#include file="mainWonPoolCustomizeBuildSQLTable.asp"-->
		
		<!------------------------------------------------------------------------------>	
		<!-- END include file that builds a temp SQL table based on selected filters !-->
		<!------------------------------------------------------------------------------>

		<%

		
		'SQL8 = "SELECT *, PR_Prospects.InternalRecordIdentifier AS Expr1  FROM PR_Prospects "
		'SQL8 = SQL8 & " WHERE PR_Prospects.InternalRecordIdentifier IN (SELECT InternalRecordIdentifier FROM zProspectFilter_" & Session("UserNo") & ")"
		'Response.write(SQL8)
		
		'SQL8 = "EXEC dbo.SelectProspects2 @RecsPerPage = " & RecsPerPage & " , @PgNum = " & PgNum & ", @UserNumber = " & Session("UserNo") & " , @db = " & "CorpCoffeedev" & " , @dbown = " & MUV_READ("SQL_OWNER")
		
		SQL8 = "SELECT *, PR_Prospects.InternalRecordIdentifier AS Expr1  FROM PR_Prospects "
		SQL8 = SQL8 & " Inner Join zProspectFilter_" & Session("UserNo") & " ON PR_Prospects.InternalRecordIdentifier = "
		SQL8 = SQL8 & " zProspectFilter_" & Session("UserNo") & ".InternalRecordIdentifier"
		SQL8 = SQL8 & " ORDER BY NextActivityDueDate DESC"

		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.CursorLocation = adUseClient ' Wierd but must do this
		cnn8.open (Session("ClientCnnString"))
		Set rs8 = Server.CreateObject("ADODB.Recordset")
		rs8.CursorLocation = adUseClient  
		
		RecsPerPage = Request.QueryString("RecsPerPage")
		
		If RecsPerPage = "" Then
			RecsPerPage = Request.Form("selProspectsPerPage")
		End If
		
		PgNum = Request.QueryString("PgNum")
		totalProspectCount = TotalNumberOfWonProspects()
		
		If NOT IsNumeric(RecsPerPage) OR RecsPerPage = "" Then RecsPerPage = totalProspectCount
		If NOT IsNumeric(PgNum) OR PgNum = "" Then PgNum = 1
	
		If cInt(RecsPerPage) > 0 Then
			rs8.CacheSize = RecsPerPage
		End If
		
		Set rs8 = cnn8.Execute(SQL8)
		
		
		If not rs8.EOF Then

	
		    rs8.MoveFirst
	   		rs8.PageSize = RecsPerPage
			
			prospectsPerPage = RecsPerPage
			currentlyViewedPage = PgNum
	
			currentViewPageCount  = rs8.PageCount
				
			rs8.AbsolutePage = PgNum
		
			RecCounter = 0 

			Do While Not rs8.EOF AND RecCounter < rs8.PageSize
			
				RecCounter = RecCounter + 1
			
				InternalRecordIdentifier = rs8("Expr1")
				Company = rs8("Company")
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

				ActivityRecID = GetLastProspectActivityNumberByProspectNumber(InternalRecordIdentifier)
				ActivityDueDate = GetLastProspectActivityDueDateByProspectNumber(InternalRecordIdentifier)

		%>
	    
			 <tr>
			    <td class="col_checkbox"><input type="checkbox" class="checksingle" name="checksingle" id="<%= InternalRecordIdentifier %>" /></td>
			    <td class="col_company"><a href="viewProspectDetailWonPool.asp?i=<%= InternalRecordIdentifier %>"><%= Company %></a></td>
				<td class="col_address"><%= Street %></td>
				<td class="col_city"><%= City %></td>
				<td class="col_state"><%= State %></td>
				<td class="col_zip"><%= PostalCode %></td>
				<td class="col_leadsource"><%= GetLeadSourceByNum(LeadSourceNumber) %></td>				
				<td class="col_stage"><a class="getProspectInfo" href="#"><%= GetStageByNum(GetProspectCurrentStageByProspectNumber(InternalRecordIdentifier)) %></a></td> 
				<td class="col_stagedate"><span class="activitydatetoday"><%= mmddyy(GetProspectLastStageChangeDateByProspectNumber(InternalRecordIdentifier)) %></span></td>
				<td class="col_industry"><%= GetIndustryByNum(IndustryNumber) %></td>
				<td class="col_numemployees"><%= GetEmployeeRangeByNum(EmployeeRangeNumber) %></td>
				<td class="col_owner"><%= GetUserDisplayNameByUserNo(OwnerUserNo) %></td>
				<td class="col_createddate" data-value="<%= mmddyyyy(CreatedDate) %>"><%= mmddyyyy(CreatedDate) %></td>
				<td class="col_createdby"><%= GetUserDisplayNameByUserNo(CreatedByUserNo) %></td>

				<% If TelemarketerUserNo <> 0 Then %>
					<td class="col_telemarketer"><%=  GetUserDisplayNameByUserNo(TelemarketerUserNo) %></td>
				<% Else %>
					<td class="col_telemarketer">&nbsp;</td>
				<% End If %>
				
				<td class="col_numpantries"><%= NumberOfPantries %></td>
				<td class="col_prospectid"><%= InternalRecordIdentifier %></td>


			    <td class="col_edit"><a href="viewProspectDetailWonPool.asp?i=<%= InternalRecordIdentifier %>"><p data-placement="top" data-toggle="tooltip" title="Edit"><button class="btn btn-success btn-xs" data-title="Edit" data-toggle="modal" data-target="#edit" ><span class="glyphicon glyphicon-pencil"></span></button></p></a></td>
			    

			</tr>
		<%	    

	  		rs8.MoveNext		
		  	Loop
		  
		End If
		
		Set rs8 = Nothing
		cnn8.Close
		Set cnn8 = Nothing
		  
		%>	

	    
	    </tbody>
	        
	</table>

	</div></div><!-- end panels -->
	
	<div class="clearfix"></div>
	
	<form name="frmRecordsPerPage" id="frmRecordsPerPage" method="post" action="mainWonPool.asp">
		<div class="col-lg-2 col-md-2 pull-left" style="margin-left:-15px;">
			<select class="form-control" name="selProspectsPerPage" onchange="this.form.submit()">
				<option value="5" <% If prospectsPerPage = 5 Then Response.Write("selected") %>>5 <%= GetTerm("Prospects") %> per Page</option>
				<option value="10" <% If prospectsPerPage = 10 Then Response.Write("selected") %>>10 <%= GetTerm("Prospects") %> per Page</option>
				<option value="15" <% If prospectsPerPage = 15 Then Response.Write("selected") %>>15 <%= GetTerm("Prospects") %> per Page</option>
				<option value="20" <% If prospectsPerPage = 20 Then Response.Write("selected") %>>20 <%= GetTerm("Prospects") %> per Page</option>
				<option value="25" <% If prospectsPerPage = 25 Then Response.Write("selected") %>>25 <%= GetTerm("Prospects") %> per Page</option>
				<option value="50" <% If prospectsPerPage = 50 Then Response.Write("selected") %>>50 <%= GetTerm("Prospects") %> per Page</option>
				<option value="100" <% If prospectsPerPage = 100 Then Response.Write("selected") %>>100 <%= GetTerm("Prospects") %> per Page</option>
				<option value="<%= totalProspectCount %>" <% If prospectsPerPage = totalProspectCount Then Response.Write("selected") %>>Show All</option>
			</select>
		</div>
	</form>	
	
	<ul class="pagination pull-right" style="margin:0px;">
	<%
	
	
		If cInt(currentlyViewedPage) = cInt(currentViewPageCount) Then
			curPage = currentlyViewedPage - 1
		Else
			curPage = currentlyViewedPage
		End if		
				
		If cInt(currentlyViewedPage) > 1 Then		
			PageNav = "<li><a href='mainWonPool.asp?PgNum=" & 1 & " & RecsPerPage="&RecsPerPage & "'>First</a></li>"
			PageNav = PageNav & "<li><a href='mainWonPool.asp?PgNum=" & currentlyViewedPage - 1 & "&RecsPerPage="& RecsPerPage & "'><span class='glyphicon glyphicon-chevron-left'></span></a></li>"
		End If
	
		If cInt(currentViewPageCount) > 7 Then

			If (currentlyViewedPage) > (currentViewPageCount - 7) Then currentlyViewedPage = (currentlyViewedPage - 6)
			If (currentlyViewedPage) < 1 Then currentlyViewedPage = 2
		
			For Pages = currentlyViewedPage to (currentlyViewedPage + 7)
			
				If cInt(Pages) <> cInt(currentlyViewedPage) Then
					if (Pages <= currentViewPageCount) then
						PageNav = PageNav & "<li><a href='mainWonPool.asp?PgNum=" & Pages & "&RecsPerPage=" & RecsPerPage & "'>" & Pages & "</a></li>"
					end if
				Else
					if (Pages <= currentViewPageCount) then
						PageNav = PageNav & "<li class='active'><a href='mainWonPool.asp?PgNum=" & Pages & "&RecsPerPage="& RecsPerPage & "'>" & Pages & "</a></li>"					
					end if
				End If
			Next	
		
		Else
		
			For Pages = 1 to currentViewPageCount
			
				If cInt(Pages) <> cInt(currentlyViewedPage) Then
					if (Pages <= currentViewPageCount) then
						PageNav = PageNav & "<li><a href='mainWonPool.asp?PgNum=" & Pages & "&RecsPerPage=" & RecsPerPage & "'>" & Pages & "</a></li>"
					end if
				Else
					if (Pages <= currentViewPageCount) then
						PageNav = PageNav & "<li class='active'><a href='mainWonPool.asp?PgNum=" & Pages & "&RecsPerPage="& RecsPerPage & "'>" & Pages & "</a></li>"					
					end if
				End If
			Next
				
		End If

		If cInt(currentlyViewedPage) < cInt(currentViewPageCount) Then
			PageNav = PageNav & "<li><a href='mainWonPool.asp?PgNum="& curPage +1 &"& RecsPerPage="& RecsPerPage & "'><span class='glyphicon glyphicon-chevron-right'></span></a></li>"
			PageNav = PageNav & "<li><a href='mainWonPool.asp?PgNum="& currentViewPageCount & "&RecsPerPage="& RecsPerPage & "'>Last</a></li>"
		End If
		
		Response.Write(PageNav)
	%>
	</ul>
	                

</div>

<!-- Bootstrap table -->
<div class="col-lg-3">
  <div class="panel panel-default">
   <div class="panel-heading" id="TotalNumberOfProspects">Currently Viewing <strong><%= TotalNumberOfWonProspects() %></strong> Total Prospects In <%= GetTerm("New Customer Pool") %></div>
	<div class="panel-body fixed-panel panel-padding">


	<div class="row">
	
		<div class="col-lg-12">
			<div class="col-lg-6">
				<button class="btn btn-indigo btn-lg btn-block" id="customizeView" data-toggle="modal" data-target=".bs-modal-show-hide-columns-won-pool">
					<i class="fa fa-columns" aria-hidden="true"></i>&nbsp;Column View
				</button>
			</div>
			
			<div class="col-lg-6">
				<button class="btn btn-red btn-lg btn-block" id="filterView" data-toggle="modal" data-target=".bs-modal-filter-prospecting-data-won-pool">
					<i class="fa fa-filter " aria-hidden="true"></i>&nbsp;Filter View
				</button>			
			</div>
			
		</div>
		
	</div>
	
	<div class="row">

		<% If UseSettings_Reports = True OR MUV_READ("CRMSTARTDATE") <> "" OR MUV_READ("CRMENDDATE") <> "" Then %>
			<a href="mainWonPoolCustomizeClearValues.asp"><button class="btn btn-warning btn-lg btn-block" id="resetView">
				<i class="fa fa-refresh" aria-hidden="true"></i>&nbsp;Clear Filters &amp; Reset View
			</button></a>
		<% End If %>
		
	</div>
	
	<hr class="style7">
	
	<% If userIsAdmin(Session("userNo")) = True OR userIsInsideSalesManager(Session("userNo")) = True OR userIsOutsideSalesManager(Session("userNo")) = True Then %>
		<div class="row" style="margin-top:-15px;">
		    <button type="button" class="btn btn-darkblue btn-lg btn-block" data-toggle="modal" data-target="#viewProspectFilterViewAllWonPool" onclick="location.href='mainCreateWonPoolViewAllProspectsFilter.asp';">
		        <i class="fa fa-users"></i>&nbsp;View All Prospects
		    </button>
		</div>
	<% Else %>
		<div class="row" style="margin-top:-15px;">
		    <button type="button" class="btn btn-lightgreen btn-lg btn-block" data-toggle="modal" data-target="#viewProspectFilterViewAllWonPool" onclick="location.href='mainCreateWonPoolViewAllProspectsFilter.asp';">
		        <i class="fa fa-users"></i>&nbsp;View All My Prospects
		    </button>
		</div>
	<% End If %>

		
	<form action="<%= BaseURL %>prospecting/mainWonPool.asp" method="POST" name="frmToggleLeadAndDateView" id="frmToggleLeadAndDateView">
		
	<div class="row">
	
			<hr class="style7">
			
	      	<%'Report View Name Dropdown
	      	
	      	userHasSavedViews = false
	      	 
	  	  	SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND "
	  	  	SQL = SQL & " PoolForProspecting = 'Won' AND UserReportName <> 'Current'  AND UserReportName <> 'Default' AND UserReportName <> 'All Prospects' ORDER BY UserReportName "
	
			Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.open (Session("ClientCnnString"))
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3 
			Set rs = cnn8.Execute(SQL)
		
			If NOT rs.EOF Then
			
				userHasSavedViews = true
			%>
				<!-- Display Report View Names -->
				<select class="form-control when-line" style="width: 100%; display :inline; margin-left:0px;" name="selectFilteredView" id="selectFilteredView" onchange="$('#loadingmodal').show();this.form.submit()">
				<option value=""> -- Select Custom View -- </option>
				<option value="Default" <% If MUV_READ("CRMVIEWSTATEWONPOOL") = "Default" Then Response.Write("selected") %>>Default View (My Leads, All Dates)</option>
				<%
					Do
						selReportName = Replace(rs("UserReportName"),"''","'")
						If MUV_READ("CRMVIEWSTATEWONPOOL") = selReportName Then
							%><option value="<%= selReportName %>" selected="selected"><%= selReportName %></option><%
						Else
							%><option value="<%= selReportName %>"><%= selReportName %></option><%
						End If
						rs.movenext
					Loop until rs.eof
				%>			
				</select>
				<!-- End Display Report View Names -->
			<%
			End If
			set rs = Nothing
			cnn8.close
			set cnn8 = Nothing
	      	%>


	</div>
	
	<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
		<div class="row">
 			
		</div>
	<% Else %>
			
		<% If userHasSavedViews = true Then %>

			<div class="row">
			    <button type="button" class="btn btn-primary btn-lg btn-block" data-toggle="modal" data-target="#saveAsNewProspectFilterViewWonPool" id="btnSaveAsNewProspectFilterViewWonPool">
			        <i class="far fa-save"></i>&nbsp;Save View
			    </button>
			</div>
			
			<div class="row" style="margin-bottom:15px;">
				<div class="col-lg-12">
					<div class="col-lg-6">
						<button type="button" class="btn btn-darkgray btn-lg btn-block" data-toggle="modal" data-target="#editFilterViewNameWonPool" id="btnRenameProspectFilterViewWonPool">
							<i class="fa fa-pencil-square" aria-hidden="true"></i>&nbsp;Rename View
						</button>
					</div>
					
					<div class="col-lg-6">
						<button type="button" class="btn btn-red btn-lg btn-block" data-toggle="modal" data-target="#deleteProspectViewWonPool" id="btnDeleteProspectFilterViewWonPool">
							<i class="fas fa-trash-alt-o" aria-hidden="true"></i>&nbsp;Delete View
						</button>					    
					</div>
				</div>
			</div>
		
		<% Else %>
		
			<div class="row" style="margin-bottom:15px;">
			    <button type="button" class="btn btn-primary btn-lg btn-block" data-toggle="modal" data-target="#saveAsNewProspectFilterViewWonPool" id="btnSaveAsNewProspectFilterViewWonPool">
			        <i class="far fa-save"></i>&nbsp;Save View
			    </button>
			</div>
			
		<% End If %>


	
	<% End If %>
	
	</form>
</div>
</div>
</div>
</div>

<!--***********************************************************************************-->
<!--ALL MODAL WINDOWS USED ON THE PAGE - CUSTOMIZATION, EDIT ACTIVITY, EDIT STAGE, ETC.-->
<!--ARE STORED IN THIS FILE: -->
<!--***********************************************************************************-->
<!--#include file="mainWonPoolModals.asp"-->
<!--***********************************************************************************-->
<!--***********************************************************************************-->
<script type="text/javascript" src="http://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>
<script type="text/javascript" src="http://cdn.jsdelivr.net/bootstrap.daterangepicker/2/daterangepicker.js"></script>
<link rel="stylesheet" type="text/css" href="http://cdn.jsdelivr.net/bootstrap.daterangepicker/2/daterangepicker.css">


<!-- datepicker for edit next activity modal !-->
	<link href="<%= baseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.css" rel="stylesheet" type="text/css">
	<script src="<%= baseURL %>js/bootstrap-datetimepicker/bootstrap-datetimepicker.js" type="text/javascript"></script>
<!-- end datepicker for edit next activity moda !-->
	

<!--#include file="../inc/footer-main.asp"-->


 