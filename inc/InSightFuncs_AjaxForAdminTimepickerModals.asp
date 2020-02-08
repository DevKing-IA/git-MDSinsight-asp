<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
 
'Sub ClearLoginAccessForExistingUser()
'Sub ClearLoginAccessNewForUser()
'Sub UpdateLoginAccessForExistingUser()
'Sub UpdateLoginAccessForNewUser()
'Sub GetContentForAutoFilterGenerationScheduler()
'Sub GetContentForFieldServiceNotesReportScheduler()
'Sub GetContentForServiceTicketCarryoverReportScheduler()
'Sub GetContentForServiceTicketThresholdReportScheduler()
'Sub GetContentForProspectingSnapshotReportScheduler()
'Sub GetContentForProspectingWeeklyAgendaReportScheduler()
'Sub GetContentForDailyAPIActivityByPartnerReportScheduler()
'Sub GetContentForDailyInventoryAPIActivityByPartnerReportScheduler()
'Sub GetContentForInventoryProductChangesReportScheduler()
'Sub GetContentForAutomaticCustomerAnalysisSummary1ReportScheduler()
'Sub GetContentForMCSActivityReportScheduler()
'Sub GetContentForOrderAPIN2KReportScheduler()
'Sub GetContentForAccountsReceivableN2KReportScheduler()
'Sub GetContentForEquipmentN2KReportScheduler()
'Sub GetContentForGlobalSettingsN2KReportScheduler()
'Sub GetContentForInventoryN2KReportScheduler()

'***************************************************
'End List of all the AJAX functions & subs
'***************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'ALL AJAX MODAL SUBROUTINES AND FUNCTIONS BELOW THIS AREA

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

action = Request("action")

Select Case action
	Case "ClearLoginAccessForExistingUser" 
		ClearLoginAccessForExistingUser()
	Case "ClearLoginAccessNewForUser"
		ClearLoginAccessNewForUser()	
	Case "UpdateLoginAccessForExistingUser"
		UpdateLoginAccessForExistingUser()
	Case "UpdateLoginAccessForNewUser"
		UpdateLoginAccessForNewUser()
    Case "updateCustomOrDefault"
        updateCustomOrDefault()
	Case "GetContentForAutoFilterGenerationScheduler"
		GetContentForAutoFilterGenerationScheduler()  
	Case "GetContentForFieldServiceNotesReportScheduler"
		GetContentForFieldServiceNotesReportScheduler()
	Case "GetContentForServiceTicketCarryoverReportScheduler"
		GetContentForServiceTicketCarryoverReportScheduler()
	Case "GetContentForServiceTicketThresholdReportScheduler"
		GetContentForServiceTicketThresholdReportScheduler()
	Case "GetContentForProspectingSnapshotReportScheduler"
		GetContentForProspectingSnapshotReportScheduler()
	Case "GetContentForProspectingWeeklyAgendaReportScheduler"
		GetContentForProspectingWeeklyAgendaReportScheduler()
	Case "GetContentForDailyAPIActivityByPartnerReportScheduler"
		GetContentForDailyAPIActivityByPartnerReportScheduler()
	Case "GetContentForDailyInventoryAPIActivityByPartnerReportScheduler"
		GetContentForDailyInventoryAPIActivityByPartnerReportScheduler()
	Case "GetContentForInventoryProductChangesReportScheduler"
		GetContentForInventoryProductChangesReportScheduler()
	Case "GetContentForAutomaticCustomerAnalysisSummary1ReportScheduler"
		GetContentForAutomaticCustomerAnalysisSummary1ReportScheduler()
	Case "GetContentForMCSActivityReportScheduler"
		GetContentForMCSActivityReportScheduler()
	Case "GetContentForOrderAPIN2KReportScheduler"
		GetContentForOrderAPIN2KReportScheduler()
	Case "GetContentForAccountsReceivableN2KReportScheduler"
		GetContentForAccountsReceivableN2KReportScheduler()
	Case "GetContentForEquipmentN2KReportScheduler"
		GetContentForEquipmentN2KReportScheduler()
	Case "GetContentForGlobalSettingsN2KReportScheduler"
		GetContentForGlobalSettingsN2KReportScheduler()
	Case "GetContentForInventoryN2KReportScheduler"
		GetContentForInventoryN2KReportScheduler()

End Select

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub UpdateLoginAccessForExistingUser()

	userNo = Request.Form("userNo") 
	jsonString = Request.Form("jsonString") 
	
	'Response.Write(jsonString)
	
	'********************************************************************
	'When a user selects new login restricted access times, we are rebuilding ALL records in SC_UserRestrictedLoginSchedule,
	'so we need to delete all existing records first
	'********************************************************************
	
	SQLDelete = "DELETE FROM SC_UserRestrictedLoginSchedule WHERE userNo = " & userNo
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	cnnDelete.close
	
	'********************************************************************
	'Prepare the jsonString for parsing by removing extraneous characters
	'********************************************************************
	'remove the opening [ in the string
	jsonString = Right(jsonString, Len(jsonString) - 1)
	
	'remove the closing ] in the string
	jsonString = Left(jsonString,len(jsonString)-1)
	
	'remove double quotes from the string
	jsonString = Replace(jsonString, """","")
	
	'********************************************************************


	'********************************************************************
	'Now build the new login records and insert them into SC_UserRestrictedLoginSchedule
	
	Set cnnInsert = Server.CreateObject("ADODB.Connection")
	cnnInsert.open (Session("ClientCnnString"))
	Set rsInsert = Server.CreateObject("ADODB.Recordset")
	rsInsert.CursorLocation = 3 

	If InStr(jsonString,"},{") Then
	
		'Multiple days have been selected with restricted login access
		
		jsonArray = Split(jsonString, "},{")
		
		for i = 0 to Ubound(jsonArray)
		
			singleDayString = Split(jsonArray(i),",")
			
			''singleDayString[0] = Contains Day Number
			''singleDayString[1] = Contains Day Number Restricted Start Time
			''singleDayString[2] = Contains Day Number Restricted End Time
			
			'remove opening bracket from day string
			singleDayString(0) = Replace(singleDayString(0), "{", "")
			'remove closing bracket from day string
			singleDayString(2) = Replace(singleDayString(2), "}", "")
			
			dayNumber = cInt(Right(singleDayString(0), 1))
			startTime = Right(singleDayString(1), 5)
			endTime = Right(singleDayString(2), 5)
			
			SQLInsert = "INSERT INTO SC_UserRestrictedLoginSchedule (UserNo, DayNo, StartRestrictedTime, EndRestrictedTime) VALUES "
			SQLInsert = SQLInsert & " (" & userNo & "," & dayNumber & ",'" & startTime & "','" & endTime & "')"
			
			Set rsInsert = cnnInsert.Execute(SQLInsert)
			
		next

	
	Else
		'Then only one day has been selected
		
		singleDayString = Split(jsonString,",")
		
		''singleDayString[0] = Contains Day Number
		''singleDayString[1] = Contains Day Number Restricted Start Time
		''singleDayString[2] = Contains Day Number Restricted End Time

		'remove opening bracket from day string
		singleDayString(0) = Replace(singleDayString(0), "{", "")
		'remove closing bracket from day string
		singleDayString(2) = Replace(singleDayString(2), "}", "")

		dayNumber = cInt(Right(singleDayString(0), 1))
		startTime = Right(singleDayString(1), 5)
		endTime = Right(singleDayString(2), 5)
		
		SQLInsert = "INSERT INTO SC_UserRestrictedLoginSchedule (UserNo, DayNo, StartRestrictedTime, EndRestrictedTime) VALUES "
		SQLInsert = SQLInsert & " (" & userNo & "," & dayNumber & ",'" & startTime & "','" & endTime & "')"
		
		Set rsInsert = cnnInsert.Execute(SQLInsert)

	End If
	
	cnnInsert.close
	
	'********************************************************************

	'day:0,start:00:00,end:24:00
	'[{"day":0,"start":"00:00","end":"24:00"}]
	'[{"day":0,"start":"00:00","end":"24:00"},{"day":6,"start":"01:00","end":"24:00"}]
	'[{"day":0,"start":"00:00","end":"24:00"},{"day":3,"start":"00:00","end":"08:00"},{"day":3,"start":"13:00","end":"19:00"},{"day":6,"start":"00:00","end":"24:00"}]

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub UpdateLoginAccessForNewUser()

	jsonString = Request.Form("jsonString") 
	
	'Response.Write(jsonString)
	
	'********************************************************************
	'When a user selects new login restricted access times, we are rebuilding ALL records in SC_UserRestrictedLoginSchedule,
	'so we need to delete all existing records first
	'********************************************************************
	
	SQLDelete = "DELETE FROM SC_UserRestrictedLoginSchedule WHERE userNo = -1"
	
	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	cnnDelete.close
	
	'********************************************************************
	'Prepare the jsonString for parsing by removing extraneous characters
	'********************************************************************
	'remove the opening [ in the string
	jsonString = Right(jsonString, Len(jsonString) - 1)
	
	'remove the closing ] in the string
	jsonString = Left(jsonString,len(jsonString)-1)
	
	'remove double quotes from the string
	jsonString = Replace(jsonString, """","")
	
	'********************************************************************


	'********************************************************************
	'Now build the new login records and insert them into SC_UserRestrictedLoginSchedule
	
	Set cnnInsert = Server.CreateObject("ADODB.Connection")
	cnnInsert.open (Session("ClientCnnString"))
	Set rsInsert = Server.CreateObject("ADODB.Recordset")
	rsInsert.CursorLocation = 3 

	If InStr(jsonString,"},{") Then
	
		'Multiple days have been selected with restricted login access
		
		jsonArray = Split(jsonString, "},{")
		
		for i = 0 to Ubound(jsonArray)
		
			singleDayString = Split(jsonArray(i),",")
			
			''singleDayString[0] = Contains Day Number
			''singleDayString[1] = Contains Day Number Restricted Start Time
			''singleDayString[2] = Contains Day Number Restricted End Time
			
			'remove opening bracket from day string
			singleDayString(0) = Replace(singleDayString(0), "{", "")
			'remove closing bracket from day string
			singleDayString(2) = Replace(singleDayString(2), "}", "")
			
			dayNumber = cInt(Right(singleDayString(0), 1))
			startTime = Right(singleDayString(1), 5)
			endTime = Right(singleDayString(2), 5)
			
			SQLInsert = "INSERT INTO SC_UserRestrictedLoginSchedule (UserNo, DayNo, StartRestrictedTime, EndRestrictedTime) VALUES "
			SQLInsert = SQLInsert & " (-1," & dayNumber & ",'" & startTime & "','" & endTime & "')"
			
			Set rsInsert = cnnInsert.Execute(SQLInsert)
			
		next

	
	Else
		'Then only one day has been selected
		
		singleDayString = Split(jsonString,",")
		
		''singleDayString[0] = Contains Day Number
		''singleDayString[1] = Contains Day Number Restricted Start Time
		''singleDayString[2] = Contains Day Number Restricted End Time

		'remove opening bracket from day string
		singleDayString(0) = Replace(singleDayString(0), "{", "")
		'remove closing bracket from day string
		singleDayString(2) = Replace(singleDayString(2), "}", "")

		dayNumber = cInt(Right(singleDayString(0), 1))
		startTime = Right(singleDayString(1), 5)
		endTime = Right(singleDayString(2), 5)
		
		SQLInsert = "INSERT INTO SC_UserRestrictedLoginSchedule (UserNo, DayNo, StartRestrictedTime, EndRestrictedTime) VALUES "
		SQLInsert = SQLInsert & " (-1," & dayNumber & ",'" & startTime & "','" & endTime & "')"
		
		Set rsInsert = cnnInsert.Execute(SQLInsert)

	End If
	
	cnnInsert.close
	
	'********************************************************************

	'day:0,start:00:00,end:24:00
	'[{"day":0,"start":"00:00","end":"24:00"}]
	'[{"day":0,"start":"00:00","end":"24:00"},{"day":6,"start":"01:00","end":"24:00"}]
	'[{"day":0,"start":"00:00","end":"24:00"},{"day":3,"start":"00:00","end":"08:00"},{"day":3,"start":"13:00","end":"19:00"},{"day":6,"start":"00:00","end":"24:00"}]

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub ClearLoginAccessForExistingUser()

	userNo = Request.Form("userNo") 
	
	Set rsSaveEquivCustID = Server.CreateObject("ADODB.Recordset")
	rsSaveEquivCustID.CursorLocation = 3 
	
	If userNo <> "" Then

		SQLDelete = "DELETE FROM SC_UserRestrictedLoginSchedule WHERE UserNo = " & userNo

		Set cnnDelete = Server.CreateObject("ADODB.Connection")
		cnnDelete.open (Session("ClientCnnString"))
		Set rsDelete = Server.CreateObject("ADODB.Recordset")
		rsDelete.CursorLocation = 3 
		Set rsDelete = cnnDelete.Execute(SQLDelete)
		cnnDelete.close
		
		Response.Write("Success")
		
	Else
		Response.Write("Cannot Clear User Access Table, Invalid Data")
		
	End If

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


Sub ClearLoginAccessNewForUser()
	
	Set rsSaveEquivCustID = Server.CreateObject("ADODB.Recordset")
	rsSaveEquivCustID.CursorLocation = 3 

	SQLDelete = "DELETE FROM SC_UserRestrictedLoginSchedule WHERE UserNo = -1"

	Set cnnDelete = Server.CreateObject("ADODB.Connection")
	cnnDelete.open (Session("ClientCnnString"))
	Set rsDelete = Server.CreateObject("ADODB.Recordset")
	rsDelete.CursorLocation = 3 
	Set rsDelete = cnnDelete.Execute(SQLDelete)
	cnnDelete.close
	
	Response.Write("Success")
		

End Sub


Sub updateCustomOrDefault()
    emailid = Request.Form("id") 
    emailtype = Request.Form("type") 
	
	SQLUpdate = "UPDATE SC_EmailCustomization SET customOrDefault='" & emailtype & "' WHERE InternalRecordIdentifier = " & emailid

	Set cnnUpdate = Server.CreateObject("ADODB.Connection")
	cnnUpdate.open (Session("ClientCnnString"))
	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.CursorLocation = 3 
	Set rsUpdate = cnnUpdate.Execute(SQLUpdate)
	cnnUpdate.close
	
	Response.Write("Success")
		

End Sub


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForAutoFilterGenerationScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerAutoFilterGenSchedulerSunday').timepicker({
	            minuteStep: 15,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutoFilterGenSchedulerMonday').timepicker({
	            minuteStep: 15,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutoFilterGenSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutoFilterGenSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutoFilterGenSchedulerThursday').timepicker({
	            minuteStep: 15,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutoFilterGenSchedulerFriday').timepicker({
	            minuteStep: 15,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutoFilterGenSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '12:00 AM'
	        });
	        
		
			var initGenTimeSunday = $('#txtGenerateFilterTicketSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerAutoFilterGenSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtGenerateFilterTicketMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerAutoFilterGenSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtGenerateFilterTicketTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerAutoFilterGenSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtGenerateFilterTicketWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerAutoFilterGenSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtGenerateFilterTicketThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerAutoFilterGenSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtGenerateFilterTicketFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerAutoFilterGenSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtGenerateFilterTicketSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerAutoFilterGenSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerAutoFilterGenSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoAutoFilterTicketGenSunday").prop( "checked", false );		    
		    });
		    $('#timepickerAutoFilterGenSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoAutoFilterTicketGenMonday").prop( "checked", false );		    
		    });
		    $('#timepickerAutoFilterGenSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoAutoFilterTicketGenTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerAutoFilterGenSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoAutoFilterTicketGenWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerAutoFilterGenSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoAutoFilterTicketGenThursday").prop( "checked", false );		    
		    });
		    $('#timepickerAutoFilterGenSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoAutoFilterTicketGenFriday").prop( "checked", false );		    
		    });
		    $('#timepickerAutoFilterGenSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoAutoFilterTicketGenSaturday").prop( "checked", false );		    
		    });
  
			$('#btnEditDeliveryAlertSave').on('click', function(e) {
			
			    //get data-id attribute of the clicked alert
			    var invoiceNum = $("#txtIvsNum").val();
			    var condition = $("#selCondition").val();
			    var emailto = $("#selEmailto").val();
			    var addlemails = $("#txtAdditionalEmails").val();
			    var textto = $("#selTextto").val();
			    var addltexts = $("#txtAdditionalTexts").val();
			    				
				//turn off the automatic page refresh
				//$('#switchAutomaticRefresh').prop('checked', true).trigger("change");
			    		    		
		    	$.ajax({
					type:"POST",
					url:"../inc/InSightFuncs_AjaxForRoutingModals.asp",
					data: "action=EditAlertFromDeliveryBoardModal&invoiceNum=" + encodeURIComponent(invoiceNum) + "&condition=" + encodeURIComponent(condition) + "&emailto=" + encodeURIComponent(emailto) + "&addlemails=" + encodeURIComponent(addlemails) + "&textto=" + encodeURIComponent(textto) + "&addltexts=" + encodeURIComponent(addltexts),
					success: function(response)
					 {
					 	location.reload();
		             }
				});
	    	});	
	    	
			$("#chkNoAutoFilterTicketGenSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutoFilterGenSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutoFilterGenSchedulerSunday').timepicker('setTime', '12:00 AM');
			    }
			});
			    	
			$("#chkNoAutoFilterTicketGenMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutoFilterGenSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutoFilterGenSchedulerMonday').timepicker('setTime', '12:00 AM');
			    }
			});
	    	
			$("#chkNoAutoFilterTicketGenTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutoFilterGenSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutoFilterGenSchedulerTuesday').timepicker('setTime', '12:00 AM');
			    }
			});

			$("#chkNoAutoFilterTicketGenWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutoFilterGenSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutoFilterGenSchedulerWednesday').timepicker('setTime', '12:00 AM');
			    }
			});

			$("#chkNoAutoFilterTicketGenThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutoFilterGenSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutoFilterGenSchedulerThursday').timepicker('setTime', '12:00 AM');
			    }
			});

			$("#chkNoAutoFilterTicketGenFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutoFilterGenSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutoFilterGenSchedulerFriday').timepicker('setTime', '12:00 AM');
			    }
			});

			$("#chkNoAutoFilterTicketGenSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutoFilterGenSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutoFilterGenSchedulerSaturday').timepicker('setTime', '12:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing filter generation schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,12:00 AM,12:00 AM,12:00 AM,12:00 AM,12:00 AM,12:00 AM,12:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_FilterGeneration = ""
	GenerateFilterTicketSunday = ""
	GenerateFilterTicketMonday = ""
	GenerateFilterTicketTuesday = ""
	GenerateFilterTicketWednesday = ""
	GenerateFilterTicketThursday = ""
	GenerateFilterTicketFriday = ""
	GenerateFilterTicketSaturday = ""
	GenerateFilterTicketSundayTime = ""
	GenerateFilterTicketMondayTime = ""
	GenerateFilterTicketTuesdayTime = ""
	GenerateFilterTicketWednesdayTime = ""
	GenerateFilterTicketThursdayTime = ""
	GenerateFilterTicketFridayTime = ""
	GenerateFilterTicketSaturdayTime = ""
	RunFieldServiceNotesReportIfClosed = ""
	RunFieldServiceNotesReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_FieldService"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_FilterGeneration = rsFieldServiceSettings("Schedule_FilterGeneration")
		
		Schedule_FilterGenerationSettings = Split(Schedule_FilterGeneration,",")

		GenerateFilterTicketSunday = cInt(Schedule_FilterGenerationSettings(0))
		GenerateFilterTicketMonday = cInt(Schedule_FilterGenerationSettings(1))
		GenerateFilterTicketTuesday = cInt(Schedule_FilterGenerationSettings(2))
		GenerateFilterTicketWednesday = cInt(Schedule_FilterGenerationSettings(3))
		GenerateFilterTicketThursday = cInt(Schedule_FilterGenerationSettings(4))
		GenerateFilterTicketFriday = cInt(Schedule_FilterGenerationSettings(5))
		GenerateFilterTicketSaturday = cInt(Schedule_FilterGenerationSettings(6))
		GenerateFilterTicketSundayTime = Schedule_FilterGenerationSettings(7)
		GenerateFilterTicketMondayTime = Schedule_FilterGenerationSettings(8)
		GenerateFilterTicketTuesdayTime = Schedule_FilterGenerationSettings(9)
		GenerateFilterTicketWednesdayTime = Schedule_FilterGenerationSettings(10)
		GenerateFilterTicketThursdayTime = Schedule_FilterGenerationSettings(11)
		GenerateFilterTicketFridayTime = Schedule_FilterGenerationSettings(12)
		GenerateFilterTicketSaturdayTime = Schedule_FilterGenerationSettings(13)
		RunFilterTicketAutoGenIfClosed = cInt(Schedule_FilterGenerationSettings(14))
		RunFilterTicketAutoGenIfClosingEarly = cInt(Schedule_FilterGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to auto generate filter tickets on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to generate tickets on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> Filter Tickets Can Only Be Generated 3:00 PM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GenerateFilterTicketSunday) = 0 Then %>
				  		<input id="timepickerAutoFilterGenSchedulerSunday" name="txtAutoFilterGenSchedulerSundayTime" type="text" value="" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketSundayInit" id="txtGenerateFilterTicketSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutoFilterGenSchedulerSunday" name="txtAutoFilterGenSchedulerSundayTime" type="text" value="<%= GenerateFilterTicketSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketSundayInit" id="txtGenerateFilterTicketSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(GenerateFilterTicketSunday) = 0 Then %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenSunday" name="chkNoAutoFilterTicketGenSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenSunday" name="chkNoAutoFilterTicketGenSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GenerateFilterTicketMonday) = 0 Then %>
				  		<input id="timepickerAutoFilterGenSchedulerMonday" type="text" name="txtAutoFilterGenSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketMondayInit" id="txtGenerateFilterTicketMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutoFilterGenSchedulerMonday" type="text" name="txtAutoFilterGenSchedulerMondayTime" value="<%= GenerateFilterTicketMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketMondayInit" id="txtGenerateFilterTicketMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GenerateFilterTicketMonday) = 0 Then %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenMonday" name="chkNoAutoFilterTicketGenMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenMonday" name="chkNoAutoFilterTicketGenMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GenerateFilterTicketTuesday) = 0 Then %>
				  		<input id="timepickerAutoFilterGenSchedulerTuesday" type="text" name="txtAutoFilterGenSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketTuesdayInit" id="txtGenerateFilterTicketTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutoFilterGenSchedulerTuesday" type="text" name="txtAutoFilterGenSchedulerTuesdayTime" value="<%= GenerateFilterTicketTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketTuesdayInit" id="txtGenerateFilterTicketTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GenerateFilterTicketTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenTuesday" name="chkNoAutoFilterTicketGenTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenTuesday" name="chkNoAutoFilterTicketGenTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GenerateFilterTicketWednesday) = 0 Then %>
				  		<input id="timepickerAutoFilterGenSchedulerWednesday" type="text" name="txtAutoFilterGenSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketWednesdayInit" id="txtGenerateFilterTicketWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutoFilterGenSchedulerWednesday" type="text" name="txtAutoFilterGenSchedulerWednesdayTime" value="<%= GenerateFilterTicketWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketWednesdayInit" id="txtGenerateFilterTicketWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GenerateFilterTicketWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenWednesday" name="chkNoAutoFilterTicketGenWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenWednesday" name="chkNoAutoFilterTicketGenWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GenerateFilterTicketThursday) = 0 Then %>
				  		<input id="timepickerAutoFilterGenSchedulerThursday" type="text" name="txtAutoFilterGenSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketThursdayInit" id="txtGenerateFilterTicketThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutoFilterGenSchedulerThursday" type="text" name="txtAutoFilterGenSchedulerThursdayTime" value="<%= GenerateFilterTicketThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketThursdayInit" id="txtGenerateFilterTicketThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GenerateFilterTicketThursday) = 0 Then %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenThursday" name="chkNoAutoFilterTicketGenThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenThursday" name="chkNoAutoFilterTicketGenThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GenerateFilterTicketFriday) = 0 Then %>
				  		<input id="timepickerAutoFilterGenSchedulerFriday" type="text" name="txtAutoFilterGenSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketFridayInit" id="txtGenerateFilterTicketFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutoFilterGenSchedulerFriday" type="text" name="txtAutoFilterGenSchedulerFridayTime" value="<%= GenerateFilterTicketFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketFridayInit" id="txtGenerateFilterTicketFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GenerateFilterTicketFriday) = 0 Then %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenFriday" name="chkNoAutoFilterTicketGenFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenFriday" name="chkNoAutoFilterTicketGenFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GenerateFilterTicketSaturday) = 0 Then %>
				  		<input id="timepickerAutoFilterGenSchedulerSaturday" type="text" name="txtAutoFilterGenSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketSaturdayInit" id="txtGenerateFilterTicketSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutoFilterGenSchedulerSaturday" type="text" name="txtAutoFilterGenSchedulerSaturdayTime" value="<%= GenerateFilterTicketSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtGenerateFilterTicketSaturdayInit" id="txtGenerateFilterTicketSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GenerateFilterTicketSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenSaturday" name="chkNoAutoFilterTicketGenSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenSaturday" name="chkNoAutoFilterTicketGenSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunFilterTicketAutoGenIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenIfClosed" name="chkNoAutoFilterTicketGenIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenIfClosed" name="chkNoAutoFilterTicketGenIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunFilterTicketAutoGenIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenIfClosingEarly" name="chkNoAutoFilterTicketGenIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutoFilterTicketGenIfClosingEarly" name="chkNoAutoFilterTicketGenIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
 

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForFieldServiceNotesReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerFieldServiceNotesReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFieldServiceNotesReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFieldServiceNotesReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFieldServiceNotesReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFieldServiceNotesReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFieldServiceNotesReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFieldServiceNotesReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });	
	
			var initGenTimeSunday = $('#txtFieldServiceNotesReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerFieldServiceNotesReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtFieldServiceNotesReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerFieldServiceNotesReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtFieldServiceNotesReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerFieldServiceNotesReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtFieldServiceNotesReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerFieldServiceNotesReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtFieldServiceNotesReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerFieldServiceNotesReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtFieldServiceNotesReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerFieldServiceNotesReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtFieldServiceNotesReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerFieldServiceNotesReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerFieldServiceNotesReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoFieldServiceNotesReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerFieldServiceNotesReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoFieldServiceNotesReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerFieldServiceNotesReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoFieldServiceNotesReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerFieldServiceNotesReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoFieldServiceNotesReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerFieldServiceNotesReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoFieldServiceNotesReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerFieldServiceNotesReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoFieldServiceNotesReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerFieldServiceNotesReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoFieldServiceNotesReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoFieldServiceNotesReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerFieldServiceNotesReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFieldServiceNotesReportSchedulerSunday').timepicker('setTime', '8:00 AM');
			    }
			});
			    	
			$("#chkNoFieldServiceNotesReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerFieldServiceNotesReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFieldServiceNotesReportSchedulerMonday').timepicker('setTime', '8:00 AM');
			    }
			});
	    	
			$("#chkNoFieldServiceNotesReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerFieldServiceNotesReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFieldServiceNotesReportSchedulerTuesday').timepicker('setTime', '8:00 AM');
			    }
			});

			$("#chkNoFieldServiceNotesReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerFieldServiceNotesReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFieldServiceNotesReportSchedulerWednesday').timepicker('setTime', '8:00 AM');
			    }
			});

			$("#chkNoFieldServiceNotesReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerFieldServiceNotesReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFieldServiceNotesReportSchedulerThursday').timepicker('setTime', '8:00 AM');
			    }
			});

			$("#chkNoFieldServiceNotesReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerFieldServiceNotesReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFieldServiceNotesReportSchedulerFriday').timepicker('setTime', '8:00 AM');
			    }
			});

			$("#chkNoFieldServiceNotesReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerFieldServiceNotesReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFieldServiceNotesReportSchedulerSaturday').timepicker('setTime', '8:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing field service notes report generation schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,8:00 AM,8:00 AM,8:00 AM,8:00 AM,8:00 AM,8:00 AM,8:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_FieldServiceNotesReportGeneration = ""
	FieldServiceNotesReportSunday = ""
	FieldServiceNotesReportMonday = ""
	FieldServiceNotesReportTuesday = ""
	FieldServiceNotesReportWednesday = ""
	FieldServiceNotesReportThursday = ""
	FieldServiceNotesReportFriday = ""
	FieldServiceNotesReportSaturday = ""
	FieldServiceNotesReportSundayTime = ""
	FieldServiceNotesReportMondayTime = ""
	FieldServiceNotesReportTuesdayTime = ""
	FieldServiceNotesReportWednesdayTime = ""
	FieldServiceNotesReportThursdayTime = ""
	FieldServiceNotesReportFridayTime = ""
	FieldServiceNotesReportSaturdayTime = ""
	RunFilterTicketAutoGenIfClosed = ""
	RunFilterTicketAutoGenIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_FieldService"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_FieldServiceNotesReportGeneration = rsFieldServiceSettings("Schedule_FieldServiceNotesReportGeneration")
		
		Schedule_FieldServiceNotesReportGenerationSettings = Split(Schedule_FieldServiceNotesReportGeneration,",")

		FieldServiceNotesReportSunday = cInt(Schedule_FieldServiceNotesReportGenerationSettings(0))
		FieldServiceNotesReportMonday = cInt(Schedule_FieldServiceNotesReportGenerationSettings(1))
		FieldServiceNotesReportTuesday = cInt(Schedule_FieldServiceNotesReportGenerationSettings(2))
		FieldServiceNotesReportWednesday = cInt(Schedule_FieldServiceNotesReportGenerationSettings(3))
		FieldServiceNotesReportThursday = cInt(Schedule_FieldServiceNotesReportGenerationSettings(4))
		FieldServiceNotesReportFriday = cInt(Schedule_FieldServiceNotesReportGenerationSettings(5))
		FieldServiceNotesReportSaturday = cInt(Schedule_FieldServiceNotesReportGenerationSettings(6))
		FieldServiceNotesReportSundayTime = Schedule_FieldServiceNotesReportGenerationSettings(7)
		FieldServiceNotesReportMondayTime = Schedule_FieldServiceNotesReportGenerationSettings(8)
		FieldServiceNotesReportTuesdayTime = Schedule_FieldServiceNotesReportGenerationSettings(9)
		FieldServiceNotesReportWednesdayTime = Schedule_FieldServiceNotesReportGenerationSettings(10)
		FieldServiceNotesReportThursdayTime = Schedule_FieldServiceNotesReportGenerationSettings(11)
		FieldServiceNotesReportFridayTime = Schedule_FieldServiceNotesReportGenerationSettings(12)
		FieldServiceNotesReportSaturdayTime = Schedule_FieldServiceNotesReportGenerationSettings(13)
		RunFieldServiceNotesReportIfClosed = cInt(Schedule_FieldServiceNotesReportGenerationSettings(14))
		RunFieldServiceNotesReportIfClosingEarly = cInt(Schedule_FieldServiceNotesReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the <%= GetTerm("field service") %> notes report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the <%= GetTerm("field service") %> notes report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The <%= GetTerm("Field Service") %> Notes Report Can Only Be Generated 6:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FieldServiceNotesReportSunday) = 0 Then %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerSunday" type="text" name="txtFieldServiceNotesReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportSundayInit" id="txtFieldServiceNotesReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerSunday" type="text" name="txtFieldServiceNotesReportSchedulerSundayTime" value="<%= FieldServiceNotesReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportSundayInit" id="txtFieldServiceNotesReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(FieldServiceNotesReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportSunday" name="chkNoFieldServiceNotesReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportSunday" name="chkNoFieldServiceNotesReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FieldServiceNotesReportMonday) = 0 Then %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerMonday" type="text" name="txtFieldServiceNotesReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportMondayInit" id="txtFieldServiceNotesReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerMonday" type="text" name="txtFieldServiceNotesReportSchedulerMondayTime" value="<%= FieldServiceNotesReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportMondayInit" id="txtFieldServiceNotesReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FieldServiceNotesReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportMonday" name="chkNoFieldServiceNotesReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportMonday" name="chkNoFieldServiceNotesReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FieldServiceNotesReportTuesday) = 0 Then %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerTuesday" type="text" value="" name="txtFieldServiceNotesReportSchedulerTuesdayTime" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportTuesdayInit" id="txtFieldServiceNotesReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerTuesday" type="text"  name="txtFieldServiceNotesReportSchedulerTuesdayTime" value="<%= FieldServiceNotesReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportTuesdayInit" id="txtFieldServiceNotesReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FieldServiceNotesReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportTuesday" name="chkNoFieldServiceNotesReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportTuesday" name="chkNoFieldServiceNotesReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FieldServiceNotesReportWednesday) = 0 Then %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerWednesday" type="text" name="txtFieldServiceNotesReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportWednesdayInit" id="txtFieldServiceNotesReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerWednesday" type="text" name="txtFieldServiceNotesReportSchedulerWednesdayTime" value="<%= FieldServiceNotesReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportWednesdayInit" id="txtFieldServiceNotesReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FieldServiceNotesReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportWednesday" name="chkNoFieldServiceNotesReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportWednesday" name="chkNoFieldServiceNotesReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FieldServiceNotesReportThursday) = 0 Then %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerThursday" type="text" name="txtFieldServiceNotesReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportThursdayInit" id="txtFieldServiceNotesReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerThursday" type="text" name="txtFieldServiceNotesReportSchedulerThursdayTime" value="<%= FieldServiceNotesReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportThursdayInit" id="txtFieldServiceNotesReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FieldServiceNotesReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportThursday" name="chkNoFieldServiceNotesReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportThursday" name="chkNoFieldServiceNotesReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FieldServiceNotesReportFriday) = 0 Then %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerFriday" type="text" name="txtFieldServiceNotesReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportFridayInit" id="txtFieldServiceNotesReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerFriday" type="text" name="txtFieldServiceNotesReportSchedulerFridayTime" value="<%= FieldServiceNotesReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportFridayInit" id="txtFieldServiceNotesReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FieldServiceNotesReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportFriday" name="chkNoFieldServiceNotesReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportFriday" name="chkNoFieldServiceNotesReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FieldServiceNotesReportSaturday) = 0 Then %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerSaturday" type="text" name="txtFieldServiceNotesReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportSaturdayInit" id="txtFieldServiceNotesReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFieldServiceNotesReportSchedulerSaturday" type="text" name="txtFieldServiceNotesReportSchedulerSaturdayTime" value="<%= FieldServiceNotesReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtFieldServiceNotesReportSaturdayInit" id="txtFieldServiceNotesReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FieldServiceNotesReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportSaturday" name="chkNoFieldServiceNotesReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportSaturday" name="chkNoFieldServiceNotesReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunFieldServiceNotesReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportIfClosed" name="chkNoFieldServiceNotesReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportIfClosed" name="chkNoFieldServiceNotesReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunFieldServiceNotesReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportIfClosingEarly" name="chkNoFieldServiceNotesReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFieldServiceNotesReportIfClosingEarly" name="chkNoFieldServiceNotesReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************





'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForServiceTicketCarryoverReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerServiceTicketCarryoverReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketCarryoverReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketCarryoverReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketCarryoverReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketCarryoverReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketCarryoverReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketCarryoverReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '3:00 PM',
            	maxTime: '11:45 PM'
	        });

		
			var initGenTimeSunday = $('#txtServiceTicketCarryoverReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerServiceTicketCarryoverReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtServiceTicketCarryoverReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerServiceTicketCarryoverReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtServiceTicketCarryoverReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerServiceTicketCarryoverReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtServiceTicketCarryoverReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerServiceTicketCarryoverReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtServiceTicketCarryoverReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerServiceTicketCarryoverReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtServiceTicketCarryoverReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerServiceTicketCarryoverReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtServiceTicketCarryoverReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerServiceTicketCarryoverReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerServiceTicketCarryoverReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketCarryoverReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketCarryoverReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketCarryoverReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketCarryoverReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketCarryoverReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketCarryoverReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketCarryoverReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketCarryoverReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketCarryoverReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketCarryoverReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketCarryoverReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketCarryoverReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketCarryoverReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoServiceTicketCarryoverReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketCarryoverReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketCarryoverReportSchedulerSunday').timepicker('setTime', '6:00 PM');
			    }
			});
			    	
			$("#chkNoServiceTicketCarryoverReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketCarryoverReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketCarryoverReportSchedulerMonday').timepicker('setTime', '6:00 PM');
			    }
			});
	    	
			$("#chkNoServiceTicketCarryoverReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketCarryoverReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketCarryoverReportSchedulerTuesday').timepicker('setTime', '6:00 PM');
			    }
			});

			$("#chkNoServiceTicketCarryoverReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketCarryoverReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketCarryoverReportSchedulerWednesday').timepicker('setTime', '6:00 PM');
			    }
			});

			$("#chkNoServiceTicketCarryoverReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketCarryoverReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketCarryoverReportSchedulerThursday').timepicker('setTime', '6:00 PM');
			    }
			});

			$("#chkNoServiceTicketCarryoverReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketCarryoverReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketCarryoverReportSchedulerFriday').timepicker('setTime', '6:00 PM');
			    }
			});

			$("#chkNoServiceTicketCarryoverReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketCarryoverReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketCarryoverReportSchedulerSaturday').timepicker('setTime', '6:00 PM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing service ticket carryover report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,1,1,1,1,1,0,6:00 PM,6:00 PM,6:00 PM,6:00 PM,6:00 PM,6:00 PM,6:00 PM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_ServiceTicketCarryoverReportGeneration = ""
	ServiceTicketCarryoverReportSunday = ""
	ServiceTicketCarryoverReportMonday = ""
	ServiceTicketCarryoverReportTuesday = ""
	ServiceTicketCarryoverReportWednesday = ""
	ServiceTicketCarryoverReportThursday = ""
	ServiceTicketCarryoverReportFriday = ""
	ServiceTicketCarryoverReportSaturday = ""
	ServiceTicketCarryoverReportSundayTime = ""
	ServiceTicketCarryoverReportMondayTime = ""
	ServiceTicketCarryoverReportTuesdayTime = ""
	ServiceTicketCarryoverReportWednesdayTime = ""
	ServiceTicketCarryoverReportThursdayTime = ""
	ServiceTicketCarryoverReportFridayTime = ""
	ServiceTicketCarryoverReportSaturdayTime = ""
	RunServiceTicketCarryoverReportIfClosed = ""
	RunServiceTicketCarryoverReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_FieldService"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_ServiceTicketCarryoverReportGeneration = rsFieldServiceSettings("Schedule_ServiceTicketCarryoverReportGeneration")
		
		Schedule_ServiceTicketCarryoverReportGenerationSettings = Split(Schedule_ServiceTicketCarryoverReportGeneration,",")

		ServiceTicketCarryoverReportSunday = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(0))
		ServiceTicketCarryoverReportMonday = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(1))
		ServiceTicketCarryoverReportTuesday = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(2))
		ServiceTicketCarryoverReportWednesday = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(3))
		ServiceTicketCarryoverReportThursday = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(4))
		ServiceTicketCarryoverReportFriday = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(5))
		ServiceTicketCarryoverReportSaturday = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(6))
		ServiceTicketCarryoverReportSundayTime = Schedule_ServiceTicketCarryoverReportGenerationSettings(7)
		ServiceTicketCarryoverReportMondayTime = Schedule_ServiceTicketCarryoverReportGenerationSettings(8)
		ServiceTicketCarryoverReportTuesdayTime = Schedule_ServiceTicketCarryoverReportGenerationSettings(9)
		ServiceTicketCarryoverReportWednesdayTime = Schedule_ServiceTicketCarryoverReportGenerationSettings(10)
		ServiceTicketCarryoverReportThursdayTime = Schedule_ServiceTicketCarryoverReportGenerationSettings(11)
		ServiceTicketCarryoverReportFridayTime = Schedule_ServiceTicketCarryoverReportGenerationSettings(12)
		ServiceTicketCarryoverReportSaturdayTime = Schedule_ServiceTicketCarryoverReportGenerationSettings(13)
		RunServiceTicketCarryoverReportIfClosed = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(14))
		RunServiceTicketCarryoverReportIfClosingEarly = cInt(Schedule_ServiceTicketCarryoverReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the service ticket carryover report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the service ticket carryover report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The Service Ticket Carryover Report Can Only Be Generated 3:00 PM - 11:45 PM each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketCarryoverReportSunday) = 0 Then %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerSunday" type="text" name="txtServiceTicketCarryoverReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportSundayInit" id="txtServiceTicketCarryoverReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerSunday" type="text" name="txtServiceTicketCarryoverReportSchedulerSundayTime" value="<%= ServiceTicketCarryoverReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportSundayInit" id="txtServiceTicketCarryoverReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(ServiceTicketCarryoverReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportSunday" name="chkNoServiceTicketCarryoverReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportSunday" name="chkNoServiceTicketCarryoverReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketCarryoverReportMonday) = 0 Then %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerMonday" type="text" name="txtServiceTicketCarryoverReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportMondayInit" id="txtServiceTicketCarryoverReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerMonday" type="text" name="txtServiceTicketCarryoverReportSchedulerMondayTime" value="<%= ServiceTicketCarryoverReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportMondayInit" id="txtServiceTicketCarryoverReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketCarryoverReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportMonday" name="chkNoServiceTicketCarryoverReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportMonday" name="chkNoServiceTicketCarryoverReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketCarryoverReportTuesday) = 0 Then %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerTuesday" type="text" name="txtServiceTicketCarryoverReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportTuesdayInit" id="txtServiceTicketCarryoverReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerTuesday" type="text" name="txtServiceTicketCarryoverReportSchedulerTuesdayTime" value="<%= ServiceTicketCarryoverReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportTuesdayInit" id="txtServiceTicketCarryoverReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketCarryoverReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportTuesday" name="chkNoServiceTicketCarryoverReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportTuesday" name="chkNoServiceTicketCarryoverReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketCarryoverReportWednesday) = 0 Then %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerWednesday" type="text" name="txtServiceTicketCarryoverReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportWednesdayInit" id="txtServiceTicketCarryoverReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerWednesday" type="text" name="txtServiceTicketCarryoverReportSchedulerWednesdayTime" value="<%= ServiceTicketCarryoverReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportWednesdayInit" id="txtServiceTicketCarryoverReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketCarryoverReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportWednesday" name="chkNoServiceTicketCarryoverReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportWednesday" name="chkNoServiceTicketCarryoverReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketCarryoverReportThursday) = 0 Then %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerThursday" type="text" name="txtServiceTicketCarryoverReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportThursdayInit" id="txtServiceTicketCarryoverReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerThursday" type="text" name="txtServiceTicketCarryoverReportSchedulerThursdayTime" value="<%= ServiceTicketCarryoverReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportThursdayInit" id="txtServiceTicketCarryoverReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketCarryoverReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportThursday" name="chkNoServiceTicketCarryoverReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportThursday" name="chkNoServiceTicketCarryoverReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketCarryoverReportFriday) = 0 Then %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerFriday" type="text" name="txtServiceTicketCarryoverReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportFridayInit" id="txtServiceTicketCarryoverReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerFriday" type="text" name="txtServiceTicketCarryoverReportSchedulerFridayTime" value="<%= ServiceTicketCarryoverReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportFridayInit" id="txtServiceTicketCarryoverReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketCarryoverReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportFriday" name="chkNoServiceTicketCarryoverReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportFriday" name="chkNoServiceTicketCarryoverReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketCarryoverReportSaturday) = 0 Then %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerSaturday" type="text" name="txtServiceTicketCarryoverReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportSaturdayInit" id="txtServiceTicketCarryoverReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketCarryoverReportSchedulerSaturday" type="text" name="txtServiceTicketCarryoverReportSchedulerSaturdayTime" value="<%= ServiceTicketCarryoverReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketCarryoverReportSaturdayInit" id="txtServiceTicketCarryoverReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketCarryoverReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportSaturday" name="chkNoServiceTicketCarryoverReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportSaturday" name="chkNoServiceTicketCarryoverReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunServiceTicketCarryoverReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportIfClosed" name="chkNoServiceTicketCarryoverReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportIfClosed" name="chkNoServiceTicketCarryoverReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunServiceTicketCarryoverReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportIfClosingEarly" name="chkNoServiceTicketCarryoverReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketCarryoverReportIfClosingEarly" name="chkNoServiceTicketCarryoverReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForServiceTicketThresholdReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerServiceTicketThresholdReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketThresholdReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketThresholdReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketThresholdReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketThresholdReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketThresholdReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '11:45 PM'
	        });
	        $('#timepickerServiceTicketThresholdReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '11:45 PM'
	        });

		
			var initGenTimeSunday = $('#txtServiceTicketThresholdReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerServiceTicketThresholdReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtServiceTicketThresholdReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerServiceTicketThresholdReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtServiceTicketThresholdReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerServiceTicketThresholdReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtServiceTicketThresholdReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerServiceTicketThresholdReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtServiceTicketThresholdReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerServiceTicketThresholdReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtServiceTicketThresholdReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerServiceTicketThresholdReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtServiceTicketThresholdReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerServiceTicketThresholdReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerServiceTicketThresholdReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketThresholdReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketThresholdReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketThresholdReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketThresholdReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketThresholdReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketThresholdReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketThresholdReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketThresholdReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketThresholdReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketThresholdReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketThresholdReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerServiceTicketThresholdReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoServiceTicketThresholdReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoServiceTicketThresholdReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketThresholdReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketThresholdReportSchedulerSunday').timepicker('setTime', '6:00 PM');
			    }
			});
			    	
			$("#chkNoServiceTicketThresholdReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketThresholdReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketThresholdReportSchedulerMonday').timepicker('setTime', '6:00 PM');
			    }
			});
	    	
			$("#chkNoServiceTicketThresholdReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketThresholdReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketThresholdReportSchedulerTuesday').timepicker('setTime', '6:00 PM');
			    }
			});

			$("#chkNoServiceTicketThresholdReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketThresholdReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketThresholdReportSchedulerWednesday').timepicker('setTime', '6:00 PM');
			    }
			});

			$("#chkNoServiceTicketThresholdReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketThresholdReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketThresholdReportSchedulerThursday').timepicker('setTime', '6:00 PM');
			    }
			});

			$("#chkNoServiceTicketThresholdReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketThresholdReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketThresholdReportSchedulerFriday').timepicker('setTime', '6:00 PM');
			    }
			});

			$("#chkNoServiceTicketThresholdReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerServiceTicketThresholdReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerServiceTicketThresholdReportSchedulerSaturday').timepicker('setTime', '6:00 PM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing service ticket Threshold report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,1,1,1,1,1,0,8:30 AM,8:30 AM,8:30 AM,8:30 AM,8:30 AM,8:30 AM,8:30 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_ServiceTicketThresholdReportGeneration = ""
	ServiceTicketThresholdReportSunday = ""
	ServiceTicketThresholdReportMonday = ""
	ServiceTicketThresholdReportTuesday = ""
	ServiceTicketThresholdReportWednesday = ""
	ServiceTicketThresholdReportThursday = ""
	ServiceTicketThresholdReportFriday = ""
	ServiceTicketThresholdReportSaturday = ""
	ServiceTicketThresholdReportSundayTime = ""
	ServiceTicketThresholdReportMondayTime = ""
	ServiceTicketThresholdReportTuesdayTime = ""
	ServiceTicketThresholdReportWednesdayTime = ""
	ServiceTicketThresholdReportThursdayTime = ""
	ServiceTicketThresholdReportFridayTime = ""
	ServiceTicketThresholdReportSaturdayTime = ""
	RunServiceTicketThresholdReportIfClosed = ""
	RunServiceTicketThresholdReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_FieldService"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_ServiceTicketThresholdReportGeneration = rsFieldServiceSettings("Schedule_ServiceTicketThresholdReportGeneration")
		
		Schedule_ServiceTicketThresholdReportGenerationSettings = Split(Schedule_ServiceTicketThresholdReportGeneration,",")

		ServiceTicketThresholdReportSunday = cInt(Schedule_ServiceTicketThresholdReportGenerationSettings(0))
		ServiceTicketThresholdReportMonday = cInt(Schedule_ServiceTicketThresholdReportGenerationSettings(1))
		ServiceTicketThresholdReportTuesday = cInt(Schedule_ServiceTicketThresholdReportGenerationSettings(2))
		ServiceTicketThresholdReportWednesday = cInt(Schedule_ServiceTicketThresholdReportGenerationSettings(3))
		ServiceTicketThresholdReportThursday = cInt(Schedule_ServiceTicketThresholdReportGenerationSettings(4))
		ServiceTicketThresholdReportFriday = cInt(Schedule_ServiceTicketThresholdReportGenerationSettings(5))
		ServiceTicketThresholdReportSaturday = cInt(Schedule_ServiceTicketThresholdReportGenerationSettings(6))
		ServiceTicketThresholdReportSundayTime = Schedule_ServiceTicketThresholdReportGenerationSettings(7)
		ServiceTicketThresholdReportMondayTime = Schedule_ServiceTicketThresholdReportGenerationSettings(8)
		ServiceTicketThresholdReportTuesdayTime = Schedule_ServiceTicketThresholdReportGenerationSettings(9)
		ServiceTicketThresholdReportWednesdayTime = Schedule_ServiceTicketThresholdReportGenerationSettings(10)
		ServiceTicketThresholdReportThursdayTime = Schedule_ServiceTicketThresholdReportGenerationSettings(11)
		ServiceTicketThresholdReportFridayTime = Schedule_ServiceTicketThresholdReportGenerationSettings(12)
		ServiceTicketThresholdReportSaturdayTime = Schedule_ServiceTicketThresholdReportGenerationSettings(13)
		RunServiceTicketThresholdReportIfClosed = cInt(Schedule_ServiceTicketThresholdReportGenerationSettings(14))
		RunServiceTicketThresholdReportIfClosingEarly = cInt(Schedule_ServiceTicketThresholdReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the service ticket Threshold report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the service ticket Threshold report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The Service Ticket Threshold Report Can Only Be Generated 6:00 AM - 11:45 PM each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketThresholdReportSunday) = 0 Then %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerSunday" type="text" name="txtServiceTicketThresholdReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportSundayInit" id="txtServiceTicketThresholdReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerSunday" type="text" name="txtServiceTicketThresholdReportSchedulerSundayTime" value="<%= ServiceTicketThresholdReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportSundayInit" id="txtServiceTicketThresholdReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(ServiceTicketThresholdReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportSunday" name="chkNoServiceTicketThresholdReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportSunday" name="chkNoServiceTicketThresholdReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketThresholdReportMonday) = 0 Then %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerMonday" type="text" name="txtServiceTicketThresholdReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportMondayInit" id="txtServiceTicketThresholdReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerMonday" type="text" name="txtServiceTicketThresholdReportSchedulerMondayTime" value="<%= ServiceTicketThresholdReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportMondayInit" id="txtServiceTicketThresholdReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketThresholdReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportMonday" name="chkNoServiceTicketThresholdReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportMonday" name="chkNoServiceTicketThresholdReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketThresholdReportTuesday) = 0 Then %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerTuesday" type="text" name="txtServiceTicketThresholdReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportTuesdayInit" id="txtServiceTicketThresholdReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerTuesday" type="text" name="txtServiceTicketThresholdReportSchedulerTuesdayTime" value="<%= ServiceTicketThresholdReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportTuesdayInit" id="txtServiceTicketThresholdReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketThresholdReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportTuesday" name="chkNoServiceTicketThresholdReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportTuesday" name="chkNoServiceTicketThresholdReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketThresholdReportWednesday) = 0 Then %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerWednesday" type="text" name="txtServiceTicketThresholdReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportWednesdayInit" id="txtServiceTicketThresholdReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerWednesday" type="text" name="txtServiceTicketThresholdReportSchedulerWednesdayTime" value="<%= ServiceTicketThresholdReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportWednesdayInit" id="txtServiceTicketThresholdReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketThresholdReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportWednesday" name="chkNoServiceTicketThresholdReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportWednesday" name="chkNoServiceTicketThresholdReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketThresholdReportThursday) = 0 Then %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerThursday" type="text" name="txtServiceTicketThresholdReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportThursdayInit" id="txtServiceTicketThresholdReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerThursday" type="text" name="txtServiceTicketThresholdReportSchedulerThursdayTime" value="<%= ServiceTicketThresholdReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportThursdayInit" id="txtServiceTicketThresholdReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketThresholdReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportThursday" name="chkNoServiceTicketThresholdReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportThursday" name="chkNoServiceTicketThresholdReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketThresholdReportFriday) = 0 Then %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerFriday" type="text" name="txtServiceTicketThresholdReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportFridayInit" id="txtServiceTicketThresholdReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerFriday" type="text" name="txtServiceTicketThresholdReportSchedulerFridayTime" value="<%= ServiceTicketThresholdReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportFridayInit" id="txtServiceTicketThresholdReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketThresholdReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportFriday" name="chkNoServiceTicketThresholdReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportFriday" name="chkNoServiceTicketThresholdReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ServiceTicketThresholdReportSaturday) = 0 Then %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerSaturday" type="text" name="txtServiceTicketThresholdReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportSaturdayInit" id="txtServiceTicketThresholdReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerServiceTicketThresholdReportSchedulerSaturday" type="text" name="txtServiceTicketThresholdReportSchedulerSaturdayTime" value="<%= ServiceTicketThresholdReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtServiceTicketThresholdReportSaturdayInit" id="txtServiceTicketThresholdReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ServiceTicketThresholdReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportSaturday" name="chkNoServiceTicketThresholdReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportSaturday" name="chkNoServiceTicketThresholdReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunServiceTicketThresholdReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportIfClosed" name="chkNoServiceTicketThresholdReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportIfClosed" name="chkNoServiceTicketThresholdReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunServiceTicketThresholdReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportIfClosingEarly" name="chkNoServiceTicketThresholdReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoServiceTicketThresholdReportIfClosingEarly" name="chkNoServiceTicketThresholdReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForProspectingSnapshotReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerProspectingSnapshotReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingSnapshotReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingSnapshotReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingSnapshotReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingSnapshotReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingSnapshotReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingSnapshotReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtProspectingSnapshotReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerProspectingSnapshotReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtProspectingSnapshotReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerProspectingSnapshotReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtProspectingSnapshotReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerProspectingSnapshotReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtProspectingSnapshotReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerProspectingSnapshotReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtProspectingSnapshotReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerProspectingSnapshotReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtProspectingSnapshotReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerProspectingSnapshotReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtProspectingSnapshotReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerProspectingSnapshotReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerProspectingSnapshotReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingSnapshotReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingSnapshotReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingSnapshotReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingSnapshotReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingSnapshotReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingSnapshotReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingSnapshotReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingSnapshotReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingSnapshotReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingSnapshotReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingSnapshotReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingSnapshotReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingSnapshotReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoProspectingSnapshotReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingSnapshotReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingSnapshotReportSchedulerSunday').timepicker('setTime', '6:00 AM');
			    }
			});
			    	
			$("#chkNoProspectingSnapshotReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingSnapshotReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingSnapshotReportSchedulerMonday').timepicker('setTime', '6:00 AM');
			    }
			});
	    	
			$("#chkNoProspectingSnapshotReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingSnapshotReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingSnapshotReportSchedulerTuesday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoProspectingSnapshotReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingSnapshotReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingSnapshotReportSchedulerWednesday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoProspectingSnapshotReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingSnapshotReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingSnapshotReportSchedulerThursday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoProspectingSnapshotReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingSnapshotReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingSnapshotReportSchedulerFriday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoProspectingSnapshotReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingSnapshotReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingSnapshotReportSchedulerSaturday').timepicker('setTime', '6:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing prospecting snapshot report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_ProspectingSnapshotReportGeneration = ""
	ProspectingSnapshotReportSunday = ""
	ProspectingSnapshotReportMonday = ""
	ProspectingSnapshotReportTuesday = ""
	ProspectingSnapshotReportWednesday = ""
	ProspectingSnapshotReportThursday = ""
	ProspectingSnapshotReportFriday = ""
	ProspectingSnapshotReportSaturday = ""
	ProspectingSnapshotReportSundayTime = ""
	ProspectingSnapshotReportMondayTime = ""
	ProspectingSnapshotReportTuesdayTime = ""
	ProspectingSnapshotReportWednesdayTime = ""
	ProspectingSnapshotReportThursdayTime = ""
	ProspectingSnapshotReportFridayTime = ""
	ProspectingSnapshotReportSaturdayTime = ""
	RunProspectingSnapshotReportIfClosed = ""
	RunProspectingSnapshotReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_Prospecting"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_ProspectingSnapshotReportGeneration = rsFieldServiceSettings("Schedule_ProspectingSnapshotReportGeneration")
		
		Schedule_ProspectingSnapshotReportGenerationSettings = Split(Schedule_ProspectingSnapshotReportGeneration,",")

		ProspectingSnapshotReportSunday = cInt(Schedule_ProspectingSnapshotReportGenerationSettings(0))
		ProspectingSnapshotReportMonday = cInt(Schedule_ProspectingSnapshotReportGenerationSettings(1))
		ProspectingSnapshotReportTuesday = cInt(Schedule_ProspectingSnapshotReportGenerationSettings(2))
		ProspectingSnapshotReportWednesday = cInt(Schedule_ProspectingSnapshotReportGenerationSettings(3))
		ProspectingSnapshotReportThursday = cInt(Schedule_ProspectingSnapshotReportGenerationSettings(4))
		ProspectingSnapshotReportFriday = cInt(Schedule_ProspectingSnapshotReportGenerationSettings(5))
		ProspectingSnapshotReportSaturday = cInt(Schedule_ProspectingSnapshotReportGenerationSettings(6))
		ProspectingSnapshotReportSundayTime = Schedule_ProspectingSnapshotReportGenerationSettings(7)
		ProspectingSnapshotReportMondayTime = Schedule_ProspectingSnapshotReportGenerationSettings(8)
		ProspectingSnapshotReportTuesdayTime = Schedule_ProspectingSnapshotReportGenerationSettings(9)
		ProspectingSnapshotReportWednesdayTime = Schedule_ProspectingSnapshotReportGenerationSettings(10)
		ProspectingSnapshotReportThursdayTime = Schedule_ProspectingSnapshotReportGenerationSettings(11)
		ProspectingSnapshotReportFridayTime = Schedule_ProspectingSnapshotReportGenerationSettings(12)
		ProspectingSnapshotReportSaturdayTime = Schedule_ProspectingSnapshotReportGenerationSettings(13)
		RunProspectingSnapshotReportIfClosed = cInt(Schedule_ProspectingSnapshotReportGenerationSettings(14))
		RunProspectingSnapshotReportIfClosingEarly = cInt(Schedule_ProspectingSnapshotReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the <%= GetTerm("prospecting") %> snapshot report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the <%= GetTerm("prospecting") %> snapshot report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The <%= GetTerm("Prospecting") %> Snapshot Report Can Only Be Generated 6:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingSnapshotReportSunday) = 0 Then %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerSunday" type="text" name="txtProspectingSnapshotReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportSundayInit" id="txtProspectingSnapshotReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerSunday" type="text" name="txtProspectingSnapshotReportSchedulerSundayTime" value="<%= ProspectingSnapshotReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportSundayInit" id="txtProspectingSnapshotReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(ProspectingSnapshotReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportSunday" name="chkNoProspectingSnapshotReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportSunday" name="chkNoProspectingSnapshotReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingSnapshotReportMonday) = 0 Then %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerMonday" type="text" name="txtProspectingSnapshotReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportMondayInit" id="txtProspectingSnapshotReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerMonday" type="text" name="txtProspectingSnapshotReportSchedulerMondayTime" value="<%= ProspectingSnapshotReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportMondayInit" id="txtProspectingSnapshotReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingSnapshotReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportMonday" name="chkNoProspectingSnapshotReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportMonday" name="chkNoProspectingSnapshotReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingSnapshotReportTuesday) = 0 Then %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerTuesday" type="text" name="txtProspectingSnapshotReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportTuesdayInit" id="txtProspectingSnapshotReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerTuesday" type="text" name="txtProspectingSnapshotReportSchedulerTuesdayTime" value="<%= ProspectingSnapshotReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportTuesdayInit" id="txtProspectingSnapshotReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingSnapshotReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportTuesday" name="chkNoProspectingSnapshotReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportTuesday" name="chkNoProspectingSnapshotReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingSnapshotReportWednesday) = 0 Then %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerWednesday" type="text" name="txtProspectingSnapshotReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportWednesdayInit" id="txtProspectingSnapshotReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerWednesday" type="text" name="txtProspectingSnapshotReportSchedulerWednesdayTime" value="<%= ProspectingSnapshotReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportWednesdayInit" id="txtProspectingSnapshotReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingSnapshotReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportWednesday" name="chkNoProspectingSnapshotReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportWednesday" name="chkNoProspectingSnapshotReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingSnapshotReportThursday) = 0 Then %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerThursday" type="text" name="txtProspectingSnapshotReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportThursdayInit" id="txtProspectingSnapshotReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerThursday" type="text" name="txtProspectingSnapshotReportSchedulerThursdayTime" value="<%= ProspectingSnapshotReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportThursdayInit" id="txtProspectingSnapshotReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingSnapshotReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportThursday" name="chkNoProspectingSnapshotReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportThursday" name="chkNoProspectingSnapshotReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingSnapshotReportFriday) = 0 Then %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerFriday" type="text" name="txtProspectingSnapshotReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportFridayInit" id="txtProspectingSnapshotReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerFriday" type="text" name="txtProspectingSnapshotReportSchedulerFridayTime" value="<%= ProspectingSnapshotReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportFridayInit" id="txtProspectingSnapshotReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingSnapshotReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportFriday" name="chkNoProspectingSnapshotReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportFriday" name="chkNoProspectingSnapshotReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingSnapshotReportSaturday) = 0 Then %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerSaturday" type="text" name="txtProspectingSnapshotReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportSaturdayInit" id="txtProspectingSnapshotReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingSnapshotReportSchedulerSaturday" type="text" name="txtProspectingSnapshotReportSchedulerSaturdayTime" value="<%= ProspectingSnapshotReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingSnapshotReportSaturdayInit" id="txtProspectingSnapshotReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingSnapshotReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportSaturday" name="chkNoProspectingSnapshotReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportSaturday" name="chkNoProspectingSnapshotReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunProspectingSnapshotReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportIfClosed" name="chkNoProspectingSnapshotReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportIfClosed" name="chkNoProspectingSnapshotReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunProspectingSnapshotReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportIfClosingEarly" name="chkNoProspectingSnapshotReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingSnapshotReportIfClosingEarly" name="chkNoProspectingSnapshotReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForProspectingWeeklyAgendaReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerProspectingWeeklyAgendaReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingWeeklyAgendaReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingWeeklyAgendaReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingWeeklyAgendaReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingWeeklyAgendaReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingWeeklyAgendaReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerProspectingWeeklyAgendaReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtProspectingWeeklyAgendaReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerProspectingWeeklyAgendaReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtProspectingWeeklyAgendaReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerProspectingWeeklyAgendaReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtProspectingWeeklyAgendaReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerProspectingWeeklyAgendaReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtProspectingWeeklyAgendaReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerProspectingWeeklyAgendaReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtProspectingWeeklyAgendaReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerProspectingWeeklyAgendaReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtProspectingWeeklyAgendaReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerProspectingWeeklyAgendaReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtProspectingWeeklyAgendaReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerProspectingWeeklyAgendaReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerProspectingWeeklyAgendaReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingWeeklyAgendaReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingWeeklyAgendaReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingWeeklyAgendaReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingWeeklyAgendaReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingWeeklyAgendaReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingWeeklyAgendaReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingWeeklyAgendaReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingWeeklyAgendaReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingWeeklyAgendaReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingWeeklyAgendaReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingWeeklyAgendaReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerProspectingWeeklyAgendaReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoProspectingWeeklyAgendaReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoProspectingWeeklyAgendaReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingWeeklyAgendaReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingWeeklyAgendaReportSchedulerSunday').timepicker('setTime', '6:00 AM');
			    }
			});
			    	
			$("#chkNoProspectingWeeklyAgendaReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingWeeklyAgendaReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingWeeklyAgendaReportSchedulerMonday').timepicker('setTime', '6:00 AM');
			    }
			});
	    	
			$("#chkNoProspectingWeeklyAgendaReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingWeeklyAgendaReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingWeeklyAgendaReportSchedulerTuesday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoProspectingWeeklyAgendaReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingWeeklyAgendaReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingWeeklyAgendaReportSchedulerWednesday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoProspectingWeeklyAgendaReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingWeeklyAgendaReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingWeeklyAgendaReportSchedulerThursday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoProspectingWeeklyAgendaReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingWeeklyAgendaReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingWeeklyAgendaReportSchedulerFriday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoProspectingWeeklyAgendaReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerProspectingWeeklyAgendaReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerProspectingWeeklyAgendaReportSchedulerSaturday').timepicker('setTime', '6:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing prospecting weekly agenda report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,1,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_ProspectingWeeklyAgendaReportGeneration = ""
	ProspectingWeeklyAgendaReportSunday = ""
	ProspectingWeeklyAgendaReportMonday = ""
	ProspectingWeeklyAgendaReportTuesday = ""
	ProspectingWeeklyAgendaReportWednesday = ""
	ProspectingWeeklyAgendaReportThursday = ""
	ProspectingWeeklyAgendaReportFriday = ""
	ProspectingWeeklyAgendaReportSaturday = ""
	ProspectingWeeklyAgendaReportSundayTime = ""
	ProspectingWeeklyAgendaReportMondayTime = ""
	ProspectingWeeklyAgendaReportTuesdayTime = ""
	ProspectingWeeklyAgendaReportWednesdayTime = ""
	ProspectingWeeklyAgendaReportThursdayTime = ""
	ProspectingWeeklyAgendaReportFridayTime = ""
	ProspectingWeeklyAgendaReportSaturdayTime = ""
	RunProspectingWeeklyAgendaReportIfClosed = ""
	RunProspectingWeeklyAgendaReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_Prospecting"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_ProspectingWeeklyAgendaReportGeneration = rsFieldServiceSettings("Schedule_ProspectingWeeklyAgendaReportGeneration")
		
		Schedule_ProspectingWeeklyAgendaReportGenerationSettings = Split(Schedule_ProspectingWeeklyAgendaReportGeneration,",")

		ProspectingWeeklyAgendaReportSunday = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(0))
		ProspectingWeeklyAgendaReportMonday = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(1))
		ProspectingWeeklyAgendaReportTuesday = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(2))
		ProspectingWeeklyAgendaReportWednesday = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(3))
		ProspectingWeeklyAgendaReportThursday = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(4))
		ProspectingWeeklyAgendaReportFriday = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(5))
		ProspectingWeeklyAgendaReportSaturday = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(6))
		ProspectingWeeklyAgendaReportSundayTime = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(7)
		ProspectingWeeklyAgendaReportMondayTime = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(8)
		ProspectingWeeklyAgendaReportTuesdayTime = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(9)
		ProspectingWeeklyAgendaReportWednesdayTime = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(10)
		ProspectingWeeklyAgendaReportThursdayTime = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(11)
		ProspectingWeeklyAgendaReportFridayTime = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(12)
		ProspectingWeeklyAgendaReportSaturdayTime = Schedule_ProspectingWeeklyAgendaReportGenerationSettings(13)
		RunProspectingWeeklyAgendaReportIfClosed = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(14))
		RunProspectingWeeklyAgendaReportIfClosingEarly = cInt(Schedule_ProspectingWeeklyAgendaReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the <%= GetTerm("prospecting") %> weekly agenda report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the <%= GetTerm("prospecting") %> weekly agenda report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The <%= GetTerm("Prospecting") %> Weekly Agenda Report Can Only Be Generated 6:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingWeeklyAgendaReportSunday) = 0 Then %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerSunday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportSundayInit" id="txtProspectingWeeklyAgendaReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerSunday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerSundayTime" value="<%= ProspectingWeeklyAgendaReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportSundayInit" id="txtProspectingWeeklyAgendaReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(ProspectingWeeklyAgendaReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportSunday" name="chkNoProspectingWeeklyAgendaReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportSunday" name="chkNoProspectingWeeklyAgendaReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingWeeklyAgendaReportMonday) = 0 Then %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerMonday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportMondayInit" id="txtProspectingWeeklyAgendaReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerMonday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerMondayTime" value="<%= ProspectingWeeklyAgendaReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportMondayInit" id="txtProspectingWeeklyAgendaReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingWeeklyAgendaReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportMonday" name="chkNoProspectingWeeklyAgendaReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportMonday" name="chkNoProspectingWeeklyAgendaReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingWeeklyAgendaReportTuesday) = 0 Then %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerTuesday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportTuesdayInit" id="txtProspectingWeeklyAgendaReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerTuesday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerTuesdayTime" value="<%= ProspectingWeeklyAgendaReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportTuesdayInit" id="txtProspectingWeeklyAgendaReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingWeeklyAgendaReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportTuesday" name="chkNoProspectingWeeklyAgendaReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportTuesday" name="chkNoProspectingWeeklyAgendaReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingWeeklyAgendaReportWednesday) = 0 Then %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerWednesday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportWednesdayInit" id="txtProspectingWeeklyAgendaReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerWednesday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerWednesdayTime" value="<%= ProspectingWeeklyAgendaReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportWednesdayInit" id="txtProspectingWeeklyAgendaReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingWeeklyAgendaReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportWednesday" name="chkNoProspectingWeeklyAgendaReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportWednesday" name="chkNoProspectingWeeklyAgendaReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingWeeklyAgendaReportThursday) = 0 Then %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerThursday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportThursdayInit" id="txtProspectingWeeklyAgendaReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerThursday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerThursdayTime" value="<%= ProspectingWeeklyAgendaReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportThursdayInit" id="txtProspectingWeeklyAgendaReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingWeeklyAgendaReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportThursday" name="chkNoProspectingWeeklyAgendaReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportThursday" name="chkNoProspectingWeeklyAgendaReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingWeeklyAgendaReportFriday) = 0 Then %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerFriday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportFridayInit" id="txtProspectingWeeklyAgendaReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerFriday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerFridayTime" value="<%= ProspectingWeeklyAgendaReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportFridayInit" id="txtProspectingWeeklyAgendaReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingWeeklyAgendaReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportFriday" name="chkNoProspectingWeeklyAgendaReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportFriday" name="chkNoProspectingWeeklyAgendaReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(ProspectingWeeklyAgendaReportSaturday) = 0 Then %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerSaturday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportSaturdayInit" id="txtProspectingWeeklyAgendaReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerProspectingWeeklyAgendaReportSchedulerSaturday" type="text" name="txtProspectingWeeklyAgendaReportSchedulerSaturdayTime" value="<%= ProspectingWeeklyAgendaReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtProspectingWeeklyAgendaReportSaturdayInit" id="txtProspectingWeeklyAgendaReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(ProspectingWeeklyAgendaReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportSaturday" name="chkNoProspectingWeeklyAgendaReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportSaturday" name="chkNoProspectingWeeklyAgendaReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunProspectingWeeklyAgendaReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportIfClosed" name="chkNoProspectingWeeklyAgendaReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportIfClosed" name="chkNoProspectingWeeklyAgendaReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunProspectingWeeklyAgendaReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportIfClosingEarly" name="chkNoProspectingWeeklyAgendaReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoProspectingWeeklyAgendaReportIfClosingEarly" name="chkNoProspectingWeeklyAgendaReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************







'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForDailyAPIActivityByPartnerReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerDailyAPIActivityByPartnerReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyAPIActivityByPartnerReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyAPIActivityByPartnerReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyAPIActivityByPartnerReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyAPIActivityByPartnerReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyAPIActivityByPartnerReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyAPIActivityByPartnerReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtDailyAPIActivityByPartnerReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerDailyAPIActivityByPartnerReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtDailyAPIActivityByPartnerReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerDailyAPIActivityByPartnerReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtDailyAPIActivityByPartnerReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerDailyAPIActivityByPartnerReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtDailyAPIActivityByPartnerReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerDailyAPIActivityByPartnerReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtDailyAPIActivityByPartnerReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerDailyAPIActivityByPartnerReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtDailyAPIActivityByPartnerReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerDailyAPIActivityByPartnerReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtDailyAPIActivityByPartnerReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerDailyAPIActivityByPartnerReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerDailyAPIActivityByPartnerReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyAPIActivityByPartnerReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyAPIActivityByPartnerReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyAPIActivityByPartnerReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyAPIActivityByPartnerReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyAPIActivityByPartnerReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyAPIActivityByPartnerReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyAPIActivityByPartnerReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyAPIActivityByPartnerReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyAPIActivityByPartnerReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyAPIActivityByPartnerReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyAPIActivityByPartnerReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyAPIActivityByPartnerReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyAPIActivityByPartnerReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoDailyAPIActivityByPartnerReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyAPIActivityByPartnerReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyAPIActivityByPartnerReportSchedulerSunday').timepicker('setTime', '6:00 AM');
			    }
			});
			    	
			$("#chkNoDailyAPIActivityByPartnerReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyAPIActivityByPartnerReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyAPIActivityByPartnerReportSchedulerMonday').timepicker('setTime', '6:00 AM');
			    }
			});
	    	
			$("#chkNoDailyAPIActivityByPartnerReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyAPIActivityByPartnerReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyAPIActivityByPartnerReportSchedulerTuesday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoDailyAPIActivityByPartnerReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyAPIActivityByPartnerReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyAPIActivityByPartnerReportSchedulerWednesday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoDailyAPIActivityByPartnerReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyAPIActivityByPartnerReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyAPIActivityByPartnerReportSchedulerThursday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoDailyAPIActivityByPartnerReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyAPIActivityByPartnerReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyAPIActivityByPartnerReportSchedulerFriday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoDailyAPIActivityByPartnerReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyAPIActivityByPartnerReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyAPIActivityByPartnerReportSchedulerSaturday').timepicker('setTime', '6:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing daily api activity by partner report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_DailyAPIActivityByPartnerReportGeneration = ""
	DailyAPIActivityByPartnerReportSunday = ""
	DailyAPIActivityByPartnerReportMonday = ""
	DailyAPIActivityByPartnerReportTuesday = ""
	DailyAPIActivityByPartnerReportWednesday = ""
	DailyAPIActivityByPartnerReportThursday = ""
	DailyAPIActivityByPartnerReportFriday = ""
	DailyAPIActivityByPartnerReportSaturday = ""
	DailyAPIActivityByPartnerReportSundayTime = ""
	DailyAPIActivityByPartnerReportMondayTime = ""
	DailyAPIActivityByPartnerReportTuesdayTime = ""
	DailyAPIActivityByPartnerReportWednesdayTime = ""
	DailyAPIActivityByPartnerReportThursdayTime = ""
	DailyAPIActivityByPartnerReportFridayTime = ""
	DailyAPIActivityByPartnerReportSaturdayTime = ""
	RunDailyAPIActivityByPartnerReportIfClosed = ""
	RunDailyAPIActivityByPartnerReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_API"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_DailyAPIActivityByPartnerReportGeneration = rsFieldServiceSettings("Schedule_DailyAPIActivityByPartnerReportGeneration")
		
		Schedule_DailyAPIActivityByPartnerReportGenerationSettings = Split(Schedule_DailyAPIActivityByPartnerReportGeneration,",")

		DailyAPIActivityByPartnerReportSunday = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(0))
		DailyAPIActivityByPartnerReportMonday = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(1))
		DailyAPIActivityByPartnerReportTuesday = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(2))
		DailyAPIActivityByPartnerReportWednesday = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(3))
		DailyAPIActivityByPartnerReportThursday = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(4))
		DailyAPIActivityByPartnerReportFriday = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(5))
		DailyAPIActivityByPartnerReportSaturday = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(6))
		DailyAPIActivityByPartnerReportSundayTime = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(7)
		DailyAPIActivityByPartnerReportMondayTime = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(8)
		DailyAPIActivityByPartnerReportTuesdayTime = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(9)
		DailyAPIActivityByPartnerReportWednesdayTime = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(10)
		DailyAPIActivityByPartnerReportThursdayTime = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(11)
		DailyAPIActivityByPartnerReportFridayTime = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(12)
		DailyAPIActivityByPartnerReportSaturdayTime = Schedule_DailyAPIActivityByPartnerReportGenerationSettings(13)
		RunDailyAPIActivityByPartnerReportIfClosed = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(14))
		RunDailyAPIActivityByPartnerReportIfClosingEarly = cInt(Schedule_DailyAPIActivityByPartnerReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the Daily API Activity Summary By Partner Report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the Daily API Activity Summary By Partner Report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The Daily API Activity Summary By Partner Report Can Only Be Generated 6:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyAPIActivityByPartnerReportSunday) = 0 Then %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerSunday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportSundayInit" id="txtDailyAPIActivityByPartnerReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerSunday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerSundayTime" value="<%= DailyAPIActivityByPartnerReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportSundayInit" id="txtDailyAPIActivityByPartnerReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(DailyAPIActivityByPartnerReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportSunday" name="chkNoDailyAPIActivityByPartnerReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportSunday" name="chkNoDailyAPIActivityByPartnerReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyAPIActivityByPartnerReportMonday) = 0 Then %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerMonday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportMondayInit" id="txtDailyAPIActivityByPartnerReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerMonday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerMondayTime" value="<%= DailyAPIActivityByPartnerReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportMondayInit" id="txtDailyAPIActivityByPartnerReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyAPIActivityByPartnerReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportMonday" name="chkNoDailyAPIActivityByPartnerReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportMonday" name="chkNoDailyAPIActivityByPartnerReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyAPIActivityByPartnerReportTuesday) = 0 Then %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerTuesday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportTuesdayInit" id="txtDailyAPIActivityByPartnerReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerTuesday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerTuesdayTime" value="<%= DailyAPIActivityByPartnerReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportTuesdayInit" id="txtDailyAPIActivityByPartnerReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyAPIActivityByPartnerReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportTuesday" name="chkNoDailyAPIActivityByPartnerReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportTuesday" name="chkNoDailyAPIActivityByPartnerReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyAPIActivityByPartnerReportWednesday) = 0 Then %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerWednesday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportWednesdayInit" id="txtDailyAPIActivityByPartnerReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerWednesday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerWednesdayTime" value="<%= DailyAPIActivityByPartnerReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportWednesdayInit" id="txtDailyAPIActivityByPartnerReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyAPIActivityByPartnerReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportWednesday" name="chkNoDailyAPIActivityByPartnerReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportWednesday" name="chkNoDailyAPIActivityByPartnerReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyAPIActivityByPartnerReportThursday) = 0 Then %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerThursday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportThursdayInit" id="txtDailyAPIActivityByPartnerReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerThursday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerThursdayTime" value="<%= DailyAPIActivityByPartnerReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportThursdayInit" id="txtDailyAPIActivityByPartnerReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyAPIActivityByPartnerReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportThursday" name="chkNoDailyAPIActivityByPartnerReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportThursday" name="chkNoDailyAPIActivityByPartnerReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyAPIActivityByPartnerReportFriday) = 0 Then %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerFriday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportFridayInit" id="txtDailyAPIActivityByPartnerReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerFriday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerFridayTime" value="<%= DailyAPIActivityByPartnerReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportFridayInit" id="txtDailyAPIActivityByPartnerReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyAPIActivityByPartnerReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportFriday" name="chkNoDailyAPIActivityByPartnerReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportFriday" name="chkNoDailyAPIActivityByPartnerReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyAPIActivityByPartnerReportSaturday) = 0 Then %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerSaturday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportSaturdayInit" id="txtDailyAPIActivityByPartnerReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyAPIActivityByPartnerReportSchedulerSaturday" type="text" name="txtDailyAPIActivityByPartnerReportSchedulerSaturdayTime" value="<%= DailyAPIActivityByPartnerReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyAPIActivityByPartnerReportSaturdayInit" id="txtDailyAPIActivityByPartnerReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyAPIActivityByPartnerReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportSaturday" name="chkNoDailyAPIActivityByPartnerReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportSaturday" name="chkNoDailyAPIActivityByPartnerReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunDailyAPIActivityByPartnerReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportIfClosed" name="chkNoDailyAPIActivityByPartnerReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportIfClosed" name="chkNoDailyAPIActivityByPartnerReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunDailyAPIActivityByPartnerReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportIfClosingEarly" name="chkNoDailyAPIActivityByPartnerReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyAPIActivityByPartnerReportIfClosingEarly" name="chkNoDailyAPIActivityByPartnerReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForDailyInventoryAPIActivityByPartnerReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtDailyInventoryAPIActivityByPartnerReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtDailyInventoryAPIActivityByPartnerReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtDailyInventoryAPIActivityByPartnerReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtDailyInventoryAPIActivityByPartnerReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtDailyInventoryAPIActivityByPartnerReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtDailyInventoryAPIActivityByPartnerReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtDailyInventoryAPIActivityByPartnerReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyInventoryAPIActivityByPartnerReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyInventoryAPIActivityByPartnerReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyInventoryAPIActivityByPartnerReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyInventoryAPIActivityByPartnerReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyInventoryAPIActivityByPartnerReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyInventoryAPIActivityByPartnerReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoDailyInventoryAPIActivityByPartnerReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoDailyInventoryAPIActivityByPartnerReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSunday').timepicker('setTime', '6:00 AM');
			    }
			});
			    	
			$("#chkNoDailyInventoryAPIActivityByPartnerReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerMonday').timepicker('setTime', '6:00 AM');
			    }
			});
	    	
			$("#chkNoDailyInventoryAPIActivityByPartnerReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerTuesday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoDailyInventoryAPIActivityByPartnerReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerWednesday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoDailyInventoryAPIActivityByPartnerReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerThursday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoDailyInventoryAPIActivityByPartnerReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerFriday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoDailyInventoryAPIActivityByPartnerReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSaturday').timepicker('setTime', '6:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing daily inventory api activity by partner report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_DailyInventoryAPIActivityByPartnerReportGeneration = ""
	DailyInventoryAPIActivityByPartnerReportSunday = ""
	DailyInventoryAPIActivityByPartnerReportMonday = ""
	DailyInventoryAPIActivityByPartnerReportTuesday = ""
	DailyInventoryAPIActivityByPartnerReportWednesday = ""
	DailyInventoryAPIActivityByPartnerReportThursday = ""
	DailyInventoryAPIActivityByPartnerReportFriday = ""
	DailyInventoryAPIActivityByPartnerReportSaturday = ""
	DailyInventoryAPIActivityByPartnerReportSundayTime = ""
	DailyInventoryAPIActivityByPartnerReportMondayTime = ""
	DailyInventoryAPIActivityByPartnerReportTuesdayTime = ""
	DailyInventoryAPIActivityByPartnerReportWednesdayTime = ""
	DailyInventoryAPIActivityByPartnerReportThursdayTime = ""
	DailyInventoryAPIActivityByPartnerReportFridayTime = ""
	DailyInventoryAPIActivityByPartnerReportSaturdayTime = ""
	RunDailyInventoryAPIActivityByPartnerReportIfClosed = ""
	RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_InventoryControl"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_DailyInventoryAPIActivityByPartnerReportGeneration = rsFieldServiceSettings("Schedule_DailyInventoryAPIActivityByPartnerReportGeneration")
		
		Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings = Split(Schedule_DailyInventoryAPIActivityByPartnerReportGeneration,",")

		DailyInventoryAPIActivityByPartnerReportSunday = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(0))
		DailyInventoryAPIActivityByPartnerReportMonday = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(1))
		DailyInventoryAPIActivityByPartnerReportTuesday = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(2))
		DailyInventoryAPIActivityByPartnerReportWednesday = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(3))
		DailyInventoryAPIActivityByPartnerReportThursday = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(4))
		DailyInventoryAPIActivityByPartnerReportFriday = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(5))
		DailyInventoryAPIActivityByPartnerReportSaturday = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(6))
		DailyInventoryAPIActivityByPartnerReportSundayTime = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(7)
		DailyInventoryAPIActivityByPartnerReportMondayTime = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(8)
		DailyInventoryAPIActivityByPartnerReportTuesdayTime = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(9)
		DailyInventoryAPIActivityByPartnerReportWednesdayTime = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(10)
		DailyInventoryAPIActivityByPartnerReportThursdayTime = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(11)
		DailyInventoryAPIActivityByPartnerReportFridayTime = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(12)
		DailyInventoryAPIActivityByPartnerReportSaturdayTime = Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(13)
		RunDailyInventoryAPIActivityByPartnerReportIfClosed = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(14))
		RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly = cInt(Schedule_DailyInventoryAPIActivityByPartnerReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the Daily Inventory API Activity Summary By Partner Report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the Daily Inventory API Activity Summary By Partner Report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The Daily Inventory API Activity Summary By Partner Report Can Only Be Generated 6:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyInventoryAPIActivityByPartnerReportSunday) = 0 Then %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSunday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportSundayInit" id="txtDailyInventoryAPIActivityByPartnerReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSunday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerSundayTime" value="<%= DailyInventoryAPIActivityByPartnerReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportSundayInit" id="txtDailyInventoryAPIActivityByPartnerReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(DailyInventoryAPIActivityByPartnerReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportSunday" name="chkNoDailyInventoryAPIActivityByPartnerReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportSunday" name="chkNoDailyInventoryAPIActivityByPartnerReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyInventoryAPIActivityByPartnerReportMonday) = 0 Then %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerMonday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportMondayInit" id="txtDailyInventoryAPIActivityByPartnerReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerMonday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerMondayTime" value="<%= DailyInventoryAPIActivityByPartnerReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportMondayInit" id="txtDailyInventoryAPIActivityByPartnerReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyInventoryAPIActivityByPartnerReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportMonday" name="chkNoDailyInventoryAPIActivityByPartnerReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportMonday" name="chkNoDailyInventoryAPIActivityByPartnerReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyInventoryAPIActivityByPartnerReportTuesday) = 0 Then %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerTuesday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportTuesdayInit" id="txtDailyInventoryAPIActivityByPartnerReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerTuesday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerTuesdayTime" value="<%= DailyInventoryAPIActivityByPartnerReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportTuesdayInit" id="txtDailyInventoryAPIActivityByPartnerReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyInventoryAPIActivityByPartnerReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportTuesday" name="chkNoDailyInventoryAPIActivityByPartnerReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportTuesday" name="chkNoDailyInventoryAPIActivityByPartnerReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyInventoryAPIActivityByPartnerReportWednesday) = 0 Then %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerWednesday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportWednesdayInit" id="txtDailyInventoryAPIActivityByPartnerReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerWednesday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerWednesdayTime" value="<%= DailyInventoryAPIActivityByPartnerReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportWednesdayInit" id="txtDailyInventoryAPIActivityByPartnerReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyInventoryAPIActivityByPartnerReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportWednesday" name="chkNoDailyInventoryAPIActivityByPartnerReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportWednesday" name="chkNoDailyInventoryAPIActivityByPartnerReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyInventoryAPIActivityByPartnerReportThursday) = 0 Then %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerThursday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportThursdayInit" id="txtDailyInventoryAPIActivityByPartnerReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerThursday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerThursdayTime" value="<%= DailyInventoryAPIActivityByPartnerReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportThursdayInit" id="txtDailyInventoryAPIActivityByPartnerReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyInventoryAPIActivityByPartnerReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportThursday" name="chkNoDailyInventoryAPIActivityByPartnerReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportThursday" name="chkNoDailyInventoryAPIActivityByPartnerReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyInventoryAPIActivityByPartnerReportFriday) = 0 Then %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerFriday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportFridayInit" id="txtDailyInventoryAPIActivityByPartnerReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerFriday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerFridayTime" value="<%= DailyInventoryAPIActivityByPartnerReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportFridayInit" id="txtDailyInventoryAPIActivityByPartnerReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyInventoryAPIActivityByPartnerReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportFriday" name="chkNoDailyInventoryAPIActivityByPartnerReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportFriday" name="chkNoDailyInventoryAPIActivityByPartnerReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(DailyInventoryAPIActivityByPartnerReportSaturday) = 0 Then %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSaturday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportSaturdayInit" id="txtDailyInventoryAPIActivityByPartnerReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerDailyInventoryAPIActivityByPartnerReportSchedulerSaturday" type="text" name="txtDailyInventoryAPIActivityByPartnerReportSchedulerSaturdayTime" value="<%= DailyInventoryAPIActivityByPartnerReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtDailyInventoryAPIActivityByPartnerReportSaturdayInit" id="txtDailyInventoryAPIActivityByPartnerReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(DailyInventoryAPIActivityByPartnerReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportSaturday" name="chkNoDailyInventoryAPIActivityByPartnerReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportSaturday" name="chkNoDailyInventoryAPIActivityByPartnerReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunDailyInventoryAPIActivityByPartnerReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportIfClosed" name="chkNoDailyInventoryAPIActivityByPartnerReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportIfClosed" name="chkNoDailyInventoryAPIActivityByPartnerReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunDailyInventoryAPIActivityByPartnerReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportIfClosingEarly" name="chkNoDailyInventoryAPIActivityByPartnerReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoDailyInventoryAPIActivityByPartnerReportIfClosingEarly" name="chkNoDailyInventoryAPIActivityByPartnerReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForInventoryProductChangesReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerInventoryProductChangesReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryProductChangesReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryProductChangesReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryProductChangesReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryProductChangesReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryProductChangesReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryProductChangesReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '6:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtInventoryProductChangesReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerInventoryProductChangesReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtInventoryProductChangesReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerInventoryProductChangesReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtInventoryProductChangesReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerInventoryProductChangesReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtInventoryProductChangesReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerInventoryProductChangesReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtInventoryProductChangesReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerInventoryProductChangesReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtInventoryProductChangesReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerInventoryProductChangesReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtInventoryProductChangesReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerInventoryProductChangesReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerInventoryProductChangesReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryProductChangesReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryProductChangesReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryProductChangesReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryProductChangesReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryProductChangesReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryProductChangesReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryProductChangesReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryProductChangesReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryProductChangesReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryProductChangesReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryProductChangesReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryProductChangesReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryProductChangesReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoInventoryProductChangesReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryProductChangesReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryProductChangesReportSchedulerSunday').timepicker('setTime', '6:00 AM');
			    }
			});
			    	
			$("#chkNoInventoryProductChangesReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryProductChangesReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryProductChangesReportSchedulerMonday').timepicker('setTime', '6:00 AM');
			    }
			});
	    	
			$("#chkNoInventoryProductChangesReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryProductChangesReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryProductChangesReportSchedulerTuesday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoInventoryProductChangesReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryProductChangesReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryProductChangesReportSchedulerWednesday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoInventoryProductChangesReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryProductChangesReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryProductChangesReportSchedulerThursday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoInventoryProductChangesReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryProductChangesReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryProductChangesReportSchedulerFriday').timepicker('setTime', '6:00 AM');
			    }
			});

			$("#chkNoInventoryProductChangesReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryProductChangesReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryProductChangesReportSchedulerSaturday').timepicker('setTime', '6:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing inventory product changes report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,6:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_InventoryProductChangesReportGeneration = ""
	InventoryProductChangesReportSunday = ""
	InventoryProductChangesReportMonday = ""
	InventoryProductChangesReportTuesday = ""
	InventoryProductChangesReportWednesday = ""
	InventoryProductChangesReportThursday = ""
	InventoryProductChangesReportFriday = ""
	InventoryProductChangesReportSaturday = ""
	InventoryProductChangesReportSundayTime = ""
	InventoryProductChangesReportMondayTime = ""
	InventoryProductChangesReportTuesdayTime = ""
	InventoryProductChangesReportWednesdayTime = ""
	InventoryProductChangesReportThursdayTime = ""
	InventoryProductChangesReportFridayTime = ""
	InventoryProductChangesReportSaturdayTime = ""
	RunInventoryProductChangesReportIfClosed = ""
	RunInventoryProductChangesReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_InventoryControl"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_InventoryProductChangesReportGeneration = rsFieldServiceSettings("Schedule_InventoryProductChangesReportGeneration")
		
		Schedule_InventoryProductChangesReportGenerationSettings = Split(Schedule_InventoryProductChangesReportGeneration,",")

		InventoryProductChangesReportSunday = cInt(Schedule_InventoryProductChangesReportGenerationSettings(0))
		InventoryProductChangesReportMonday = cInt(Schedule_InventoryProductChangesReportGenerationSettings(1))
		InventoryProductChangesReportTuesday = cInt(Schedule_InventoryProductChangesReportGenerationSettings(2))
		InventoryProductChangesReportWednesday = cInt(Schedule_InventoryProductChangesReportGenerationSettings(3))
		InventoryProductChangesReportThursday = cInt(Schedule_InventoryProductChangesReportGenerationSettings(4))
		InventoryProductChangesReportFriday = cInt(Schedule_InventoryProductChangesReportGenerationSettings(5))
		InventoryProductChangesReportSaturday = cInt(Schedule_InventoryProductChangesReportGenerationSettings(6))
		InventoryProductChangesReportSundayTime = Schedule_InventoryProductChangesReportGenerationSettings(7)
		InventoryProductChangesReportMondayTime = Schedule_InventoryProductChangesReportGenerationSettings(8)
		InventoryProductChangesReportTuesdayTime = Schedule_InventoryProductChangesReportGenerationSettings(9)
		InventoryProductChangesReportWednesdayTime = Schedule_InventoryProductChangesReportGenerationSettings(10)
		InventoryProductChangesReportThursdayTime = Schedule_InventoryProductChangesReportGenerationSettings(11)
		InventoryProductChangesReportFridayTime = Schedule_InventoryProductChangesReportGenerationSettings(12)
		InventoryProductChangesReportSaturdayTime = Schedule_InventoryProductChangesReportGenerationSettings(13)
		RunInventoryProductChangesReportIfClosed = cInt(Schedule_InventoryProductChangesReportGenerationSettings(14))
		RunInventoryProductChangesReportIfClosingEarly = cInt(Schedule_InventoryProductChangesReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the Inventory Product Changes Report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the Inventory Product Changes Report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The Inventory Product Changes Report Can Only Be Generated 6:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryProductChangesReportSunday) = 0 Then %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerSunday" type="text" name="txtInventoryProductChangesReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportSundayInit" id="txtInventoryProductChangesReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerSunday" type="text" name="txtInventoryProductChangesReportSchedulerSundayTime" value="<%= InventoryProductChangesReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportSundayInit" id="txtInventoryProductChangesReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(InventoryProductChangesReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportSunday" name="chkNoInventoryProductChangesReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportSunday" name="chkNoInventoryProductChangesReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryProductChangesReportMonday) = 0 Then %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerMonday" type="text" name="txtInventoryProductChangesReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportMondayInit" id="txtInventoryProductChangesReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerMonday" type="text" name="txtInventoryProductChangesReportSchedulerMondayTime" value="<%= InventoryProductChangesReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportMondayInit" id="txtInventoryProductChangesReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryProductChangesReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportMonday" name="chkNoInventoryProductChangesReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportMonday" name="chkNoInventoryProductChangesReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryProductChangesReportTuesday) = 0 Then %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerTuesday" type="text" name="txtInventoryProductChangesReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportTuesdayInit" id="txtInventoryProductChangesReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerTuesday" type="text" name="txtInventoryProductChangesReportSchedulerTuesdayTime" value="<%= InventoryProductChangesReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportTuesdayInit" id="txtInventoryProductChangesReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryProductChangesReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportTuesday" name="chkNoInventoryProductChangesReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportTuesday" name="chkNoInventoryProductChangesReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryProductChangesReportWednesday) = 0 Then %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerWednesday" type="text" name="txtInventoryProductChangesReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportWednesdayInit" id="txtInventoryProductChangesReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerWednesday" type="text" name="txtInventoryProductChangesReportSchedulerWednesdayTime" value="<%= InventoryProductChangesReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportWednesdayInit" id="txtInventoryProductChangesReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryProductChangesReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportWednesday" name="chkNoInventoryProductChangesReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportWednesday" name="chkNoInventoryProductChangesReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryProductChangesReportThursday) = 0 Then %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerThursday" type="text" name="txtInventoryProductChangesReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportThursdayInit" id="txtInventoryProductChangesReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerThursday" type="text" name="txtInventoryProductChangesReportSchedulerThursdayTime" value="<%= InventoryProductChangesReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportThursdayInit" id="txtInventoryProductChangesReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryProductChangesReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportThursday" name="chkNoInventoryProductChangesReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportThursday" name="chkNoInventoryProductChangesReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryProductChangesReportFriday) = 0 Then %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerFriday" type="text" name="txtInventoryProductChangesReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportFridayInit" id="txtInventoryProductChangesReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerFriday" type="text" name="txtInventoryProductChangesReportSchedulerFridayTime" value="<%= InventoryProductChangesReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportFridayInit" id="txtInventoryProductChangesReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryProductChangesReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportFriday" name="chkNoInventoryProductChangesReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportFriday" name="chkNoInventoryProductChangesReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryProductChangesReportSaturday) = 0 Then %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerSaturday" type="text" name="txtInventoryProductChangesReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportSaturdayInit" id="txtInventoryProductChangesReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryProductChangesReportSchedulerSaturday" type="text" name="txtInventoryProductChangesReportSchedulerSaturdayTime" value="<%= InventoryProductChangesReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryProductChangesReportSaturdayInit" id="txtInventoryProductChangesReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryProductChangesReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportSaturday" name="chkNoInventoryProductChangesReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportSaturday" name="chkNoInventoryProductChangesReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunInventoryProductChangesReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportIfClosed" name="chkNoInventoryProductChangesReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportIfClosed" name="chkNoInventoryProductChangesReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunInventoryProductChangesReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportIfClosingEarly" name="chkNoInventoryProductChangesReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryProductChangesReportIfClosingEarly" name="chkNoInventoryProductChangesReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForAutomaticCustomerAnalysisSummary1ReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtAutomaticCustomerAnalysisSummary1ReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtAutomaticCustomerAnalysisSummary1ReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtAutomaticCustomerAnalysisSummary1ReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtAutomaticCustomerAnalysisSummary1ReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtAutomaticCustomerAnalysisSummary1ReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtAutomaticCustomerAnalysisSummary1ReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtAutomaticCustomerAnalysisSummary1ReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoAutomaticCustomerAnalysisSummary1ReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoAutomaticCustomerAnalysisSummary1ReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoAutomaticCustomerAnalysisSummary1ReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoAutomaticCustomerAnalysisSummary1ReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoAutomaticCustomerAnalysisSummary1ReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoAutomaticCustomerAnalysisSummary1ReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoAutomaticCustomerAnalysisSummary1ReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoAutomaticCustomerAnalysisSummary1ReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSunday').timepicker('setTime', '10:00 AM');
			    }
			});
			    	
			$("#chkNoAutomaticCustomerAnalysisSummary1ReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerMonday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	
			$("#chkNoAutomaticCustomerAnalysisSummary1ReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerTuesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoAutomaticCustomerAnalysisSummary1ReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerWednesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoAutomaticCustomerAnalysisSummary1ReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerThursday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoAutomaticCustomerAnalysisSummary1ReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerFriday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoAutomaticCustomerAnalysisSummary1ReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSaturday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing inventory product changes report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration = ""
	AutomaticCustomerAnalysisSummary1ReportSunday = ""
	AutomaticCustomerAnalysisSummary1ReportMonday = ""
	AutomaticCustomerAnalysisSummary1ReportTuesday = ""
	AutomaticCustomerAnalysisSummary1ReportWednesday = ""
	AutomaticCustomerAnalysisSummary1ReportThursday = ""
	AutomaticCustomerAnalysisSummary1ReportFriday = ""
	AutomaticCustomerAnalysisSummary1ReportSaturday = ""
	AutomaticCustomerAnalysisSummary1ReportSundayTime = ""
	AutomaticCustomerAnalysisSummary1ReportMondayTime = ""
	AutomaticCustomerAnalysisSummary1ReportTuesdayTime = ""
	AutomaticCustomerAnalysisSummary1ReportWednesdayTime = ""
	AutomaticCustomerAnalysisSummary1ReportThursdayTime = ""
	AutomaticCustomerAnalysisSummary1ReportFridayTime = ""
	AutomaticCustomerAnalysisSummary1ReportSaturdayTime = ""
	RunAutomaticCustomerAnalysisSummary1ReportIfClosed = ""
	RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_BizIntel"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration = rsFieldServiceSettings("Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration")
		
		Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings = Split(Schedule_AutomaticCustomerAnalysisSummary1ReportGeneration,",")

		AutomaticCustomerAnalysisSummary1ReportSunday = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(0))
		AutomaticCustomerAnalysisSummary1ReportMonday = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(1))
		AutomaticCustomerAnalysisSummary1ReportTuesday = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(2))
		AutomaticCustomerAnalysisSummary1ReportWednesday = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(3))
		AutomaticCustomerAnalysisSummary1ReportThursday = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(4))
		AutomaticCustomerAnalysisSummary1ReportFriday = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(5))
		AutomaticCustomerAnalysisSummary1ReportSaturday = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(6))
		AutomaticCustomerAnalysisSummary1ReportSundayTime = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(7)
		AutomaticCustomerAnalysisSummary1ReportMondayTime = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(8)
		AutomaticCustomerAnalysisSummary1ReportTuesdayTime = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(9)
		AutomaticCustomerAnalysisSummary1ReportWednesdayTime = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(10)
		AutomaticCustomerAnalysisSummary1ReportThursdayTime = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(11)
		AutomaticCustomerAnalysisSummary1ReportFridayTime = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(12)
		AutomaticCustomerAnalysisSummary1ReportSaturdayTime = Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(13)
		RunAutomaticCustomerAnalysisSummary1ReportIfClosed = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(14))
		RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly = cInt(Schedule_AutomaticCustomerAnalysisSummary1ReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the Automatic Customer Analysis Summary 1 Report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the Automatic Customer Analysis Summary 1 Report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The Automatic Customer Analysis Summary 1 Report Can Only Be Generated 10:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(AutomaticCustomerAnalysisSummary1ReportSunday) = 0 Then %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSunday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportSundayInit" id="txtAutomaticCustomerAnalysisSummary1ReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSunday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerSundayTime" value="<%= AutomaticCustomerAnalysisSummary1ReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportSundayInit" id="txtAutomaticCustomerAnalysisSummary1ReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(AutomaticCustomerAnalysisSummary1ReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportSunday" name="chkNoAutomaticCustomerAnalysisSummary1ReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportSunday" name="chkNoAutomaticCustomerAnalysisSummary1ReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(AutomaticCustomerAnalysisSummary1ReportMonday) = 0 Then %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerMonday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportMondayInit" id="txtAutomaticCustomerAnalysisSummary1ReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerMonday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerMondayTime" value="<%= AutomaticCustomerAnalysisSummary1ReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportMondayInit" id="txtAutomaticCustomerAnalysisSummary1ReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(AutomaticCustomerAnalysisSummary1ReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportMonday" name="chkNoAutomaticCustomerAnalysisSummary1ReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportMonday" name="chkNoAutomaticCustomerAnalysisSummary1ReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(AutomaticCustomerAnalysisSummary1ReportTuesday) = 0 Then %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerTuesday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportTuesdayInit" id="txtAutomaticCustomerAnalysisSummary1ReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerTuesday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerTuesdayTime" value="<%= AutomaticCustomerAnalysisSummary1ReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportTuesdayInit" id="txtAutomaticCustomerAnalysisSummary1ReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(AutomaticCustomerAnalysisSummary1ReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportTuesday" name="chkNoAutomaticCustomerAnalysisSummary1ReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportTuesday" name="chkNoAutomaticCustomerAnalysisSummary1ReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(AutomaticCustomerAnalysisSummary1ReportWednesday) = 0 Then %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerWednesday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportWednesdayInit" id="txtAutomaticCustomerAnalysisSummary1ReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerWednesday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerWednesdayTime" value="<%= AutomaticCustomerAnalysisSummary1ReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportWednesdayInit" id="txtAutomaticCustomerAnalysisSummary1ReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(AutomaticCustomerAnalysisSummary1ReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportWednesday" name="chkNoAutomaticCustomerAnalysisSummary1ReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportWednesday" name="chkNoAutomaticCustomerAnalysisSummary1ReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(AutomaticCustomerAnalysisSummary1ReportThursday) = 0 Then %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerThursday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportThursdayInit" id="txtAutomaticCustomerAnalysisSummary1ReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerThursday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerThursdayTime" value="<%= AutomaticCustomerAnalysisSummary1ReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportThursdayInit" id="txtAutomaticCustomerAnalysisSummary1ReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(AutomaticCustomerAnalysisSummary1ReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportThursday" name="chkNoAutomaticCustomerAnalysisSummary1ReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportThursday" name="chkNoAutomaticCustomerAnalysisSummary1ReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(AutomaticCustomerAnalysisSummary1ReportFriday) = 0 Then %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerFriday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportFridayInit" id="txtAutomaticCustomerAnalysisSummary1ReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerFriday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerFridayTime" value="<%= AutomaticCustomerAnalysisSummary1ReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportFridayInit" id="txtAutomaticCustomerAnalysisSummary1ReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(AutomaticCustomerAnalysisSummary1ReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportFriday" name="chkNoAutomaticCustomerAnalysisSummary1ReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportFriday" name="chkNoAutomaticCustomerAnalysisSummary1ReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(AutomaticCustomerAnalysisSummary1ReportSaturday) = 0 Then %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSaturday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportSaturdayInit" id="txtAutomaticCustomerAnalysisSummary1ReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerAutomaticCustomerAnalysisSummary1ReportSchedulerSaturday" type="text" name="txtAutomaticCustomerAnalysisSummary1ReportSchedulerSaturdayTime" value="<%= AutomaticCustomerAnalysisSummary1ReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtAutomaticCustomerAnalysisSummary1ReportSaturdayInit" id="txtAutomaticCustomerAnalysisSummary1ReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(AutomaticCustomerAnalysisSummary1ReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportSaturday" name="chkNoAutomaticCustomerAnalysisSummary1ReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportSaturday" name="chkNoAutomaticCustomerAnalysisSummary1ReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunAutomaticCustomerAnalysisSummary1ReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportIfClosed" name="chkNoAutomaticCustomerAnalysisSummary1ReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportIfClosed" name="chkNoAutomaticCustomerAnalysisSummary1ReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunAutomaticCustomerAnalysisSummary1ReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportIfClosingEarly" name="chkNoAutomaticCustomerAnalysisSummary1ReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoAutomaticCustomerAnalysisSummary1ReportIfClosingEarly" name="chkNoAutomaticCustomerAnalysisSummary1ReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForMCSActivityReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerMCSActivityReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerMCSActivityReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerMCSActivityReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerMCSActivityReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerMCSActivityReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerMCSActivityReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerMCSActivityReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtMCSActivityReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerMCSActivityReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtMCSActivityReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerMCSActivityReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtMCSActivityReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerMCSActivityReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtMCSActivityReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerMCSActivityReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtMCSActivityReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerMCSActivityReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtMCSActivityReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerMCSActivityReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtMCSActivityReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerMCSActivityReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerMCSActivityReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoMCSActivityReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerMCSActivityReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoMCSActivityReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerMCSActivityReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoMCSActivityReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerMCSActivityReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoMCSActivityReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerMCSActivityReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoMCSActivityReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerMCSActivityReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoMCSActivityReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerMCSActivityReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoMCSActivityReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoMCSActivityReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerMCSActivityReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerMCSActivityReportSchedulerSunday').timepicker('setTime', '10:00 AM');
			    }
			});
			    	
			$("#chkNoMCSActivityReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerMCSActivityReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerMCSActivityReportSchedulerMonday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	
			$("#chkNoMCSActivityReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerMCSActivityReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerMCSActivityReportSchedulerTuesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoMCSActivityReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerMCSActivityReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerMCSActivityReportSchedulerWednesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoMCSActivityReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerMCSActivityReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerMCSActivityReportSchedulerThursday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoMCSActivityReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerMCSActivityReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerMCSActivityReportSchedulerFriday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoMCSActivityReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerMCSActivityReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerMCSActivityReportSchedulerSaturday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing mcs activity report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_MCSActivityReportGeneration = ""
	MCSActivityReportSunday = ""
	MCSActivityReportMonday = ""
	MCSActivityReportTuesday = ""
	MCSActivityReportWednesday = ""
	MCSActivityReportThursday = ""
	MCSActivityReportFriday = ""
	MCSActivityReportSaturday = ""
	MCSActivityReportSundayTime = ""
	MCSActivityReportMondayTime = ""
	MCSActivityReportTuesdayTime = ""
	MCSActivityReportWednesdayTime = ""
	MCSActivityReportThursdayTime = ""
	MCSActivityReportFridayTime = ""
	MCSActivityReportSaturdayTime = ""
	RunMCSActivityReportIfClosed = ""
	RunMCSActivityReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_BizIntel"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_MCSActivityReportGeneration = rsFieldServiceSettings("Schedule_MCSActivityReportGeneration")
		
		Schedule_MCSActivityReportGenerationSettings = Split(Schedule_MCSActivityReportGeneration,",")

		MCSActivityReportSunday = cInt(Schedule_MCSActivityReportGenerationSettings(0))
		MCSActivityReportMonday = cInt(Schedule_MCSActivityReportGenerationSettings(1))
		MCSActivityReportTuesday = cInt(Schedule_MCSActivityReportGenerationSettings(2))
		MCSActivityReportWednesday = cInt(Schedule_MCSActivityReportGenerationSettings(3))
		MCSActivityReportThursday = cInt(Schedule_MCSActivityReportGenerationSettings(4))
		MCSActivityReportFriday = cInt(Schedule_MCSActivityReportGenerationSettings(5))
		MCSActivityReportSaturday = cInt(Schedule_MCSActivityReportGenerationSettings(6))
		MCSActivityReportSundayTime = Schedule_MCSActivityReportGenerationSettings(7)
		MCSActivityReportMondayTime = Schedule_MCSActivityReportGenerationSettings(8)
		MCSActivityReportTuesdayTime = Schedule_MCSActivityReportGenerationSettings(9)
		MCSActivityReportWednesdayTime = Schedule_MCSActivityReportGenerationSettings(10)
		MCSActivityReportThursdayTime = Schedule_MCSActivityReportGenerationSettings(11)
		MCSActivityReportFridayTime = Schedule_MCSActivityReportGenerationSettings(12)
		MCSActivityReportSaturdayTime = Schedule_MCSActivityReportGenerationSettings(13)
		RunMCSActivityReportIfClosed = cInt(Schedule_MCSActivityReportGenerationSettings(14))
		RunMCSActivityReportIfClosingEarly = cInt(Schedule_MCSActivityReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the MCS Activity Report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the MCS Activity Report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The MCS Activity Report Can Only Be Generated 10:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(MCSActivityReportSunday) = 0 Then %>
				  		<input id="timepickerMCSActivityReportSchedulerSunday" type="text" name="txtMCSActivityReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportSundayInit" id="txtMCSActivityReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerMCSActivityReportSchedulerSunday" type="text" name="txtMCSActivityReportSchedulerSundayTime" value="<%= MCSActivityReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportSundayInit" id="txtMCSActivityReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(MCSActivityReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoMCSActivityReportSunday" name="chkNoMCSActivityReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoMCSActivityReportSunday" name="chkNoMCSActivityReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(MCSActivityReportMonday) = 0 Then %>
				  		<input id="timepickerMCSActivityReportSchedulerMonday" type="text" name="txtMCSActivityReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportMondayInit" id="txtMCSActivityReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerMCSActivityReportSchedulerMonday" type="text" name="txtMCSActivityReportSchedulerMondayTime" value="<%= MCSActivityReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportMondayInit" id="txtMCSActivityReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(MCSActivityReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoMCSActivityReportMonday" name="chkNoMCSActivityReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoMCSActivityReportMonday" name="chkNoMCSActivityReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(MCSActivityReportTuesday) = 0 Then %>
				  		<input id="timepickerMCSActivityReportSchedulerTuesday" type="text" name="txtMCSActivityReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportTuesdayInit" id="txtMCSActivityReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerMCSActivityReportSchedulerTuesday" type="text" name="txtMCSActivityReportSchedulerTuesdayTime" value="<%= MCSActivityReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportTuesdayInit" id="txtMCSActivityReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(MCSActivityReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoMCSActivityReportTuesday" name="chkNoMCSActivityReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoMCSActivityReportTuesday" name="chkNoMCSActivityReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(MCSActivityReportWednesday) = 0 Then %>
				  		<input id="timepickerMCSActivityReportSchedulerWednesday" type="text" name="txtMCSActivityReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportWednesdayInit" id="txtMCSActivityReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerMCSActivityReportSchedulerWednesday" type="text" name="txtMCSActivityReportSchedulerWednesdayTime" value="<%= MCSActivityReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportWednesdayInit" id="txtMCSActivityReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(MCSActivityReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoMCSActivityReportWednesday" name="chkNoMCSActivityReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoMCSActivityReportWednesday" name="chkNoMCSActivityReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(MCSActivityReportThursday) = 0 Then %>
				  		<input id="timepickerMCSActivityReportSchedulerThursday" type="text" name="txtMCSActivityReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportThursdayInit" id="txtMCSActivityReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerMCSActivityReportSchedulerThursday" type="text" name="txtMCSActivityReportSchedulerThursdayTime" value="<%= MCSActivityReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportThursdayInit" id="txtMCSActivityReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(MCSActivityReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoMCSActivityReportThursday" name="chkNoMCSActivityReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoMCSActivityReportThursday" name="chkNoMCSActivityReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(MCSActivityReportFriday) = 0 Then %>
				  		<input id="timepickerMCSActivityReportSchedulerFriday" type="text" name="txtMCSActivityReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportFridayInit" id="txtMCSActivityReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerMCSActivityReportSchedulerFriday" type="text" name="txtMCSActivityReportSchedulerFridayTime" value="<%= MCSActivityReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportFridayInit" id="txtMCSActivityReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(MCSActivityReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoMCSActivityReportFriday" name="chkNoMCSActivityReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoMCSActivityReportFriday" name="chkNoMCSActivityReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(MCSActivityReportSaturday) = 0 Then %>
				  		<input id="timepickerMCSActivityReportSchedulerSaturday" type="text" name="txtMCSActivityReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportSaturdayInit" id="txtMCSActivityReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerMCSActivityReportSchedulerSaturday" type="text" name="txtMCSActivityReportSchedulerSaturdayTime" value="<%= MCSActivityReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtMCSActivityReportSaturdayInit" id="txtMCSActivityReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(MCSActivityReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoMCSActivityReportSaturday" name="chkNoMCSActivityReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoMCSActivityReportSaturday" name="chkNoMCSActivityReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunMCSActivityReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoMCSActivityReportIfClosed" name="chkNoMCSActivityReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoMCSActivityReportIfClosed" name="chkNoMCSActivityReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunMCSActivityReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoMCSActivityReportIfClosingEarly" name="chkNoMCSActivityReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoMCSActivityReportIfClosingEarly" name="chkNoMCSActivityReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForOrderAPIN2KReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerOrderAPINeedToKnowReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerOrderAPINeedToKnowReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerOrderAPINeedToKnowReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerOrderAPINeedToKnowReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerOrderAPINeedToKnowReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerOrderAPINeedToKnowReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerOrderAPINeedToKnowReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtOrderAPINeedToKnowReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerOrderAPINeedToKnowReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtOrderAPINeedToKnowReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerOrderAPINeedToKnowReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtOrderAPINeedToKnowReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerOrderAPINeedToKnowReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtOrderAPINeedToKnowReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerOrderAPINeedToKnowReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtOrderAPINeedToKnowReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerOrderAPINeedToKnowReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtOrderAPINeedToKnowReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerOrderAPINeedToKnowReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtOrderAPINeedToKnowReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerOrderAPINeedToKnowReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerOrderAPINeedToKnowReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoOrderAPINeedToKnowReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerOrderAPINeedToKnowReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoOrderAPINeedToKnowReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerOrderAPINeedToKnowReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoOrderAPINeedToKnowReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerOrderAPINeedToKnowReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoOrderAPINeedToKnowReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerOrderAPINeedToKnowReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoOrderAPINeedToKnowReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerOrderAPINeedToKnowReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoOrderAPINeedToKnowReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerOrderAPINeedToKnowReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoOrderAPINeedToKnowReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoOrderAPINeedToKnowReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerOrderAPINeedToKnowReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerOrderAPINeedToKnowReportSchedulerSunday').timepicker('setTime', '10:00 AM');
			    }
			});
			    	
			$("#chkNoOrderAPINeedToKnowReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerOrderAPINeedToKnowReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerOrderAPINeedToKnowReportSchedulerMonday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	
			$("#chkNoOrderAPINeedToKnowReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerOrderAPINeedToKnowReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerOrderAPINeedToKnowReportSchedulerTuesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoOrderAPINeedToKnowReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerOrderAPINeedToKnowReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerOrderAPINeedToKnowReportSchedulerWednesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoOrderAPINeedToKnowReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerOrderAPINeedToKnowReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerOrderAPINeedToKnowReportSchedulerThursday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoOrderAPINeedToKnowReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerOrderAPINeedToKnowReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerOrderAPINeedToKnowReportSchedulerFriday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoOrderAPINeedToKnowReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerOrderAPINeedToKnowReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerOrderAPINeedToKnowReportSchedulerSaturday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing mcs activity report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_OrderAPINeedToKnowReportGeneration = ""
	OrderAPINeedToKnowReportSunday = ""
	OrderAPINeedToKnowReportMonday = ""
	OrderAPINeedToKnowReportTuesday = ""
	OrderAPINeedToKnowReportWednesday = ""
	OrderAPINeedToKnowReportThursday = ""
	OrderAPINeedToKnowReportFriday = ""
	OrderAPINeedToKnowReportSaturday = ""
	OrderAPINeedToKnowReportSundayTime = ""
	OrderAPINeedToKnowReportMondayTime = ""
	OrderAPINeedToKnowReportTuesdayTime = ""
	OrderAPINeedToKnowReportWednesdayTime = ""
	OrderAPINeedToKnowReportThursdayTime = ""
	OrderAPINeedToKnowReportFridayTime = ""
	OrderAPINeedToKnowReportSaturdayTime = ""
	RunOrderAPINeedToKnowReportIfClosed = ""
	RunOrderAPINeedToKnowReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_OrderAPINeedToKnowReportGeneration = rsFieldServiceSettings("Schedule_APINeedToKnowReportGeneration")
		
		Schedule_OrderAPINeedToKnowReportGenerationSettings = Split(Schedule_OrderAPINeedToKnowReportGeneration,",")

		OrderAPINeedToKnowReportSunday = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(0))
		OrderAPINeedToKnowReportMonday = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(1))
		OrderAPINeedToKnowReportTuesday = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(2))
		OrderAPINeedToKnowReportWednesday = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(3))
		OrderAPINeedToKnowReportThursday = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(4))
		OrderAPINeedToKnowReportFriday = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(5))
		OrderAPINeedToKnowReportSaturday = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(6))
		OrderAPINeedToKnowReportSundayTime = Schedule_OrderAPINeedToKnowReportGenerationSettings(7)
		OrderAPINeedToKnowReportMondayTime = Schedule_OrderAPINeedToKnowReportGenerationSettings(8)
		OrderAPINeedToKnowReportTuesdayTime = Schedule_OrderAPINeedToKnowReportGenerationSettings(9)
		OrderAPINeedToKnowReportWednesdayTime = Schedule_OrderAPINeedToKnowReportGenerationSettings(10)
		OrderAPINeedToKnowReportThursdayTime = Schedule_OrderAPINeedToKnowReportGenerationSettings(11)
		OrderAPINeedToKnowReportFridayTime = Schedule_OrderAPINeedToKnowReportGenerationSettings(12)
		OrderAPINeedToKnowReportSaturdayTime = Schedule_OrderAPINeedToKnowReportGenerationSettings(13)
		RunOrderAPINeedToKnowReportIfClosed = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(14))
		RunOrderAPINeedToKnowReportIfClosingEarly = cInt(Schedule_OrderAPINeedToKnowReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the Order API Need To Know Report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the Order API Need To Know Report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The Order API Need To Know Report Can Only Be Generated 10:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(OrderAPINeedToKnowReportSunday) = 0 Then %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerSunday" type="text" name="txtOrderAPINeedToKnowReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportSundayInit" id="txtOrderAPINeedToKnowReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerSunday" type="text" name="txtOrderAPINeedToKnowReportSchedulerSundayTime" value="<%= OrderAPINeedToKnowReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportSundayInit" id="txtOrderAPINeedToKnowReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(OrderAPINeedToKnowReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportSunday" name="chkNoOrderAPINeedToKnowReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportSunday" name="chkNoOrderAPINeedToKnowReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(OrderAPINeedToKnowReportMonday) = 0 Then %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerMonday" type="text" name="txtOrderAPINeedToKnowReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportMondayInit" id="txtOrderAPINeedToKnowReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerMonday" type="text" name="txtOrderAPINeedToKnowReportSchedulerMondayTime" value="<%= OrderAPINeedToKnowReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportMondayInit" id="txtOrderAPINeedToKnowReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(OrderAPINeedToKnowReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportMonday" name="chkNoOrderAPINeedToKnowReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportMonday" name="chkNoOrderAPINeedToKnowReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(OrderAPINeedToKnowReportTuesday) = 0 Then %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerTuesday" type="text" name="txtOrderAPINeedToKnowReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportTuesdayInit" id="txtOrderAPINeedToKnowReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerTuesday" type="text" name="txtOrderAPINeedToKnowReportSchedulerTuesdayTime" value="<%= OrderAPINeedToKnowReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportTuesdayInit" id="txtOrderAPINeedToKnowReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(OrderAPINeedToKnowReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportTuesday" name="chkNoOrderAPINeedToKnowReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportTuesday" name="chkNoOrderAPINeedToKnowReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(OrderAPINeedToKnowReportWednesday) = 0 Then %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerWednesday" type="text" name="txtOrderAPINeedToKnowReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportWednesdayInit" id="txtOrderAPINeedToKnowReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerWednesday" type="text" name="txtOrderAPINeedToKnowReportSchedulerWednesdayTime" value="<%= OrderAPINeedToKnowReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportWednesdayInit" id="txtOrderAPINeedToKnowReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(OrderAPINeedToKnowReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportWednesday" name="chkNoOrderAPINeedToKnowReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportWednesday" name="chkNoOrderAPINeedToKnowReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(OrderAPINeedToKnowReportThursday) = 0 Then %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerThursday" type="text" name="txtOrderAPINeedToKnowReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportThursdayInit" id="txtOrderAPINeedToKnowReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerThursday" type="text" name="txtOrderAPINeedToKnowReportSchedulerThursdayTime" value="<%= OrderAPINeedToKnowReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportThursdayInit" id="txtOrderAPINeedToKnowReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(OrderAPINeedToKnowReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportThursday" name="chkNoOrderAPINeedToKnowReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportThursday" name="chkNoOrderAPINeedToKnowReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(OrderAPINeedToKnowReportFriday) = 0 Then %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerFriday" type="text" name="txtOrderAPINeedToKnowReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportFridayInit" id="txtOrderAPINeedToKnowReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerFriday" type="text" name="txtOrderAPINeedToKnowReportSchedulerFridayTime" value="<%= OrderAPINeedToKnowReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportFridayInit" id="txtOrderAPINeedToKnowReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(OrderAPINeedToKnowReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportFriday" name="chkNoOrderAPINeedToKnowReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportFriday" name="chkNoOrderAPINeedToKnowReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(OrderAPINeedToKnowReportSaturday) = 0 Then %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerSaturday" type="text" name="txtOrderAPINeedToKnowReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportSaturdayInit" id="txtOrderAPINeedToKnowReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerOrderAPINeedToKnowReportSchedulerSaturday" type="text" name="txtOrderAPINeedToKnowReportSchedulerSaturdayTime" value="<%= OrderAPINeedToKnowReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtOrderAPINeedToKnowReportSaturdayInit" id="txtOrderAPINeedToKnowReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(OrderAPINeedToKnowReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportSaturday" name="chkNoOrderAPINeedToKnowReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportSaturday" name="chkNoOrderAPINeedToKnowReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunOrderAPINeedToKnowReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportIfClosed" name="chkNoOrderAPINeedToKnowReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportIfClosed" name="chkNoOrderAPINeedToKnowReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunOrderAPINeedToKnowReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportIfClosingEarly" name="chkNoOrderAPINeedToKnowReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoOrderAPINeedToKnowReportIfClosingEarly" name="chkNoOrderAPINeedToKnowReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForEquipmentN2KReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerEquipmentNeedToKnowReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerEquipmentNeedToKnowReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerEquipmentNeedToKnowReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerEquipmentNeedToKnowReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerEquipmentNeedToKnowReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerEquipmentNeedToKnowReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerEquipmentNeedToKnowReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtEquipmentNeedToKnowReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerEquipmentNeedToKnowReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtEquipmentNeedToKnowReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerEquipmentNeedToKnowReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtEquipmentNeedToKnowReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerEquipmentNeedToKnowReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtEquipmentNeedToKnowReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerEquipmentNeedToKnowReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtEquipmentNeedToKnowReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerEquipmentNeedToKnowReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtEquipmentNeedToKnowReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerEquipmentNeedToKnowReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtEquipmentNeedToKnowReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerEquipmentNeedToKnowReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerEquipmentNeedToKnowReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoEquipmentNeedToKnowReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerEquipmentNeedToKnowReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoEquipmentNeedToKnowReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerEquipmentNeedToKnowReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoEquipmentNeedToKnowReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerEquipmentNeedToKnowReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoEquipmentNeedToKnowReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerEquipmentNeedToKnowReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoEquipmentNeedToKnowReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerEquipmentNeedToKnowReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoEquipmentNeedToKnowReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerEquipmentNeedToKnowReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoEquipmentNeedToKnowReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoEquipmentNeedToKnowReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerEquipmentNeedToKnowReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerEquipmentNeedToKnowReportSchedulerSunday').timepicker('setTime', '10:00 AM');
			    }
			});
			    	
			$("#chkNoEquipmentNeedToKnowReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerEquipmentNeedToKnowReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerEquipmentNeedToKnowReportSchedulerMonday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	
			$("#chkNoEquipmentNeedToKnowReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerEquipmentNeedToKnowReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerEquipmentNeedToKnowReportSchedulerTuesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoEquipmentNeedToKnowReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerEquipmentNeedToKnowReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerEquipmentNeedToKnowReportSchedulerWednesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoEquipmentNeedToKnowReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerEquipmentNeedToKnowReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerEquipmentNeedToKnowReportSchedulerThursday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoEquipmentNeedToKnowReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerEquipmentNeedToKnowReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerEquipmentNeedToKnowReportSchedulerFriday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoEquipmentNeedToKnowReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerEquipmentNeedToKnowReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerEquipmentNeedToKnowReportSchedulerSaturday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing mcs activity report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_EquipmentNeedToKnowReportGeneration = ""
	EquipmentNeedToKnowReportSunday = ""
	EquipmentNeedToKnowReportMonday = ""
	EquipmentNeedToKnowReportTuesday = ""
	EquipmentNeedToKnowReportWednesday = ""
	EquipmentNeedToKnowReportThursday = ""
	EquipmentNeedToKnowReportFriday = ""
	EquipmentNeedToKnowReportSaturday = ""
	EquipmentNeedToKnowReportSundayTime = ""
	EquipmentNeedToKnowReportMondayTime = ""
	EquipmentNeedToKnowReportTuesdayTime = ""
	EquipmentNeedToKnowReportWednesdayTime = ""
	EquipmentNeedToKnowReportThursdayTime = ""
	EquipmentNeedToKnowReportFridayTime = ""
	EquipmentNeedToKnowReportSaturdayTime = ""
	RunEquipmentNeedToKnowReportIfClosed = ""
	RunEquipmentNeedToKnowReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_EquipmentNeedToKnowReportGeneration = rsFieldServiceSettings("Schedule_EquipmentNeedToKnowReportGeneration")
		
		Schedule_EquipmentNeedToKnowReportGenerationSettings = Split(Schedule_EquipmentNeedToKnowReportGeneration,",")

		EquipmentNeedToKnowReportSunday = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(0))
		EquipmentNeedToKnowReportMonday = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(1))
		EquipmentNeedToKnowReportTuesday = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(2))
		EquipmentNeedToKnowReportWednesday = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(3))
		EquipmentNeedToKnowReportThursday = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(4))
		EquipmentNeedToKnowReportFriday = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(5))
		EquipmentNeedToKnowReportSaturday = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(6))
		EquipmentNeedToKnowReportSundayTime = Schedule_EquipmentNeedToKnowReportGenerationSettings(7)
		EquipmentNeedToKnowReportMondayTime = Schedule_EquipmentNeedToKnowReportGenerationSettings(8)
		EquipmentNeedToKnowReportTuesdayTime = Schedule_EquipmentNeedToKnowReportGenerationSettings(9)
		EquipmentNeedToKnowReportWednesdayTime = Schedule_EquipmentNeedToKnowReportGenerationSettings(10)
		EquipmentNeedToKnowReportThursdayTime = Schedule_EquipmentNeedToKnowReportGenerationSettings(11)
		EquipmentNeedToKnowReportFridayTime = Schedule_EquipmentNeedToKnowReportGenerationSettings(12)
		EquipmentNeedToKnowReportSaturdayTime = Schedule_EquipmentNeedToKnowReportGenerationSettings(13)
		RunEquipmentNeedToKnowReportIfClosed = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(14))
		RunEquipmentNeedToKnowReportIfClosingEarly = cInt(Schedule_EquipmentNeedToKnowReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the Equipment Need To Know Report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the Equipment Need To Know Report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The Equipment Need To Know Report Can Only Be Generated 10:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(EquipmentNeedToKnowReportSunday) = 0 Then %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerSunday" type="text" name="txtEquipmentNeedToKnowReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportSundayInit" id="txtEquipmentNeedToKnowReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerSunday" type="text" name="txtEquipmentNeedToKnowReportSchedulerSundayTime" value="<%= EquipmentNeedToKnowReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportSundayInit" id="txtEquipmentNeedToKnowReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(EquipmentNeedToKnowReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportSunday" name="chkNoEquipmentNeedToKnowReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportSunday" name="chkNoEquipmentNeedToKnowReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(EquipmentNeedToKnowReportMonday) = 0 Then %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerMonday" type="text" name="txtEquipmentNeedToKnowReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportMondayInit" id="txtEquipmentNeedToKnowReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerMonday" type="text" name="txtEquipmentNeedToKnowReportSchedulerMondayTime" value="<%= EquipmentNeedToKnowReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportMondayInit" id="txtEquipmentNeedToKnowReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(EquipmentNeedToKnowReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportMonday" name="chkNoEquipmentNeedToKnowReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportMonday" name="chkNoEquipmentNeedToKnowReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(EquipmentNeedToKnowReportTuesday) = 0 Then %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerTuesday" type="text" name="txtEquipmentNeedToKnowReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportTuesdayInit" id="txtEquipmentNeedToKnowReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerTuesday" type="text" name="txtEquipmentNeedToKnowReportSchedulerTuesdayTime" value="<%= EquipmentNeedToKnowReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportTuesdayInit" id="txtEquipmentNeedToKnowReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(EquipmentNeedToKnowReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportTuesday" name="chkNoEquipmentNeedToKnowReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportTuesday" name="chkNoEquipmentNeedToKnowReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(EquipmentNeedToKnowReportWednesday) = 0 Then %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerWednesday" type="text" name="txtEquipmentNeedToKnowReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportWednesdayInit" id="txtEquipmentNeedToKnowReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerWednesday" type="text" name="txtEquipmentNeedToKnowReportSchedulerWednesdayTime" value="<%= EquipmentNeedToKnowReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportWednesdayInit" id="txtEquipmentNeedToKnowReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(EquipmentNeedToKnowReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportWednesday" name="chkNoEquipmentNeedToKnowReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportWednesday" name="chkNoEquipmentNeedToKnowReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(EquipmentNeedToKnowReportThursday) = 0 Then %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerThursday" type="text" name="txtEquipmentNeedToKnowReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportThursdayInit" id="txtEquipmentNeedToKnowReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerThursday" type="text" name="txtEquipmentNeedToKnowReportSchedulerThursdayTime" value="<%= EquipmentNeedToKnowReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportThursdayInit" id="txtEquipmentNeedToKnowReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(EquipmentNeedToKnowReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportThursday" name="chkNoEquipmentNeedToKnowReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportThursday" name="chkNoEquipmentNeedToKnowReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(EquipmentNeedToKnowReportFriday) = 0 Then %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerFriday" type="text" name="txtEquipmentNeedToKnowReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportFridayInit" id="txtEquipmentNeedToKnowReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerFriday" type="text" name="txtEquipmentNeedToKnowReportSchedulerFridayTime" value="<%= EquipmentNeedToKnowReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportFridayInit" id="txtEquipmentNeedToKnowReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(EquipmentNeedToKnowReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportFriday" name="chkNoEquipmentNeedToKnowReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportFriday" name="chkNoEquipmentNeedToKnowReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(EquipmentNeedToKnowReportSaturday) = 0 Then %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerSaturday" type="text" name="txtEquipmentNeedToKnowReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportSaturdayInit" id="txtEquipmentNeedToKnowReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerEquipmentNeedToKnowReportSchedulerSaturday" type="text" name="txtEquipmentNeedToKnowReportSchedulerSaturdayTime" value="<%= EquipmentNeedToKnowReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtEquipmentNeedToKnowReportSaturdayInit" id="txtEquipmentNeedToKnowReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(EquipmentNeedToKnowReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportSaturday" name="chkNoEquipmentNeedToKnowReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportSaturday" name="chkNoEquipmentNeedToKnowReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunEquipmentNeedToKnowReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportIfClosed" name="chkNoEquipmentNeedToKnowReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportIfClosed" name="chkNoEquipmentNeedToKnowReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunEquipmentNeedToKnowReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportIfClosingEarly" name="chkNoEquipmentNeedToKnowReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoEquipmentNeedToKnowReportIfClosingEarly" name="chkNoEquipmentNeedToKnowReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForInventoryN2KReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerInventoryNeedToKnowReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryNeedToKnowReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryNeedToKnowReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryNeedToKnowReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryNeedToKnowReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryNeedToKnowReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerInventoryNeedToKnowReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtInventoryNeedToKnowReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerInventoryNeedToKnowReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtInventoryNeedToKnowReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerInventoryNeedToKnowReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtInventoryNeedToKnowReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerInventoryNeedToKnowReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtInventoryNeedToKnowReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerInventoryNeedToKnowReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtInventoryNeedToKnowReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerInventoryNeedToKnowReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtInventoryNeedToKnowReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerInventoryNeedToKnowReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtInventoryNeedToKnowReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerInventoryNeedToKnowReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerInventoryNeedToKnowReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryNeedToKnowReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryNeedToKnowReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryNeedToKnowReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryNeedToKnowReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryNeedToKnowReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryNeedToKnowReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryNeedToKnowReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryNeedToKnowReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryNeedToKnowReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryNeedToKnowReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryNeedToKnowReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerInventoryNeedToKnowReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoInventoryNeedToKnowReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoInventoryNeedToKnowReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryNeedToKnowReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryNeedToKnowReportSchedulerSunday').timepicker('setTime', '10:00 AM');
			    }
			});
			    	
			$("#chkNoInventoryNeedToKnowReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryNeedToKnowReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryNeedToKnowReportSchedulerMonday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	
			$("#chkNoInventoryNeedToKnowReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryNeedToKnowReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryNeedToKnowReportSchedulerTuesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoInventoryNeedToKnowReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryNeedToKnowReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryNeedToKnowReportSchedulerWednesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoInventoryNeedToKnowReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryNeedToKnowReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryNeedToKnowReportSchedulerThursday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoInventoryNeedToKnowReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryNeedToKnowReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryNeedToKnowReportSchedulerFriday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoInventoryNeedToKnowReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerInventoryNeedToKnowReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerInventoryNeedToKnowReportSchedulerSaturday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing mcs activity report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_InventoryNeedToKnowReportGeneration = ""
	InventoryNeedToKnowReportSunday = ""
	InventoryNeedToKnowReportMonday = ""
	InventoryNeedToKnowReportTuesday = ""
	InventoryNeedToKnowReportWednesday = ""
	InventoryNeedToKnowReportThursday = ""
	InventoryNeedToKnowReportFriday = ""
	InventoryNeedToKnowReportSaturday = ""
	InventoryNeedToKnowReportSundayTime = ""
	InventoryNeedToKnowReportMondayTime = ""
	InventoryNeedToKnowReportTuesdayTime = ""
	InventoryNeedToKnowReportWednesdayTime = ""
	InventoryNeedToKnowReportThursdayTime = ""
	InventoryNeedToKnowReportFridayTime = ""
	InventoryNeedToKnowReportSaturdayTime = ""
	RunInventoryNeedToKnowReportIfClosed = ""
	RunInventoryNeedToKnowReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_InventoryNeedToKnowReportGeneration = rsFieldServiceSettings("Schedule_InventoryNeedToKnowReportGeneration")
		
		Schedule_InventoryNeedToKnowReportGenerationSettings = Split(Schedule_InventoryNeedToKnowReportGeneration,",")

		InventoryNeedToKnowReportSunday = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(0))
		InventoryNeedToKnowReportMonday = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(1))
		InventoryNeedToKnowReportTuesday = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(2))
		InventoryNeedToKnowReportWednesday = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(3))
		InventoryNeedToKnowReportThursday = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(4))
		InventoryNeedToKnowReportFriday = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(5))
		InventoryNeedToKnowReportSaturday = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(6))
		InventoryNeedToKnowReportSundayTime = Schedule_InventoryNeedToKnowReportGenerationSettings(7)
		InventoryNeedToKnowReportMondayTime = Schedule_InventoryNeedToKnowReportGenerationSettings(8)
		InventoryNeedToKnowReportTuesdayTime = Schedule_InventoryNeedToKnowReportGenerationSettings(9)
		InventoryNeedToKnowReportWednesdayTime = Schedule_InventoryNeedToKnowReportGenerationSettings(10)
		InventoryNeedToKnowReportThursdayTime = Schedule_InventoryNeedToKnowReportGenerationSettings(11)
		InventoryNeedToKnowReportFridayTime = Schedule_InventoryNeedToKnowReportGenerationSettings(12)
		InventoryNeedToKnowReportSaturdayTime = Schedule_InventoryNeedToKnowReportGenerationSettings(13)
		RunInventoryNeedToKnowReportIfClosed = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(14))
		RunInventoryNeedToKnowReportIfClosingEarly = cInt(Schedule_InventoryNeedToKnowReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the Inventory Need To Know Report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the Inventory Need To Know Report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The Inventory Need To Know Report Can Only Be Generated 10:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryNeedToKnowReportSunday) = 0 Then %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerSunday" type="text" name="txtInventoryNeedToKnowReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportSundayInit" id="txtInventoryNeedToKnowReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerSunday" type="text" name="txtInventoryNeedToKnowReportSchedulerSundayTime" value="<%= InventoryNeedToKnowReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportSundayInit" id="txtInventoryNeedToKnowReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(InventoryNeedToKnowReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportSunday" name="chkNoInventoryNeedToKnowReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportSunday" name="chkNoInventoryNeedToKnowReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryNeedToKnowReportMonday) = 0 Then %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerMonday" type="text" name="txtInventoryNeedToKnowReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportMondayInit" id="txtInventoryNeedToKnowReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerMonday" type="text" name="txtInventoryNeedToKnowReportSchedulerMondayTime" value="<%= InventoryNeedToKnowReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportMondayInit" id="txtInventoryNeedToKnowReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryNeedToKnowReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportMonday" name="chkNoInventoryNeedToKnowReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportMonday" name="chkNoInventoryNeedToKnowReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryNeedToKnowReportTuesday) = 0 Then %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerTuesday" type="text" name="txtInventoryNeedToKnowReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportTuesdayInit" id="txtInventoryNeedToKnowReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerTuesday" type="text" name="txtInventoryNeedToKnowReportSchedulerTuesdayTime" value="<%= InventoryNeedToKnowReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportTuesdayInit" id="txtInventoryNeedToKnowReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryNeedToKnowReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportTuesday" name="chkNoInventoryNeedToKnowReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportTuesday" name="chkNoInventoryNeedToKnowReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryNeedToKnowReportWednesday) = 0 Then %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerWednesday" type="text" name="txtInventoryNeedToKnowReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportWednesdayInit" id="txtInventoryNeedToKnowReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerWednesday" type="text" name="txtInventoryNeedToKnowReportSchedulerWednesdayTime" value="<%= InventoryNeedToKnowReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportWednesdayInit" id="txtInventoryNeedToKnowReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryNeedToKnowReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportWednesday" name="chkNoInventoryNeedToKnowReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportWednesday" name="chkNoInventoryNeedToKnowReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryNeedToKnowReportThursday) = 0 Then %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerThursday" type="text" name="txtInventoryNeedToKnowReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportThursdayInit" id="txtInventoryNeedToKnowReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerThursday" type="text" name="txtInventoryNeedToKnowReportSchedulerThursdayTime" value="<%= InventoryNeedToKnowReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportThursdayInit" id="txtInventoryNeedToKnowReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryNeedToKnowReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportThursday" name="chkNoInventoryNeedToKnowReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportThursday" name="chkNoInventoryNeedToKnowReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryNeedToKnowReportFriday) = 0 Then %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerFriday" type="text" name="txtInventoryNeedToKnowReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportFridayInit" id="txtInventoryNeedToKnowReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerFriday" type="text" name="txtInventoryNeedToKnowReportSchedulerFridayTime" value="<%= InventoryNeedToKnowReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportFridayInit" id="txtInventoryNeedToKnowReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryNeedToKnowReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportFriday" name="chkNoInventoryNeedToKnowReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportFriday" name="chkNoInventoryNeedToKnowReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(InventoryNeedToKnowReportSaturday) = 0 Then %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerSaturday" type="text" name="txtInventoryNeedToKnowReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportSaturdayInit" id="txtInventoryNeedToKnowReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerInventoryNeedToKnowReportSchedulerSaturday" type="text" name="txtInventoryNeedToKnowReportSchedulerSaturdayTime" value="<%= InventoryNeedToKnowReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtInventoryNeedToKnowReportSaturdayInit" id="txtInventoryNeedToKnowReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(InventoryNeedToKnowReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportSaturday" name="chkNoInventoryNeedToKnowReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportSaturday" name="chkNoInventoryNeedToKnowReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunInventoryNeedToKnowReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportIfClosed" name="chkNoInventoryNeedToKnowReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportIfClosed" name="chkNoInventoryNeedToKnowReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunInventoryNeedToKnowReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportIfClosingEarly" name="chkNoInventoryNeedToKnowReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoInventoryNeedToKnowReportIfClosingEarly" name="chkNoInventoryNeedToKnowReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForGlobalSettingsN2KReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtGlobalSettingsNeedToKnowReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerGlobalSettingsNeedToKnowReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtGlobalSettingsNeedToKnowReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerGlobalSettingsNeedToKnowReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtGlobalSettingsNeedToKnowReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerGlobalSettingsNeedToKnowReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtGlobalSettingsNeedToKnowReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerGlobalSettingsNeedToKnowReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtGlobalSettingsNeedToKnowReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerGlobalSettingsNeedToKnowReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtGlobalSettingsNeedToKnowReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerGlobalSettingsNeedToKnowReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtGlobalSettingsNeedToKnowReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerGlobalSettingsNeedToKnowReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerGlobalSettingsNeedToKnowReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoGlobalSettingsNeedToKnowReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerGlobalSettingsNeedToKnowReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoGlobalSettingsNeedToKnowReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerGlobalSettingsNeedToKnowReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoGlobalSettingsNeedToKnowReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerGlobalSettingsNeedToKnowReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoGlobalSettingsNeedToKnowReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerGlobalSettingsNeedToKnowReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoGlobalSettingsNeedToKnowReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerGlobalSettingsNeedToKnowReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoGlobalSettingsNeedToKnowReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerGlobalSettingsNeedToKnowReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoGlobalSettingsNeedToKnowReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoGlobalSettingsNeedToKnowReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerGlobalSettingsNeedToKnowReportSchedulerSunday').timepicker('setTime', '10:00 AM');
			    }
			});
			    	
			$("#chkNoGlobalSettingsNeedToKnowReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerGlobalSettingsNeedToKnowReportSchedulerMonday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	
			$("#chkNoGlobalSettingsNeedToKnowReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerGlobalSettingsNeedToKnowReportSchedulerTuesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoGlobalSettingsNeedToKnowReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerGlobalSettingsNeedToKnowReportSchedulerWednesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoGlobalSettingsNeedToKnowReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerGlobalSettingsNeedToKnowReportSchedulerThursday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoGlobalSettingsNeedToKnowReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerGlobalSettingsNeedToKnowReportSchedulerFriday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoGlobalSettingsNeedToKnowReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerGlobalSettingsNeedToKnowReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerGlobalSettingsNeedToKnowReportSchedulerSaturday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing mcs activity report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_GlobalSettingsNeedToKnowReportGeneration = ""
	GlobalSettingsNeedToKnowReportSunday = ""
	GlobalSettingsNeedToKnowReportMonday = ""
	GlobalSettingsNeedToKnowReportTuesday = ""
	GlobalSettingsNeedToKnowReportWednesday = ""
	GlobalSettingsNeedToKnowReportThursday = ""
	GlobalSettingsNeedToKnowReportFriday = ""
	GlobalSettingsNeedToKnowReportSaturday = ""
	GlobalSettingsNeedToKnowReportSundayTime = ""
	GlobalSettingsNeedToKnowReportMondayTime = ""
	GlobalSettingsNeedToKnowReportTuesdayTime = ""
	GlobalSettingsNeedToKnowReportWednesdayTime = ""
	GlobalSettingsNeedToKnowReportThursdayTime = ""
	GlobalSettingsNeedToKnowReportFridayTime = ""
	GlobalSettingsNeedToKnowReportSaturdayTime = ""
	RunGlobalSettingsNeedToKnowReportIfClosed = ""
	RunGlobalSettingsNeedToKnowReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_GlobalSettingsNeedToKnowReportGeneration = rsFieldServiceSettings("Schedule_GlobalSettingsNeedToKnowReportGeneration")
		
		Schedule_GlobalSettingsNeedToKnowReportGenerationSettings = Split(Schedule_GlobalSettingsNeedToKnowReportGeneration,",")

		GlobalSettingsNeedToKnowReportSunday = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(0))
		GlobalSettingsNeedToKnowReportMonday = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(1))
		GlobalSettingsNeedToKnowReportTuesday = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(2))
		GlobalSettingsNeedToKnowReportWednesday = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(3))
		GlobalSettingsNeedToKnowReportThursday = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(4))
		GlobalSettingsNeedToKnowReportFriday = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(5))
		GlobalSettingsNeedToKnowReportSaturday = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(6))
		GlobalSettingsNeedToKnowReportSundayTime = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(7)
		GlobalSettingsNeedToKnowReportMondayTime = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(8)
		GlobalSettingsNeedToKnowReportTuesdayTime = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(9)
		GlobalSettingsNeedToKnowReportWednesdayTime = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(10)
		GlobalSettingsNeedToKnowReportThursdayTime = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(11)
		GlobalSettingsNeedToKnowReportFridayTime = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(12)
		GlobalSettingsNeedToKnowReportSaturdayTime = Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(13)
		RunGlobalSettingsNeedToKnowReportIfClosed = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(14))
		RunGlobalSettingsNeedToKnowReportIfClosingEarly = cInt(Schedule_GlobalSettingsNeedToKnowReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the Global Settings Need To Know Report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the Global Settings Need To Know Report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The Global Settings Need To Know Report Can Only Be Generated 10:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GlobalSettingsNeedToKnowReportSunday) = 0 Then %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerSunday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportSundayInit" id="txtGlobalSettingsNeedToKnowReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerSunday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerSundayTime" value="<%= GlobalSettingsNeedToKnowReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportSundayInit" id="txtGlobalSettingsNeedToKnowReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(GlobalSettingsNeedToKnowReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportSunday" name="chkNoGlobalSettingsNeedToKnowReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportSunday" name="chkNoGlobalSettingsNeedToKnowReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GlobalSettingsNeedToKnowReportMonday) = 0 Then %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerMonday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportMondayInit" id="txtGlobalSettingsNeedToKnowReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerMonday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerMondayTime" value="<%= GlobalSettingsNeedToKnowReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportMondayInit" id="txtGlobalSettingsNeedToKnowReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GlobalSettingsNeedToKnowReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportMonday" name="chkNoGlobalSettingsNeedToKnowReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportMonday" name="chkNoGlobalSettingsNeedToKnowReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GlobalSettingsNeedToKnowReportTuesday) = 0 Then %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerTuesday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportTuesdayInit" id="txtGlobalSettingsNeedToKnowReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerTuesday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerTuesdayTime" value="<%= GlobalSettingsNeedToKnowReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportTuesdayInit" id="txtGlobalSettingsNeedToKnowReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GlobalSettingsNeedToKnowReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportTuesday" name="chkNoGlobalSettingsNeedToKnowReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportTuesday" name="chkNoGlobalSettingsNeedToKnowReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GlobalSettingsNeedToKnowReportWednesday) = 0 Then %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerWednesday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportWednesdayInit" id="txtGlobalSettingsNeedToKnowReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerWednesday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerWednesdayTime" value="<%= GlobalSettingsNeedToKnowReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportWednesdayInit" id="txtGlobalSettingsNeedToKnowReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GlobalSettingsNeedToKnowReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportWednesday" name="chkNoGlobalSettingsNeedToKnowReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportWednesday" name="chkNoGlobalSettingsNeedToKnowReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GlobalSettingsNeedToKnowReportThursday) = 0 Then %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerThursday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportThursdayInit" id="txtGlobalSettingsNeedToKnowReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerThursday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerThursdayTime" value="<%= GlobalSettingsNeedToKnowReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportThursdayInit" id="txtGlobalSettingsNeedToKnowReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GlobalSettingsNeedToKnowReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportThursday" name="chkNoGlobalSettingsNeedToKnowReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportThursday" name="chkNoGlobalSettingsNeedToKnowReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GlobalSettingsNeedToKnowReportFriday) = 0 Then %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerFriday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportFridayInit" id="txtGlobalSettingsNeedToKnowReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerFriday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerFridayTime" value="<%= GlobalSettingsNeedToKnowReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportFridayInit" id="txtGlobalSettingsNeedToKnowReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GlobalSettingsNeedToKnowReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportFriday" name="chkNoGlobalSettingsNeedToKnowReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportFriday" name="chkNoGlobalSettingsNeedToKnowReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(GlobalSettingsNeedToKnowReportSaturday) = 0 Then %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerSaturday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportSaturdayInit" id="txtGlobalSettingsNeedToKnowReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerGlobalSettingsNeedToKnowReportSchedulerSaturday" type="text" name="txtGlobalSettingsNeedToKnowReportSchedulerSaturdayTime" value="<%= GlobalSettingsNeedToKnowReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtGlobalSettingsNeedToKnowReportSaturdayInit" id="txtGlobalSettingsNeedToKnowReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(GlobalSettingsNeedToKnowReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportSaturday" name="chkNoGlobalSettingsNeedToKnowReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportSaturday" name="chkNoGlobalSettingsNeedToKnowReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunGlobalSettingsNeedToKnowReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportIfClosed" name="chkNoGlobalSettingsNeedToKnowReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportIfClosed" name="chkNoGlobalSettingsNeedToKnowReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunGlobalSettingsNeedToKnowReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportIfClosingEarly" name="chkNoGlobalSettingsNeedToKnowReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoGlobalSettingsNeedToKnowReportIfClosingEarly" name="chkNoGlobalSettingsNeedToKnowReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetContentForAccountsReceivableN2KReportScheduler() 

	%>
	
	<script type="text/javascript">
	
		$(document).ready(function() {

	        $('#timepickerFinanceNeedToKnowReportSchedulerSunday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFinanceNeedToKnowReportSchedulerMonday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFinanceNeedToKnowReportSchedulerTuesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFinanceNeedToKnowReportSchedulerWednesday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFinanceNeedToKnowReportSchedulerThursday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFinanceNeedToKnowReportSchedulerFriday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });
	        $('#timepickerFinanceNeedToKnowReportSchedulerSaturday').timepicker({
	            minuteStep: 15,
	            secondStep: 5,
	            showInputs: false,
	            template: 'dropdown',
	            showSeconds: false,
	            showMeridian: true,
           		minTime: '10:00 AM',
            	maxTime: '12:00 AM'
	        });

		
			var initGenTimeSunday = $('#txtFinanceNeedToKnowReportSundayInit').val();
			
			if (initGenTimeSunday == 0) {
				$('#timepickerFinanceNeedToKnowReportSchedulerSunday').timepicker('clear');
			}

			var initGenTimeMonday = $('#txtFinanceNeedToKnowReportMondayInit').val();
			
			if (initGenTimeMonday == 0) {
				$('#timepickerFinanceNeedToKnowReportSchedulerMonday').timepicker('clear');
			}

			var initGenTimeTuesday = $('#txtFinanceNeedToKnowReportTuesdayInit').val();
			
			if (initGenTimeTuesday == 0) {
				$('#timepickerFinanceNeedToKnowReportSchedulerTuesday').timepicker('clear');
			}

			var initGenTimeWednesday = $('#txtFinanceNeedToKnowReportWednesdayInit').val();
			
			if (initGenTimeWednesday == 0) {
				$('#timepickerFinanceNeedToKnowReportSchedulerWednesday').timepicker('clear');
			}

			var initGenTimeThursday = $('#txtFinanceNeedToKnowReportThursdayInit').val();
			
			if (initGenTimeThursday == 0) {
				$('#timepickerFinanceNeedToKnowReportSchedulerThursday').timepicker('clear');
			}

			var initGenTimeFriday = $('#txtFinanceNeedToKnowReportFridayInit').val();
			
			if (initGenTimeFriday == 0) {
				$('#timepickerFinanceNeedToKnowReportSchedulerFriday').timepicker('clear');
			}

			var initGenTimeSaturday = $('#txtFinanceNeedToKnowReportSaturdayInit').val();
			
			if (initGenTimeSaturday == 0) {
				$('#timepickerFinanceNeedToKnowReportSchedulerSaturday').timepicker('clear');
			}
			
		    $('#timepickerFinanceNeedToKnowReportSchedulerSunday').on('show.timepicker', function(e) {
		    	$("#chkNoFinanceNeedToKnowReportSunday").prop( "checked", false );		    
		    });
		    $('#timepickerFinanceNeedToKnowReportSchedulerMonday').on('show.timepicker', function(e) {
		    	$("#chkNoFinanceNeedToKnowReportMonday").prop( "checked", false );		    
		    });
		    $('#timepickerFinanceNeedToKnowReportSchedulerTuesday').on('show.timepicker', function(e) {
		    	$("#chkNoFinanceNeedToKnowReportTuesday").prop( "checked", false );		    
		    });
		    $('#timepickerFinanceNeedToKnowReportSchedulerWednesday').on('show.timepicker', function(e) {
		    	$("#chkNoFinanceNeedToKnowReportWednesday").prop( "checked", false );		    
		    });
		    $('#timepickerFinanceNeedToKnowReportSchedulerThursday').on('show.timepicker', function(e) {
		    	$("#chkNoFinanceNeedToKnowReportThursday").prop( "checked", false );		    
		    });
		    $('#timepickerFinanceNeedToKnowReportSchedulerFriday').on('show.timepicker', function(e) {
		    	$("#chkNoFinanceNeedToKnowReportFriday").prop( "checked", false );		    
		    });
		    $('#timepickerFinanceNeedToKnowReportSchedulerSaturday').on('show.timepicker', function(e) {
		    	$("#chkNoFinanceNeedToKnowReportSaturday").prop( "checked", false );		    
		    });
  
	    	
			$("#chkNoFinanceNeedToKnowReportSunday").change(function() {
			    if(this.checked) {
			        $('#timepickerFinanceNeedToKnowReportSchedulerSunday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFinanceNeedToKnowReportSchedulerSunday').timepicker('setTime', '10:00 AM');
			    }
			});
			    	
			$("#chkNoFinanceNeedToKnowReportMonday").change(function() {
			    if(this.checked) {
			        $('#timepickerFinanceNeedToKnowReportSchedulerMonday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFinanceNeedToKnowReportSchedulerMonday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	
			$("#chkNoFinanceNeedToKnowReportTuesday").change(function() {
			    if(this.checked) {
			        $('#timepickerFinanceNeedToKnowReportSchedulerTuesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFinanceNeedToKnowReportSchedulerTuesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoFinanceNeedToKnowReportWednesday").change(function() {
			    if(this.checked) {
			        $('#timepickerFinanceNeedToKnowReportSchedulerWednesday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFinanceNeedToKnowReportSchedulerWednesday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoFinanceNeedToKnowReportThursday").change(function() {
			    if(this.checked) {
			        $('#timepickerFinanceNeedToKnowReportSchedulerThursday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFinanceNeedToKnowReportSchedulerThursday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoFinanceNeedToKnowReportFriday").change(function() {
			    if(this.checked) {
			        $('#timepickerFinanceNeedToKnowReportSchedulerFriday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFinanceNeedToKnowReportSchedulerFriday').timepicker('setTime', '10:00 AM');
			    }
			});

			$("#chkNoFinanceNeedToKnowReportSaturday").change(function() {
			    if(this.checked) {
			        $('#timepickerFinanceNeedToKnowReportSchedulerSaturday').timepicker('clear');
			    }
			    else {
			    	$('#timepickerFinanceNeedToKnowReportSchedulerSaturday').timepicker('setTime', '10:00 AM');
			    }
			});
	    	    		
		});
	</script>
	
	<%
	'***************************************************************************************
	'Get values for editing an existing mcs activity report gen schedule
	'***************************************************************************************
	
	'DEFAULT VALUES ARE:
	'0,0,0,0,0,0,0,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,10:00 AM,0,0
	
	'***************************************************************************************
	
	'FIELDS 1-7
	'S on/off, M on/off, T on/off, W on/off, Th on/off, F on/off, S on/off,
	
	'***************************************************************************************
	
	'FIELDS 8-14
	'S gen time, M gen time, T gen time, W gen time, Th gen time, F gen time, S gen time
	
	'***************************************************************************************
	
	'FIELDS 15-16
	'Do not run if closed (on/off), Do not run if closing early (on/off)
	
	'***************************************************************************************
	
	Schedule_FinanceNeedToKnowReportGeneration = ""
	FinanceNeedToKnowReportSunday = ""
	FinanceNeedToKnowReportMonday = ""
	FinanceNeedToKnowReportTuesday = ""
	FinanceNeedToKnowReportWednesday = ""
	FinanceNeedToKnowReportThursday = ""
	FinanceNeedToKnowReportFriday = ""
	FinanceNeedToKnowReportSaturday = ""
	FinanceNeedToKnowReportSundayTime = ""
	FinanceNeedToKnowReportMondayTime = ""
	FinanceNeedToKnowReportTuesdayTime = ""
	FinanceNeedToKnowReportWednesdayTime = ""
	FinanceNeedToKnowReportThursdayTime = ""
	FinanceNeedToKnowReportFridayTime = ""
	FinanceNeedToKnowReportSaturdayTime = ""
	RunFinanceNeedToKnowReportIfClosed = ""
	RunFinanceNeedToKnowReportIfClosingEarly = ""

	SQLFieldServiceSettings = "SELECT * FROM Settings_NeedToKnow"
	
	Set cnnFieldServiceSettings = Server.CreateObject("ADODB.Connection")
	cnnFieldServiceSettings.open (Session("ClientCnnString"))
	Set rsFieldServiceSettings = Server.CreateObject("ADODB.Recordset")
	rsFieldServiceSettings.CursorLocation = 3 
	Set rsFieldServiceSettings = cnnFieldServiceSettings.Execute(SQLFieldServiceSettings)
		
	If NOT rsFieldServiceSettings.EOF Then
	
		Schedule_FinanceNeedToKnowReportGeneration = rsFieldServiceSettings("Schedule_FinanceNeedToKnowReportGeneration")
		
		Schedule_FinanceNeedToKnowReportGenerationSettings = Split(Schedule_FinanceNeedToKnowReportGeneration,",")

		FinanceNeedToKnowReportSunday = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(0))
		FinanceNeedToKnowReportMonday = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(1))
		FinanceNeedToKnowReportTuesday = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(2))
		FinanceNeedToKnowReportWednesday = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(3))
		FinanceNeedToKnowReportThursday = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(4))
		FinanceNeedToKnowReportFriday = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(5))
		FinanceNeedToKnowReportSaturday = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(6))
		FinanceNeedToKnowReportSundayTime = Schedule_FinanceNeedToKnowReportGenerationSettings(7)
		FinanceNeedToKnowReportMondayTime = Schedule_FinanceNeedToKnowReportGenerationSettings(8)
		FinanceNeedToKnowReportTuesdayTime = Schedule_FinanceNeedToKnowReportGenerationSettings(9)
		FinanceNeedToKnowReportWednesdayTime = Schedule_FinanceNeedToKnowReportGenerationSettings(10)
		FinanceNeedToKnowReportThursdayTime = Schedule_FinanceNeedToKnowReportGenerationSettings(11)
		FinanceNeedToKnowReportFridayTime = Schedule_FinanceNeedToKnowReportGenerationSettings(12)
		FinanceNeedToKnowReportSaturdayTime = Schedule_FinanceNeedToKnowReportGenerationSettings(13)
		RunFinanceNeedToKnowReportIfClosed = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(14))
		RunFinanceNeedToKnowReportIfClosingEarly = cInt(Schedule_FinanceNeedToKnowReportGenerationSettings(15))
	
	End If
	
	set rsFieldServiceSettings = Nothing
	cnnFieldServiceSettings.close
	set cnnFieldServiceSettings = Nothing
	
	'***************************************************************************************
%>
		<style>
			
			.bootstrap-timepicker-widget.dropdown-menu { z-index: 3000!important; } 
			
			.row-line{
				margin-bottom:15px;
			}
			
			h4 { 
				margin-top: 10px;
			}
			
		</style>
		
		
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<h4>Select a time to run the <%= GetTerm("Accounts Receivable") %> Need To Know Report on specific days.</h4>
				<h4>Check the checkbox if you <strong>do not</strong> want to run the <%= GetTerm("Accounts Receivable") %> Need To Know Report on a particular day.</h4>
				<div class="alert alert-info">
				  <strong>Please Note:</strong> The <%= GetTerm("Accounts Receivable") %> Need To Know Report Can Only Be Generated 10:00 AM - 12:00 AM (midnight) each day.
				</div>
			</div>
		</div>
		
        
		<!-- email alert line !-->
		<div class="row row-line">

			<div class="col-lg-2 text-right">
				<strong>Sunday</strong>
			</div>
			
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FinanceNeedToKnowReportSunday) = 0 Then %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerSunday" type="text" name="txtFinanceNeedToKnowReportSchedulerSundayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportSundayInit" id="txtFinanceNeedToKnowReportSundayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerSunday" type="text" name="txtFinanceNeedToKnowReportSchedulerSundayTime" value="<%= FinanceNeedToKnowReportSundayTime %>" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportSundayInit" id="txtFinanceNeedToKnowReportSundayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>
			
			<div class="col-lg-6">			
				<% If cInt(FinanceNeedToKnowReportSunday) = 0 Then %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportSunday" name="chkNoFinanceNeedToKnowReportSunday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportSunday" name="chkNoFinanceNeedToKnowReportSunday">
				<% End If %>
				Do <strong>Not</strong> Run On Sunday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Monday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FinanceNeedToKnowReportMonday) = 0 Then %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerMonday" type="text" name="txtFinanceNeedToKnowReportSchedulerMondayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportMondayInit" id="txtFinanceNeedToKnowReportMondayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerMonday" type="text" name="txtFinanceNeedToKnowReportSchedulerMondayTime" value="<%= FinanceNeedToKnowReportMondayTime %>" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportMondayInit" id="txtFinanceNeedToKnowReportMondayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FinanceNeedToKnowReportMonday) = 0 Then %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportMonday" name="chkNoFinanceNeedToKnowReportMonday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportMonday" name="chkNoFinanceNeedToKnowReportMonday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Monday
			</div>
			
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Tuesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FinanceNeedToKnowReportTuesday) = 0 Then %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerTuesday" type="text" name="txtFinanceNeedToKnowReportSchedulerTuesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportTuesdayInit" id="txtFinanceNeedToKnowReportTuesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerTuesday" type="text" name="txtFinanceNeedToKnowReportSchedulerTuesdayTime" value="<%= FinanceNeedToKnowReportTuesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportTuesdayInit" id="txtFinanceNeedToKnowReportTuesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FinanceNeedToKnowReportTuesday) = 0 Then %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportTuesday" name="chkNoFinanceNeedToKnowReportTuesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportTuesday" name="chkNoFinanceNeedToKnowReportTuesday">
				<% End If %>
				Do <strong>Not</strong> Run On Tuesday
			</div>
			
        </div>
        <!-- eof when line !-->
        
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Wednesday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FinanceNeedToKnowReportWednesday) = 0 Then %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerWednesday" type="text" name="txtFinanceNeedToKnowReportSchedulerWednesdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportWednesdayInit" id="txtFinanceNeedToKnowReportWednesdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerWednesday" type="text" name="txtFinanceNeedToKnowReportSchedulerWednesdayTime" value="<%= FinanceNeedToKnowReportWednesdayTime %>" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportWednesdayInit" id="txtFinanceNeedToKnowReportWednesdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FinanceNeedToKnowReportWednesday) = 0 Then %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportWednesday" name="chkNoFinanceNeedToKnowReportWednesday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportWednesday" name="chkNoFinanceNeedToKnowReportWednesday">
				<% End If %>
				Do <strong>Not</strong> Run On Wednesday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Thursday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FinanceNeedToKnowReportThursday) = 0 Then %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerThursday" type="text" name="txtFinanceNeedToKnowReportSchedulerThursdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportThursdayInit" id="txtFinanceNeedToKnowReportThursdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerThursday" type="text" name="txtFinanceNeedToKnowReportSchedulerThursdayTime" value="<%= FinanceNeedToKnowReportThursdayTime %>" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportThursdayInit" id="txtFinanceNeedToKnowReportThursdayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FinanceNeedToKnowReportThursday) = 0 Then %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportThursday" name="chkNoFinanceNeedToKnowReportThursday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportThursday" name="chkNoFinanceNeedToKnowReportThursday">
				<% End If %>
				Do <strong>Not</strong> Run On Thursday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Friday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FinanceNeedToKnowReportFriday) = 0 Then %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerFriday" type="text" name="txtFinanceNeedToKnowReportSchedulerFridayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportFridayInit" id="txtFinanceNeedToKnowReportFridayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerFriday" type="text" name="txtFinanceNeedToKnowReportSchedulerFridayTime" value="<%= FinanceNeedToKnowReportFridayTime %>" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportFridayInit" id="txtFinanceNeedToKnowReportFridayInit" value="1">
				  	<% End If %>
					<span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FinanceNeedToKnowReportFriday) = 0 Then %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportFriday" name="chkNoFinanceNeedToKnowReportFriday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportFriday" name="chkNoFinanceNeedToKnowReportFriday">
				<% End If %>
				Do <strong>Not</strong> Run On Friday
			</div>
			
        </div>
        <!-- eof when line !-->
			
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			
			<div class="col-lg-2 text-right">
				<strong>Saturday</strong>
			</div>
			
			<!-- multi select !-->
			<div class="col-lg-4">		        
				<div class="input-group bootstrap-timepicker timepicker">
					<% If cInt(FinanceNeedToKnowReportSaturday) = 0 Then %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerSaturday" type="text" name="txtFinanceNeedToKnowReportSchedulerSaturdayTime" value="" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportSaturdayInit" id="txtFinanceNeedToKnowReportSaturdayInit" value="0">
				  	<% Else %>
				  		<input id="timepickerFinanceNeedToKnowReportSchedulerSaturday" type="text" name="txtFinanceNeedToKnowReportSchedulerSaturdayTime" value="<%= FinanceNeedToKnowReportSaturdayTime %>" class="form-control">
				  		<input type="hidden" name="txtFinanceNeedToKnowReportSaturdayInit" id="txtFinanceNeedToKnowReportSaturdayInit" value="1">
				  	<% End If %>
				 	 <span class="input-group-addon"><i class="glyphicon glyphicon-time"></i></span>
				</div>		        	    
			</div>

			<div class="col-lg-6">			
				<% If cInt(FinanceNeedToKnowReportSaturday) = 0 Then %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportSaturday" name="chkNoFinanceNeedToKnowReportSaturday" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportSaturday" name="chkNoFinanceNeedToKnowReportSaturday">
				<% End If %>
					
				Do <strong>Not</strong> Run On Saturday
			</div>
			
        </div>
        <!-- eof when line !-->
        

		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunFinanceNeedToKnowReportIfClosed) = 0 Then %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportIfClosed" name="chkNoFinanceNeedToKnowReportIfClosed" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportIfClosed" name="chkNoFinanceNeedToKnowReportIfClosed">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closed (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
        
		<!-- email alert line !-->
		<div class="row row-line">
			<div class="col-lg-12">
				<% If cInt(RunFinanceNeedToKnowReportIfClosingEarly) = 0 Then %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportIfClosingEarly" name="chkNoFinanceNeedToKnowReportIfClosingEarly" checked="checked">
				<% Else %>
					<input type="checkbox" id="chkNoFinanceNeedToKnowReportIfClosingEarly" name="chkNoFinanceNeedToKnowReportIfClosingEarly">
				<% End If %>
				&nbsp;&nbsp;Don't Run If Closing Early (Monday-Friday Only)
			</div>
        </div>
        <!-- eof when line !-->
        
	

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************

'END ALL AJAX MODAL SUBROUTINES AND FUNCTIONS

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

%>