<!--#include file="../../inc/header.asp"-->

<!-- function that gets the value of the tab when it is clicked and then
updates the value of a hidden form field so when the page posts, it returns
back to the tab that was previously opened -->

<script type="text/javascript">
	$(function () {
		$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
		var target = $(e.target).attr("href");
		$('input[name="txtTab"]').val(target);
		});
	})
</script>


<script src="<%= BaseURL %>js/bootstrap-yearly-calendar/bootstrap-year-calendar.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>js/bootstrap-yearly-calendar/bootstrap-year-calendar.css">

<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-clockpicker/clockpicker.js"></script>
<link rel="stylesheet" type="text/css" href="<%= BaseURL %>js/bootstrap-clockpicker/clockpicker.css" />

<script src="http://cdnjs.cloudflare.com/ajax/libs/moment.js/2.10.3/moment.js"></script>

<%


Server.ScriptTimeout = 500

SQLBuildCalendarDataSource = "SELECT * FROM Settings_CompanyCalendar ORDER BY YearNum"

Set cnnBuildCalendarDataSource = Server.CreateObject("ADODB.Connection")
cnnBuildCalendarDataSource.open (Session("ClientCnnString"))
Set rsBuildCalendarDataSource = Server.CreateObject("ADODB.Recordset")
rsBuildCalendarDataSource.CursorLocation = 3 

Set rsBuildCalendarDataSource = cnnBuildCalendarDataSource.Execute(SQLBuildCalendarDataSource)

If not rsBuildCalendarDataSource.EOF Then

	DayCount = 0
	jsonDataCalendar = ""
	
	Do While Not rsBuildCalendarDataSource.EOF
	
		MonthNam = rsBuildCalendarDataSource("MonthNam")
		MonthNum = rsBuildCalendarDataSource("MonthNum")
		MonthNum = cInt(MonthNum) - 1
		DayNum = rsBuildCalendarDataSource("DayNum")
		YearNum = rsBuildCalendarDataSource("YearNum")
		OpenClosedCloseEarly = rsBuildCalendarDataSource("OpenClosedCloseEarly")
		AlterDate=rsBuildCalendarDataSource("AlternateDeliveryDate")
		If OpenClosedCloseEarly = "Closed" Then
			ClosingTime = ""
			closingEarlyTime12Hour = ""
		Else
		
			ClosingTime = rsBuildCalendarDataSource("ClosingTime")
			ClosingTime = FormatDateTime(ClosingTime, 4)
			
			closingEarlyHour = cInt(hour(ClosingTime))
			closingEarlyMinute = cInt(minute(ClosingTime))
			
			
			If (closingEarlyHour = 0) Then
			     closingEarlyHour12Hour = 12
			     closingEarlyAMPM = "PM"
			ElseIf (closingEarlyHour > 12) Then
			     closingEarlyHour12Hour = closingEarlyHour - 12
			     closingEarlyAMPM = "PM"
			ElseIf (closingEarlyHour < 12) Then
				closingEarlyHour12Hour = closingEarlyHour
				closingEarlyAMPM = "AM"
			ElseIf (closingEarlyHour = 12) Then
				closingEarlyHour12Hour = closingEarlyHour
				closingEarlyAMPM = "PM"
			End If
			
			If closingEarlyMinute = 0 Then
				closingEarlyMinute12Hour = "00"
			Else
				closingEarlyMinute12Hour = closingEarlyMinute
			End If
			
			closingEarlyTime12Hour = closingEarlyHour12Hour & ":" & closingEarlyMinute12Hour & " " & closingEarlyAMPM
			'closingEarlyTime12Hour = FormatDateTime(closingEarlyTime12Hour, 4) 
			
		End If
		
		Description = rsBuildCalendarDataSource("Description")
		Description = Replace(Description,"'","\'")
		
		'FullDate = MonthNum & "/" & DayNum & "/" & YearNum
		'FullDate = FormatDateTime(FullDate,2)

		If DayCount = 0 Then
			jsonDataCalendar = "["
		End If

		jsonDataCalendar = jsonDataCalendar & "{businessDayID:" & DayCount & ","
		jsonDataCalendar = jsonDataCalendar & "businessDayDescription:'" & Description & "',"
		jsonDataCalendar = jsonDataCalendar & "businessDayStatus:'" & OpenClosedCloseEarly & "',"
		jsonDataCalendar = jsonDataCalendar & "closeEarlyTime:'" & closingEarlyTime12Hour & "',"
        IF NOT IsNull(AlterDate) THEN
            IF LEN(CSTR(AlterDate))>0 THEN
		        jsonDataCalendar = jsonDataCalendar & "alterDate:'" &FormatDateTime(AlterDate,2) & "',"
                ELSE
                    jsonDataCalendar = jsonDataCalendar & "alterDate:'',"
            END IF
            ELSE
            jsonDataCalendar = jsonDataCalendar & "alterDate:'',"
        END IF
		jsonDataCalendar = jsonDataCalendar & "startDate:new Date(" & YearNum & "," & MonthNum & "," & DayNum & "),"
		jsonDataCalendar = jsonDataCalendar & "endDate:new Date(" & YearNum & "," & MonthNum & "," & DayNum & ")},"

		DayCount = DayCount + 1
		rsBuildCalendarDataSource.MoveNext
		
	Loop
	
	If Len(jsonDataCalendar)>0 Then jsonDataCalendar = Left(jsonDataCalendar,Len(jsonDataCalendar)-1)
	jsonDataCalendar = jsonDataCalendar & "]"
	
	
End If

'************************************************************************************************
'Get the minimum date to show based on  the oldest year is in Settings_CompanyCalendar
'************************************************************************************************
SQLBuildCalendarDataSource = "SELECT * FROM Settings_CompanyCalendar ORDER BY YearNum ASC"

Set cnnBuildCalendarDataSource = Server.CreateObject("ADODB.Connection")
cnnBuildCalendarDataSource.open (Session("ClientCnnString"))
Set rsBuildCalendarDataSource = Server.CreateObject("ADODB.Recordset")
rsBuildCalendarDataSource.CursorLocation = 3 

Set rsBuildCalendarDataSource = cnnBuildCalendarDataSource.Execute(SQLBuildCalendarDataSource)

If not rsBuildCalendarDataSource.EOF Then
	MinYearNum = rsBuildCalendarDataSource("YearNum")
Else
	MinYearNum = Year(Date())
End If

MinDateToShow = "1/1/" & MinYearNum 
'************************************************************************************************
					
Set rsBuildCalendarDataSource = Nothing
cnnBuildCalendarDataSource.Close
Set BuildCalendarDataSource = nothing

'Response.write("<br><br><br>MinDateToShow : " & MinDateToShow)
%>


<!-- local custom css !-->
<style type="text/css">
	.form-control{
		overflow-x: hidden;
		}
		
	.post-labels{
 		padding-top: 5px;
 	}
 	.row-margin{
	 	margin-bottom: 20px;
	 	margin-top: 20px;
 	}
 	
 	h3{
	 	margin-top: 0px;
 	}
 	
 	.table-size .category{
	 	width: 35%;
	 	font-weight: normal;
 	}
 	
 	.table-size .group-name{
	 	width: 40%
 	}
 	
 	.table-size .sort-order{
	 	width: 10%;
 	}
 	
 	.table-size .display{
	 	width: 15%;
 	}
 
	 .col-line{
		 margin-bottom: 20px;
	  }
     #alterdate {    padding: 6px 12px;}
     .ui-datepicker {
   background: #ffffff;
   border: 1px solid #555;
   color: #000000;
 }

</style>
<!-- eof local custom css !-->

<h1 class="page-header"><i class="fa fa-calendar" aria-hidden="true"></i> Company Calendar</h1>


<!-- tabs start here !-->
<div class="row ">
	<div class="col-lg-12">
			<div class="row">
			<div class="col-lg-12 col-line">
				<div class="panel panel-default" style="margin:10px;">
					<div class="panel-heading">Choose dates that your company is closed or will be Close Early. By default, weekends are disabled and all weekdays are considerd open, until you click the date to change its status.</div>
					<div class="panel-body">
						<div id="calendar"></div>
					</div>
				</div>
			</div>
		</div>
		
	</div>
</div>


<script type="text/javascript">


$(function() {
    
	$('#updateCompanyCalendarModal').on('shown.bs.modal', function (e) {
	
	 	var businessDayStatus = $('#businessDayStatusHidden').val();

	 	if (businessDayStatus == 'Open') {
	 		 $("#radOpen").prop("checked",true);
	 		 $("#txtBusinessDayDescription").val('');
	 		 $("#closeEarlyTimepicker").val('');
	 		 $("#closingEarlyTimeDiv").hide();
	 	}
	 	else if (businessDayStatus == 'Closed') {
	 		$("#radClosed").prop("checked",true);
	 		$("#closingEarlyTimeDiv").hide();
	 	}
	 	else if (businessDayStatus == 'Close Early') {
	 		$("#radClosingEarly").prop("checked",true);
	 		$("#closingEarlyTimeDiv").show();
	 	}
	
	});		
	
	
	//var firstOfThisYear = new Date(new Date().getFullYear(), 0, 1);
	var firstOfThisYear = '<%= MinDateToShow %>';
	var currentYear = new Date().getFullYear();
	
	    

    $('#calendar').calendar({ 
    	/* Set options here */ 
    	
    	//This disables Saturday and Sunday as business days
    	disabledWeekDays: [0,6],
    	minDate:firstOfThisYear,
    	startYear:currentYear,
    	style: 'custom',
    	
        //This is used to color weekends gray, since they are disabled
        //It colors all other days, green, which means the company is open
        //The customDataSourceRenderer that follows will color closed and Close Early days
        customDayRenderer: function(element, date) {
        
        	var day = date.getDay();
			var isWeekend = (day == 6) || (day == 0);    // 6 = Saturday, 0 = Sunday
			
			if (isWeekend) {
                $(element).css('font-weight', 'normal');
                $(element).css('font-size', '14px');
                $(element).css('color', '#DCDCDC');
            }
            else {
            	$(element).css('font-weight', 'normal');
                $(element).css('font-size', '14px');
                $(element).css('color', '#5cb85c');

            }
        },	 
    	
    	//This will style the calendar days as red for closed and yellow for Close Early
		customDataSourceRenderer: function(element, date, events) {
		
                for(var i in events) {	                                
				    if (events[i].businessDayStatus == 'Closed') {
						$(element).css('background-color', 'red');
						$(element).css('color', 'white');
						$(element).css('border-radius', '5px');
					}   
				    if (events[i].businessDayStatus == 'Close Early') {
		                $(element).css('background-color', '#F5BB00');
		                $(element).css('color', 'white');
		                $(element).css('border-radius', '5px');
					}    
				    else if ((events[i].businessDayStatus !== 'Close Early') && (events[i].businessDayStatus !== 'Closed')){
		                $(element).css('font-weight', 'normal');
		                $(element).css('font-size', '14px');
		                $(element).css('color', '#5cb85c');
					}    												        
                }
		},
		enableContextMenu: false,
				 
		//When the user clicks on a calendar date, this function passes the needed information to the modal window.
		//A day has an "event" when it is either closed or Close Early.       
        clickDay: function(e) {
            if(e.events.length > 0) {
                for(var i in e.events) {
					
					clickedDateFormatted = moment(e.events[i].startDate).format('MM/DD/YYYY');
						                                
				    $('#updateCompanyCalendarModal input[name="businessDayID"]').val(e.events[i].businessDayID);
				    $('#updateCompanyCalendarModal #txtBusinessDayDescription').val(e.events[i].businessDayDescription);
				    $('#updateCompanyCalendarModal #selectedDate').html(clickedDateFormatted);
				    $('#updateCompanyCalendarModal input[name="dateToEdit"]').val(clickedDateFormatted);
				    $('#updateCompanyCalendarModal #closingEarlyTime').val(e.events[i].closeEarlyTime);
				    $('#updateCompanyCalendarModal input[id="closeEarlyTimepicker"]').val(e.events[i].closeEarlyTime);
				    $('#updateCompanyCalendarModal #businessDayStatus').html(e.events[i].businessDayStatus);   
                    $('#updateCompanyCalendarModal input[name="businessDayStatusHidden"]').val(e.events[i].businessDayStatus);
                    if (e.events[i].alterDate.length > 0) $("#alterdate").val(moment(e.events[i].alterDate).format('MM/DD/YYYY'));
                    else $("#alterdate").val("");
                    if (e.events[i].businessDayStatus == 'Close Early') $(".date-alter>label").html("Orders received after the cutoff time specified above should have their delivery date set to");
                    else $(".date-alter>label").html("Reschedule this day's deliveries for");
                    $(".date-alter").removeClass("hidden");
                }
            }
            else {
				    clickedDateFormatted = moment(e.date).format('MM/DD/YYYY');
				    $(".date-alter").addClass("hidden");
				    $('#updateCompanyCalendarModal #selectedDate').html(clickedDateFormatted);
				    $('#updateCompanyCalendarModal input[name="dateToEdit"]').val(clickedDateFormatted);
				    $('#updateCompanyCalendarModal input[name="businessDayStatusHidden"]').val('Open');
				    $('#updateCompanyCalendarModal #businessDayStatus').html('Open');
    

            }
            $('#updateCompanyCalendarModal').modal();
			    

        },
        
        //This function is used to show a popover div on a closed or Close Early business day
        //It will show the description, status and Close Early time (if applicable) when the user mouses over the calendar
        
        mouseOnDay: function(e) {
            if(e.events.length > 0) {
                var content = '';
                
                for(var i in e.events) {
                
                	if (e.events[i].businessDayStatus == 'Close Early') {
                	
                        content += '<div class="event-tooltip-content">'
                            + '<div class="event-description" style="color:' + e.events[i].color + '">' + e.events[i].businessDayDescription + '</div>'
                            + '<div class="event-status">' + e.events[i].businessDayStatus + ' ' + e.events[i].closeEarlyTime + '</div>';
                        if (e.events[i].alterDate.length > 0) content += '<div class="event-status">Alternate delivery date:' + e.events[i].alterDate + '</div>';

                                content += '</div>';
                	}
                	else {
                        content += '<div class="event-tooltip-content">'
                            + '<div class="event-description" style="color:' + e.events[i].color + '">' + e.events[i].businessDayDescription + '</div>'
                            + '<div class="event-status">' + e.events[i].businessDayStatus + '</div>';
                            if (e.events[i].alterDate.length > 0) content += '<div class="event-status">Alternate delivery date:' + e.events[i].alterDate + '</div>';

                                content += '</div>';
                               
                	}
                }
            
                $(e.element).popover({ 
                    trigger: 'manual',
                    container: 'body',
                    html:true,
                    content: content
                });
                
                $(e.element).popover('show');
            }
        },
        mouseOutDay: function(e) {
            if(e.events.length > 0) {
                $(e.element).popover('hide');
            }
        },
        dayContextMenu: function(e) {
            $(e.element).popover('hide');
        },
    dataSource:<%= jsonDataCalendar %>
});


});
</script>


<!-- splitter !-->
<div class="row">
	<div class="col-lg-12">
	<hr />
	</div>
</div>
<!-- eof splitter !-->


<div class="modal modal-fade" id="updateCompanyCalendarModal" style="display: none;">
	<div class="modal-dialog">
		<div class="modal-content">
			<script language="JavaScript">
			<!--

			   function validateCalendarChange()
			    {
				   var selectedBusinessDayStatus = $("input[name=radUpdatedDateStatus]:checked").val()
				   var enteredBusinessDayDesc = $("#txtBusinessDayDescription").val();
				   var enteredCloseEarlyTime = $("#closeEarlyTimepicker").val();
				   		    
			       if ((selectedBusinessDayStatus == "Closed" || selectedBusinessDayStatus == "Close Early") && enteredBusinessDayDesc == "") {
			            swal("Please enter a description for the calendar date.");
			            return false;
			       }

			       if (selectedBusinessDayStatus == "Close Early" && enteredCloseEarlyTime == "") {
			            swal("Please enter the time you will be closing early.");
			            return false;
			       }
			
			       return true;
			    }
			// -->
			</script>  
		
			<script>
			
				$(document).ready(function() {
					
					$("#closingEarlyTimeDiv").hide();
                    $("#alterdate").datepicker().next("button").button({
                    icons: {
                        primary: "glyphicon glyphicon-calendar"
                    }});

                    $("input[name='radUpdatedDateStatus']").on("click",function(){
                
                        if ($(this).val()=="Open") {
                            $(".date-alter").addClass("hidden");
                            $("#alterdate").val("");

                        }
                        else {
                            if ($(this).val()=="Closed") $(".date-alter>label").html("Reschedule this day's deliveries for");
                            else $(".date-alter>label").html("Orders received after the cutoff time specified above should have their delivery date set to");
                            $(".date-alter").removeClass("hidden");

                        }
                    });

                    $('.clockpicker').clockpicker({
					    placement: 'top',
					    align: 'left',
					    donetext: 'Done',
					    twelvehour:true
					});
						 
					$('input[type=radio][name=radUpdatedDateStatus]').on('change', function() {
					     if ($(this).val() == 'Close Early') {
					     	$("#closingEarlyTimeDiv").show();				     	
					     }
					     else if ($(this).val() == 'Open') {
					     	$("#closingEarlyTimeDiv").hide();
					     	$("#txtBusinessDayDescription").val('');
					     }
					     else if ($(this).val() == 'Closed') {
					     	$("#closingEarlyTimeDiv").hide();
						 } 
					});	  
					
				}); //end document.ready() function
			
			</script>
		
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal"><span aria-hidden="true">×</span><span class="sr-only">Close</span></button>
				<h4 class="modal-title">
					Update Company Calendar for <span id="selectedDate"></span>
				</h4>
			</div>
			<form class="form-horizontal" method="post" name="frmUpdateCompanyCalendar" id="updateCompanyCalendar" action="calendarSaveFromModal.asp" onsubmit="return validateCalendarChange()">
			<div class="modal-body">
					<input type="hidden" name="businessDayID" value="">
					<input type="hidden" name="dateToEdit" id="dateToEdit" value="">
					<input type="hidden" name="closingEarlyTime" id="closingEarlyTime" value="">
					<input type="hidden" name="businessDayStatusHidden" id="businessDayStatusHidden" value="">
					
					<div class="form-group">
						<label for="min-date" class="col-sm-4 control-label">Description</label>
						<div class="col-sm-7">
							<input name="txtBusinessDayDescription" id="txtBusinessDayDescription" type="text" class="form-control">
						</div>
					</div>
					<div class="form-group">
						<label for="min-date" class="col-sm-4 control-label">Current Status</label>
						<div class="col-sm-7">
							<span id="businessDayStatus"></span>
						</div>
					</div>
					<div class="form-group">
						<label for="min-date" class="col-sm-4 control-label">Change Status To (Open, Closed, Close Early)</label>
						<div class="col-sm-7">
				
							<div class="radio">
							  <label><input type="radio" name="radUpdatedDateStatus" value="Open" id="radOpen">Open</label>
							</div>
							<div class="radio">
							  <label><input type="radio" name="radUpdatedDateStatus" value="Closed" id="radClosed">Closed</label>
							</div>
							<div class="radio">
							  <label><input type="radio" name="radUpdatedDateStatus" value="Close Early" id="radClosingEarly">Close Early</label>
							</div>
						</div>
					</div>
					<div class="form-group" id="closingEarlyTimeDiv">
						<label for="min-date" class="col-sm-4 control-label">Early Close Time</label>
						<div class="col-sm-7">
						
							<div class="input-group clockpicker" style="width:150px">
							    <input type="text" class="form-control" name="closeEarlyTimepicker" id="closeEarlyTimepicker" value="">
							    <span class="input-group-addon">
							        <span class="glyphicon glyphicon-time"></span>
							    </span>
							</div>
						</div>
					</div>
                    <div class="form-group date-alter hidden">
                        <label for="alter-date" class="col-sm-4 control-label">Reschedule this day's deliveries for</label>

                        <div class="col-sm-7">
                            <label for="alter-date" class="col-sm-12 control-label">Alternate delivery date</label>
                            <div class=" input-group">
                                
                                <input type="text" id="alterdate" name="alterdate" class="col-md-12">
                                <span class="input-group-addon">
                                    <span class="glyphicon glyphicon-calendar"></span>
                                </span>
                            </div>
                        </div>

                       
                    </div>
			</div>
			<div class="modal-footer">
				<button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
				<button type="submit" class="btn btn-primary" id="save-event">
					Save
				</button>
			</div>
			</form>
		</div>
	</div>
</div>
<div id="context-menu">
</div>
<style>
.event-tooltip-content:not(:last-child) {
	border-bottom:1px solid #ddd;
	padding-bottom:5px;
	margin-bottom:5px;
}

.event-tooltip-content .event-title {
	font-size:18px;
}

.event-tooltip-content .event-status {
	font-size:12px;
}
</style>

<!-- eof row !-->    
<!--#include file="../../inc/footer-main.asp"-->