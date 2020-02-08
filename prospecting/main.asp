<!--#include file="../inc/header-prospecting.asp"-->
<!--#include file="../inc/InSightFuncs_Prospecting.asp"-->
<% Server.ScriptTimeout = 9000 %>
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
.tablelength {
	padding-top:10px;
	
}


	.live-pool-header{
		background: #1b51bd;
		color:#fff;
		text-align:center;
		font-weight:bold;
	}

</style>

<style>
.autocomplete-suggestions { border: 1px solid #999; background: #FFF; overflow: auto; }
.autocomplete-suggestion { padding: 2px 5px; white-space: nowrap; overflow: hidden; }
.autocomplete-selected { background: #F0F0F0; }
.autocomplete-suggestions strong { font-weight: normal; color: #3399FF; }
.autocomplete-group { padding: 2px 5px; }
.autocomplete-group strong { display: block; border-bottom: 1px solid #000; }
</style>
<script src="/js/jquery.autocomplete.js"></script>



<div id="loadingmodal"><h1>Loading Prospects</h1></div>

<script>
 
	$(window).on('load', function (e) {
	    $('#loadingmodal').fadeOut(1000);
	})
	
</script>

<script>    
$( document ).ready(function() {
	$('#autocomplete').autocomplete({
		serviceUrl: 'mainSearchSuggestion.asp',
		onSelect: function (suggestion) {
			location.href="viewProspectDetail.asp?i="+suggestion.data					
		},
		onHint: function (hint) {
			if (hint==""){
            	$(".autocomplete-suggestions").hide();
			}
        }
		
	});
});  
</script> 

<%
'Quick rebuild of the PR_ProspectContactSearch table

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.CursorLocation = adUseClient
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")

SQL = "DELETE FROM PR_ProspectContactSearch"
Set rs= cnn8.Execute(SQL)

SQL = "INSERT INTO PR_ProspectContactSearch (ProspectIntRecID, Company, City, State, FirstName, LastName) "
SQL = SQL & "SELECT PR_Prospects.InternalRecordIdentifier, PR_Prospects.Company, PR_Prospects.City, PR_Prospects.State, "
SQL = SQL & "PR_ProspectContacts.FirstName, PR_ProspectContacts.LastName "
SQL = SQL & "FROM PR_Prospects LEFT OUTER JOIN "
SQL = SQL & "PR_ProspectContacts ON PR_ProspectContacts.ProspectIntRecID = PR_Prospects.InternalRecordIdentifier "
SQL = SQL & "WHERE PR_Prospects.Pool = 'Live'"
Set rs= cnn8.Execute(SQL)

Set rs = Nothing
cnn8.Close
Set cnn8 = Nothing


' show or hide autocomplete searchbox & get # days to Show

ShowLivePoolProspectSearchBox = False

Set cnn9 = Server.CreateObject("ADODB.Connection")
cnn9.open (Session("ClientCnnString"))

Set rs9 = cnn9.Execute("SELECT * FROM Settings_Prospecting")
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

cnn9.Close
Set cnn9 = Nothing

'****************************************************************************************************
'Read Settings_Reports To See If We Are Loading A Saved Custom Report
'****************************************************************************************************

customFilterReportName = Request.Form("selectFilteredView")
customFilterReportNameQuotes = Replace(Request.Form("selectFilteredView"),"''","'")

If customFilterReportName = "" Then 
	customFilterReportName = MUV_READ("CRMVIEWSTATE")
Else
	dummy = MUV_WRITE("CRMVIEWSTATE",customFilterReportNameQuotes)
End If

If customFilterReportName = "" Then 
	customFilterReportName = "Default"
	dummy = MUV_WRITE("CRMVIEWSTATE","Default")
End If

customFilterReportNameForSQL = Replace(customFilterReportName,"'","''")


If MUV_READ("CRMVIEWSTATE") = "Default" Then

	dateTenDaysFromNow = DateAdd("d",10, Now())
	
	'Now we have a setting for this, so possibly oover-ride the 10 days
	If IsNumeric(ProspectActivityDefaultDaysToShow) Then
		dateTenDaysFromNow = DateAdd("d",ProspectActivityDefaultDaysToShow, Now())			
	End If
	
	nextActivityStartDate = dateCustomFormat("01/01/2014")
	nextActivityEndDate = dateCustomFormat(dateTenDaysFromNow)

	SQL = "SELECT * from Settings_Reports where ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = '" & customFilterReportNameForSQL & "'"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.CursorLocation = adUseClient
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rs= cnn8.Execute(SQL)
	
	If Not rs.EOF Then
		SQL = "UPDATE Settings_Reports SET  ReportSpecificData21 = '" & nextActivityStartDate & "' "
		SQL = SQL & ", ReportSpecificData22 = '" & nextActivityEndDate & "' "
		SQL = SQL & "WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = 'Default'"
		'Response.Write("<br><br><br><br>" & SQL)
		Set rs= cnn8.Execute(SQL)
	End If
	
	Set rs = Nothing
	cnn8.Close
	Set cnn8 = Nothing

End If


'****************************************************************************************************
'Read Settings_Reports To Obtain Filters For Prospecting Grid Data
'****************************************************************************************************
SQL = "SELECT * from Settings_Reports where ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = '" & customFilterReportNameForSQL & "'"
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.CursorLocation = adUseClient
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
  	SQLCreateUserFilter = SQLCreateUserFilter & " VALUES (1400," & Session("userNo") & ",'Live'," & Session("userNo") & ",'NextActivityScheduledDateRange','" & nextActivityStartDate & "','" & nextActivityEndDate & "','" & customFilterReportNameForSQL & "')"

	Set cnnCreateUserFilter = Server.CreateObject("ADODB.Connection")
	cnnCreateUserFilter.CursorLocation = adUseClient
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

<style type="text/css">
/*
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
*/
.d-none {display:none;}
/*
table{
  margin: 0 auto;
  width: 100%;
  clear: both;
  border-collapse: collapse;
  table-layout: fixed; 
  word-wrap:break-word; 
}
*/

/*
table.dataTable tbody tr.group {background-color: #f0f0f0; border-top:2px solid #000000;}
	table.dataTable.display tbody tr.group td {
    border-top: 2px solid #000000;
}

table.dataTable tbody th,
table.dataTable tbody td {
    white-space: nowrap;
}
*/


</style>

<%

	MaxActivityDaysWarningInit = GetCRMMaxActivityDaysWarning()
	MaxActivityDaysPermittedInit = GetCRMMaxActivityDaysPermitted()



col_invisible = "4,5,6,7,8,11,12,14,15,16,17,18,19"



If showHideColumns <> "" Then
	col_invisible = "19"
	Call SetColumnInVisible("col_address",4)
	Call SetColumnInVisible("col_city",5)
	Call SetColumnInVisible("col_state",6)
	Call SetColumnInVisible("col_zip",7)
	Call SetColumnInVisible("col_owner",13)
	Call SetColumnInVisible("col_createddate",14)
	Call SetColumnInVisible("col_createdby",15)
	Call SetColumnInVisible("col_numemployees",12)
	Call SetColumnInVisible("col_telemarketer",16)
	Call SetColumnInVisible("col_numpantries",17)
	Call SetColumnInVisible("col_prospectid",18)
	Call SetColumnInVisible("col_leadsource",8)
	Call SetColumnInVisible("col_industry",11)
	Call SetColumnInVisible("col_numpantries",17)
Else
	col_invisible = "4,5,6,7,8,11,12,14,15,16,17,18,19"	
End If



Sub SetColumnInVisible(col_name,col_id)
	If InStr(showHideColumns,col_name)=0 Then 
		If col_invisible="" Then
			col_invisible = col_id
		Else
			col_invisible = col_invisible & "," & col_id
		End If
	End If
End Sub
%>

<!--#include file="mainCustomizeBuildSQLTable.asp"-->

<!-- Prospecting Custom JS -->
	<!--#include file="mainJQuery_datatable.asp"-->
<!-- End Prospecting Custom JS -->

<!-- Prospecting Custom CSS -->
	<link href="maincss_new.css" rel="stylesheet" type="text/css">
<!-- End Prospecting Custom CSS -->

<!-- datatables script-->
<script>
var datatableWidget;
function getfilters(obj) {
	
	if ($(obj).find("span.data-icon img").attr("src")=="/img/details_open.png") {
		
		$(obj).find("span.data-icon img").attr("src", "/img/details_close.png");
				
		var dataID=$(obj).closest('tr').attr("data-name");
		$("tr[data-child-value='"+dataID+"']").css("display","table-row");
	}
	else  {
		$(obj).find("span.data-icon img").attr("src", "/img/details_open.png");
		
		var dataID=$(obj).closest('tr').attr("data-name");
		$("tr[data-child-value='"+dataID+"']").css("display","none");
	}
}

$(document).ready(function() {
var groupColumn = 0;
	datatableIni(-1);
	

$('#txtresultssearch').on( 'keyup', function () {
    datatableWidget.search( this.value ).draw();
} );
	
	
	$('#ViewMode a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
		

		if (typeof(datatableWidget.clear)=="function"){
			datatableWidget.clear();
			datatableWidget.destroy();
		}
	

		if($("#ViewMode li.active > a").attr("data-grouping")=="-1") {

			datatableIni(-1);
			
			
		}
		else {
			

		datatableIni($("#ViewMode li.active > a").attr("data-grouping"));
		}
		
	});
	
	
	
	function datatableIni(type) {
		switch (type) {
			case -1:
				datatableWidget=$('#prospectTable').on('preXhr.dt', function ( e, settings, data ) {
			$(".waitdiv").removeClass("d-none");
				$
			})
			.on('xhr.dt', function ( e, settings, json, xhr ) {
				$("#txttotalrecords").html(json.recordsTotal);
				$(".waitdiv").addClass("d-none");
			} )
			.DataTable({
	        scrollY: 600,
			scrollX:true,
	        scrollCollapse: true,
	        paging: true,
			dom: "<t><'row'<'col-md-3 tablelength'l><'col-md-3'i><'col-md-6'p>>",
			ajax: "main_datatablejson.asp",
			deferRender: true,
			serverSide:true,
			procesing: true,		
			lengthMenu: [[10, 25, 50, 75, 100,  500, -1],[10, 25, 50, 75, 100,  500, "All"]],
			pageLength: 100,
			order: [[3, 'asc']],
			createdRow: function ( row, data, index ) {
				if($("#ViewMode li.active > a").attr("data-grouping")=="1") {
				//$(row).attr("data-child-value",data.id).css("display","none");
				//$(row).css("table-layout","fixed");
				 
				}
			},	
			columns: [
					{ "data": "col_checkbox"},
					{ "data": "col_company"},
					{ "data": "col_nextactivity" },
					{ "data": "col_nextactivitydate" },
					{ "data": "col_address" },
					{ "data": "col_city" },
					{ "data": "col_state" },
					{ "data": "col_zip" },
					{ "data": "col_leadsource" },
					{ "data": "col_stage" },
					{ "data": "col_stagedate" },
					{ "data": "col_industry" },
					{ "data": "col_numemployees" },
					{ "data": "col_owner" },
					{ "data": "col_createddate" },
					{ "data": "col_createdby" },
					{ "data": "col_telemarketer" },
					{ "data": "col_numpantries" },
					{ "data": "col_prospectid" },
					{ "data": "col_watch" },
					{ "data": "col_edit" },

				],
				columnDefs: [
					{"orderable": false,"targets": [0,19,20] },
					{"visible": false, "targets": [<%=col_invisible%>]},
					{"className": "text-center", "targets": [0,12,17,18,19,20]}
				],

			initComplete : function() {
			var input = $('.dataTables_filter input').unbind(),
            self = this.api(),
            $searchButton = $('<button>')
                       .text('search')
                       .click(function() {
                          self.search(input.val()).draw();
                       }),
            $clearButton = $('<button>')
                       .text('clear')
                       .click(function() {
                          input.val('');
                          $searchButton.click(); 
                       }) 
        //$('.dataTables_filter').append("&nbsp;",$searchButton, $clearButton);
//-----------------------------------------------------
		// single prospect checkbox change 
		$('[name="checksingle"]').click(function () {	
		    //uncheck "select  all", if one of the listed checkbox item is unchecked
		    if(false == $(this).prop("checked")){ //if this item is unchecked
			   if ($('[name="checksingle"]:not(:checked)').length == $('[name="checksingle"]').length){
			   	$( "#addProspectToGroupSelected").hide(); 
				$( "#addNotesToProspects").hide(); 
			   	$( "#deletedSelectedProspects").hide();
			   	$( "#exportProspects").hide();
			   }
		    }
		    
		    //check "select all" if all checkbox items are checked
		    if ($('[name="checksingle"]:checked').length == $('[name="checksingle"]').length){
			    $( "#addProspectToGroupSelected" ).show(); 
				$( "#addNotesToProspects").show(); 
			    $( "#deletedSelectedProspects" ).show();
			    $( "#exportProspects" ).show();
			}
			
			if(true == $(this).prop("checked")){ //if this item is checked
				$( "#addProspectToGroupSelected" ).show(); 
				$( "#addNotesToProspects").show(); 
				$( "#deletedSelectedProspects" ).show();
				$( "#exportProspects" ).show();
				
			}
		});	
			

		//all prospects checkbox change
		$(".dataTable #checkall").click(function () {

	        if ($(".dataTable #checkall").is(':checked')) {
	            $(".dataTable tbody input[type=checkbox]").each(function () {
	                $(this).prop("checked", true);
				    $("#addProspectToGroupSelected").show(); 
				    $("#deletedSelectedProspects").show();
		            $("#exportProspects").show();    
	            });
	
	        } else {
	            $(".dataTable tbody input[type=checkbox]").each(function () {
	                $(this).prop("checked", false);
					$("#addProspectToGroupSelected").hide();
					$("#deletedSelectedProspects").hide();
					$("#exportProspects").hide();
	                
	            });
	        }
	    });
    
//-------------------------------------------------------
		
		
    }    
	,
			
			drawCallback: function ( settings ) {
				
					
					var api = this.api();
					//if(api.column(0).visible()===false) {
						
					//	$("th.sorting").removeClass("sorting");
					//}
					
					var rows = api.rows( {page:'current'} ).nodes();
					var sumData=0;
					var last=null;
					if($("#ViewMode li.active > a").attr("data-grouping")=="1") {

					}
					
				
				
				
				
        }//end of drawcalback
		
	});
	
				break;
		
		default:
		// case else start
				datatableWidget=$('#prospectTable').on('preXhr.dt', function ( e, settings, data ) {
			$(".waitdiv").removeClass("d-none");
				$
			})
			.on('xhr.dt', function ( e, settings, json, xhr ) {
				$("#txttotalrecords").html(json.recordsTotal);
				$(".waitdiv").addClass("d-none");
			} )
			.DataTable({
	        scrollY: 600,
			scrollX:true,
	        scrollCollapse: true,
	        paging: true,
			dom: "<t><'row'<'col-md-3 tablelength'l><'col-md-3'i><'col-md-6'p>>",
			ajax: "main_datatablejson.asp?owner="+type,
			deferRender: true,
			serverSide:true,
			procesing: true,		
			lengthMenu: [[10, 25, 50, 75, 100,  500, -1],[10, 25, 50, 75, 100,  500, "All"]],
			pageLength: 100,
			order: [[3, 'asc']],	
			columns: [
					{ "data": "col_checkbox"},
					{ "data": "col_company"},
					{ "data": "col_nextactivity" },
					{ "data": "col_nextactivitydate" },
					{ "data": "col_address" },
					{ "data": "col_city" },
					{ "data": "col_state" },
					{ "data": "col_zip" },
					{ "data": "col_leadsource" },
					{ "data": "col_stage" },
					{ "data": "col_stagedate" },
					{ "data": "col_industry" },
					{ "data": "col_numemployees" },
					{ "data": "col_owner" },
					{ "data": "col_createddate" },
					{ "data": "col_createdby" },
					{ "data": "col_telemarketer" },
					{ "data": "col_numpantries" },
					{ "data": "col_prospectid" },
					{ "data": "col_watch" },
					{ "data": "col_edit" },

				],
				columnDefs: [
					{"orderable": false,"targets": [0,19,20] },
					{"visible": false, "targets": [<%=col_invisible%>]},
					{"className": "text-center", "targets": [0,12,17,18,19,20]}
				],

			initComplete : function() {
			var input = $('.dataTables_filter input').unbind(),
            self = this.api(),
            $searchButton = $('<button>')
                       .text('search')
                       .click(function() {
                          self.search(input.val()).draw();
                       }),
            $clearButton = $('<button>')
                       .text('clear')
                       .click(function() {
                          input.val('');
                          $searchButton.click(); 
                       }) 
        //$('.dataTables_filter').append("&nbsp;",$searchButton, $clearButton);
//-----------------------------------------------------
		// single prospect checkbox change 
		$('[name="checksingle"]').click(function () {	
		    //uncheck "select  all", if one of the listed checkbox item is unchecked
		    if(false == $(this).prop("checked")){ //if this item is unchecked
			   if ($('[name="checksingle"]:not(:checked)').length == $('[name="checksingle"]').length){
			   	$( "#addProspectToGroupSelected").hide(); 
				$( "#addNotesToProspects").hide(); 
			   	$( "#deletedSelectedProspects").hide();
			   	$( "#exportProspects").hide();
			   }
		    }
		    
		    //check "select all" if all checkbox items are checked
		    if ($('[name="checksingle"]:checked').length == $('[name="checksingle"]').length){
			    $( "#addProspectToGroupSelected" ).show(); 
				$( "#addNotesToProspects").show(); 
			    $( "#deletedSelectedProspects" ).show();
			    $( "#exportProspects" ).show();
			}
			
			if(true == $(this).prop("checked")){ //if this item is checked
				$( "#addProspectToGroupSelected" ).show(); 
				$( "#addNotesToProspects").show(); 
				$( "#deletedSelectedProspects" ).show();
				$( "#exportProspects" ).show();
				
			}
		});	
			

		//all prospects checkbox change
		$(".dataTable #checkall").click(function () {

	        if ($(".dataTable #checkall").is(':checked')) {
	            $(".dataTable tbody input[type=checkbox]").each(function () {
	                $(this).prop("checked", true);
				    $("#addProspectToGroupSelected").show(); 
				    $("#deletedSelectedProspects").show();
		            $("#exportProspects").show();    
	            });
	
	        } else {
	            $(".dataTable tbody input[type=checkbox]").each(function () {
	                $(this).prop("checked", false);
					$("#addProspectToGroupSelected").hide();
					$("#deletedSelectedProspects").hide();
					$("#exportProspects").hide();
	                
	            });
	        }
	    });
    
//-------------------------------------------------------
		
		
    }    
	,
			
			drawCallback: function ( settings ) {
				
					
					var api = this.api();
					//if(api.column(0).visible()===false) {
						
					//	$("th.sorting").removeClass("sorting");
					//}
					
					var rows = api.rows( {page:'current'} ).nodes();
					var sumData=0;
					var last=null;
					if($("#ViewMode li.active > a").attr("data-grouping")=="1") {

					}
					
				
				
				
				
        }//end of drawcalback
		
	});
		// case else end
		
		}
		
	
	
	}	
	

				
			

		
	
});
</script>	
<!-- eof datatbales script-->

<%


Function dateCustomFormat(passeDdate)
	x = FormatDateTime(passeDdate, 2) 
	d = split(x, "/")
	dateCustomFormat = d(2) & "-" & d(0) & "-" & d(1)
End Function


'Response.Write("<div id=""PleaseWaitPanel"">")
'Response.Write(" <br><br>This may take up to a full minute, please wait...<br><br>")
'Response.Write("<img src=""../img/loading.gif"" />")
'Response.Write("</div>")
'Response.Flush()

%>	
<h1 class="page-header"><i class="fa fa-fw fa-asterisk"></i> Prospect List
	<!-- customize !-->
	
	<div class="col pull-right"><h3 style="margin-top: 5px;">Viewing: 
	<% If MUV_READ("CRMVIEWSTATE")= "Default" Then %>
		Default View (My Leads, <%=ProspectActivityDefaultDaysToShow %> Days and Older)
	<% ElseIf MUV_READ("CRMVIEWSTATE")= "Current" Then %>
		Unsaved Filter View
	<% Else %>
		<%= MUV_READ("CRMVIEWSTATE") %>
	<% End If %>
	</h3></div>
	<!-- eof customize !-->
</h1>

<div class="container">
  <div class="row">
    <div class="col-lg-9">
    
	
	<div class="form-group pull-left" style="width:100%;">
    
        <div class="row">
      	<div class="col-xs-12 col-md-8"><input id="txtresultssearch" type="text" class="search form-control" placeholder="Type here to search within current view"></div>
      	<div class="col-xs-12 col-md-4">
        <%
		
		
		If ShowLivePoolProspectSearchBox=True Then
		%>
        <input type="text" id="autocomplete" class="form-control" placeholder="Search ALL prospects by company name">
        <%End If%>
        
        </div>
    	</div>

	    
        
	</div>
	
<!-- start datatable-->
<!-- row !-->
<div class="row">
<!-- Nav tabs -->
  <ul class="nav nav-tabs" id="ViewMode" role="tablist">
    <li role="presentation" class="active"><a href="#prospects"  role="tab" data-toggle="tab" data-grouping="-1">Prospects</a></li>
    
<%
If GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then
		
		'SQL8 = "SELECT OwnerUserNo,Count(OwnerUserNo) as TotalProspects FROM PR_Prospects WHERE Pool='Live' Group BY OwnerUserNo"
		
		SQL8 = "SELECT PR_Prospects.OwnerUserNo,Count(PR_Prospects.OwnerUserNo) as TotalProspects  "
		SQL8 = SQL8 & "  FROM PR_Prospects "		
		SQL8 = SQL8 & " Inner Join zProspectFilter_" & Session("UserNo") & " ON PR_Prospects.InternalRecordIdentifier = "
		SQL8 = SQL8 & " zProspectFilter_" & Session("UserNo") & ".InternalRecordIdentifier"		
		SQL8 = SQL8 & " INNER JOIN PR_ProspectStages ON PR_Prospects.InternalRecordIdentifier=PR_ProspectStages.ProspectRecID"
		SQL8 = SQL8 & " INNER JOIN PR_Stages ON PR_ProspectStages.StageRecID = PR_Stages.InternalRecordIdentifier"
		SQL8 = SQL8 & " INNER JOIN PR_Industries ON PR_Prospects.IndustryNumber=PR_Industries.InternalRecordIdentifier"
		SQL8 = SQL8 & " INNER JOIN PR_LeadSources ON PR_Prospects.LeadSourceNumber=PR_LeadSources.InternalRecordIdentifier"
		SQL8 = SQL8 & " WHERE PR_Prospects.Pool='Live'  Group BY PR_Prospects.OwnerUserNo"
		
		
		SQL8 = "SELECT PR_Prospects.OwnerUserNo, COUNT(PR_Prospects.OwnerUserNo) AS TotalProspects FROM PR_Prospects "
		SQL8 = SQL8 & " INNER JOIN zProspectFilter_" & Session("UserNo") & " ON PR_Prospects.InternalRecordIdentifier = zProspectFilter_" & Session("UserNo") & ".InternalRecordIdentifier "
		SQL8 = SQL8 & "	INNER JOIN (SELECT InternalrecordIdentifier, RecordCreationDateTime, ProspectRecID, StageRecID, Notes, StageChangedByUserNo "
		SQL8 = SQL8 & " FROM PR_ProspectStages "
		SQL8 = SQL8 & " WHERE (InternalrecordIdentifier IN "
		SQL8 = SQL8 & " (SELECT MAX(InternalrecordIdentifier) AS Expr1 "
		SQL8 = SQL8 & " FROM PR_ProspectStages AS PR_ProspectStages_1 "
		SQL8 = SQL8 & " GROUP BY ProspectRecID))) AS derivedtbl_1 ON PR_Prospects.InternalRecordIdentifier = derivedtbl_1.ProspectRecID INNER JOIN "
		SQL8 = SQL8 & " PR_Stages ON derivedtbl_1.StageRecID = PR_Stages.InternalRecordIdentifier INNER JOIN "
		SQL8 = SQL8 & " PR_Industries ON PR_Prospects.IndustryNumber = PR_Industries.InternalRecordIdentifier INNER JOIN "
		SQL8 = SQL8 & " PR_LeadSources ON PR_Prospects.LeadSourceNumber = PR_LeadSources.InternalRecordIdentifier "
		SQL8 = SQL8 & " WHERE (PR_Prospects.Pool = 'Live') "
		SQL8 = SQL8 & " GROUP BY PR_Prospects.OwnerUserNo "
		

'Response.Write(SQL8 & "<br>")	
	
Set rsUsers = cnn8.Execute(SQL8)
Do While Not rsUsers.EOF
'	If rsUsers("OwnerUserNo") <> Session("UserNo") Then ' no tab for yourself 
	%>    
	    <li role="presentation"><a href="#bytab_<%=rsUsers("OwnerUserNo")%>" aria-controls="profile" role="tab" data-toggle="tab" data-grouping="<%=rsUsers("OwnerUserNo")%>"><%=GetUserDisplayNameByUserNo(rsUsers("OwnerUserNo"))%> (<%=rsUsers("TotalProspects")%>)</a></li>
	<%
'	End If
rsUsers.MoveNext
Loop
rsUsers.Close
Set rsUsers = Nothing
End If
%>
  </ul>
  

  <div class="container-fluid" style="padding-top:20px;">
		<div class="row">
       
           <table id="prospectTable" class="display nowrap" width="100%">              
			<thead>
            	<tr>
				<th class="col_checkbox1 live-pool-header"><input type="checkbox" id="checkall" name="checkall" /></th>
				<th class="col_company1 live-pool-header">Company</th>
				<th class="col_nextactivity1 live-pool-header">Next Activity</th>
				<th class="col_nextactivitydate1 live-pool-header">Next Activity Date</th>
				<th class="col_address1 live-pool-header">Address</th>
				<th class="col_city1 live-pool-header">City</th>
				<th class="col_state1 live-pool-header">State</th>
				<th class="col_zip1 live-pool-header">Zip</th>
				<th class="col_leadsource1 live-pool-header">Lead Source</th>
				<th class="col_stage1 live-pool-header">Stage</th>
				<th class="col_stagedate1 live-pool-header">Stage Date</th>
				<th class="col_industry1 live-pool-header">Industry</th>
				<th class="col_numemployees1 live-pool-header">Num Emp</th>
				<th class="col_owner1 live-pool-header">Owner</th>
				<th class="col_createddate1 live-pool-header">Created Date</th>
				<th class="col_createdby1 live-pool-header">Created By</th>
				<th class="col_telemarketer1 live-pool-header">Telemarketer</th>
				<th class="col_numpantries1 live-pool-header">Pantries</th>
				<th class="col_prospectid1 live-pool-header">Prospect ID</th>
				<th class="col_watch1 live-pool-header">Watch</th>
				<th class="col_edit1 live-pool-header">Edit</th>
                </tr>
			</thead>             
			</table>
            
		</div>
    </div>









<!-- ************************************************************************** --><!-- MODALS FOR EDITING CATEGORY NOTES, MEMOS AND EQUIPMENT                     --><!-- ************************************************************************** -->
<div class="waitdiv d-none" style="position: fixed;z-index: 999999999; top: 0px; left: 0px; width: 100%; height:80%; background-color:transparent; text-align: center; padding-top: 20%; filter: alpha(opacity=0); opacity:0; "></div>
	<div id="waitdiv" class="waitdiv d-none small" style="padding-bottom: 90px;text-align: center; vertical-align:middle;padding-top:50px;background-color:#ebebeb;width:300px;height:100px;margin: 0 auto; top:40%; left:40%;position:absolute;-webkit-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); -moz-box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); box-shadow: 0 5px 10px rgba(0, 0, 0, 0.2); z-index:999999999;">
		<img src="/img/loading_gray.gif" alt="" /><br /><span id="waitmsg">Loading Prospects</span> <br />Please wait ...
</div>

</div> <!-- eof row !-->

<!-- end of datatable-->
	
	

      
       

</div>

<!-- Bootstrap table -->
<div class="col-lg-3">
  <div class="panel panel-default">
   <div class="panel-heading" id="TotalNumberOfProspects">Currently Viewing <strong id="txttotalrecords"></strong> Total Prospects</div>
	<div class="panel-body fixed-panel panel-padding">


	<div class="row">
	
		<div class="col-lg-12">
			<div class="col-lg-6">
				<button class="btn btn-indigo btn-lg btn-block" id="customizeView" data-toggle="modal" data-target=".bs-modal-show-hide-columns">
					Show/Hide Columns
				</button>
			</div>
			
			<div class="col-lg-6">
				<button class="btn btn-red btn-lg btn-block" id="filterView" data-toggle="modal" data-target=".bs-modal-filter-prospecting-data">
					<i class="fa fa-filter " aria-hidden="true"></i>&nbsp;Filter View
				</button>			
			</div>
			
		</div>
		
	</div>
	
	<div class="row">

		<% If UseSettings_Reports = True OR MUV_READ("CRMSTARTDATE") <> "" OR MUV_READ("CRMENDDATE") <> "" Then %>
			<a href="mainCustomizeClearValues.asp"><button class="btn btn-warning btn-lg btn-block" id="resetView">
				<i class="fa fa-refresh" aria-hidden="true"></i>&nbsp;Clear Filters &amp; Reset View
			</button></a>
		<% End If %>
		
	</div>
	
	
		
	<form action="<%= BaseURL %>prospecting/main.asp" method="POST" name="frmToggleLeadAndDateView" id="frmToggleLeadAndDateView">
	
	<div class="row">		
		<% If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then %>
	 		<hr class="style7">
	   	<% Else
	   	
	   		'If (GetCRMPermissionLevel(Session("userNo")) <> "NONE") AND (cInt(GetProspectOwnerNoByNumber(InternalRecordIdentifier)) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then %>
	   		<hr class="style7">
		    <button type="button" class="btn btn-success btn-lg btn-block" style="height:100px; font-size:32px; border-radius:7px;" onclick="location.href='addProspectPrequalify.asp';">
		        <i class="fa fa-user"></i>&nbsp;Create New Prospect
		    </button>
		    <hr class="style7">
	   	<% End If %>
	</div>

	<% If userIsAdmin(Session("userNo")) = True OR userIsInsideSalesManager(Session("userNo")) = True OR userIsOutsideSalesManager(Session("userNo")) = True Then %>
		<div class="row" style="margin-top:-15px;">
		    <button type="button" class="btn btn-darkblue btn-lg btn-block" onclick="location.href='mainCreateViewAllProspectsFilter.asp';">
		        <i class="fa fa-users"></i>&nbsp;View All Prospects
		    </button>
		</div>
	<% Else %>
		<div class="row" style="margin-top:-15px;">
		    <button type="button" class="btn btn-lightgreen btn-lg btn-block" onclick="location.href='mainCreateViewAllProspectsFilter.asp';">
		        <i class="fa fa-users"></i>&nbsp;View All My Prospects
		    </button>
		</div>
	<% End If %>
	
	<div class="row">
	
			<hr class="style7">
	      	<%'Report View Name Dropdown
	      	
	      	userHasSavedViews = false
	      	 
	  	  	SQL = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName <> 'Current'  AND UserReportName <> 'Default' AND UserReportName <> 'All Prospects' ORDER BY UserReportName "
	
			Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.CursorLocation = adUseClient
			cnn8.open (Session("ClientCnnString"))
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3 
			Set rs = cnn8.Execute(SQL)
		
			If NOT rs.EOF Then
				
				userHasSavedViews = true
			%>
				<!-- Display Report View Names -->
				<select class="form-control when-line" style="width:100%;height:50px;display:inline;margin-left:0px;" name="selectFilteredView" id="selectFilteredView" onchange="$('#loadingmodal').show();this.form.submit()">
				<option value=""> -- Select Custom View -- </option>
				<option value="Default" <% If MUV_READ("CRMVIEWSTATE") = "Default" Then Response.Write("selected") %>>Default View (My Leads, <%=ProspectActivityDefaultDaysToShow %> Days and Older)</option>
				<%
					Do
						selReportName = Replace(rs("UserReportName"),"''","'")
						If MUV_READ("CRMVIEWSTATE") = selReportName Then
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
			    <button type="button" class="btn btn-primary btn-lg btn-block" data-toggle="modal" data-target="#saveAsNewProspectFilterView" id="btnSaveAsNewProspectFilterView">
			        <i class="far fa-save"></i>&nbsp;Save This View
			    </button>
			</div>
			
			<div class="row" style="margin-bottom:15px;">
				<div class="col-lg-12">
					<div class="col-lg-6">
						<button type="button"  class="btn btn-darkgray btn-lg btn-block" data-toggle="modal" data-target="#editFilterViewName" id="btnRenameProspectFilterView">
							<i class="fa fa-pencil-square" aria-hidden="true"></i>&nbsp;Rename View
						</button>
					</div>
					
					<div class="col-lg-6">
						<button type="button" class="btn btn-red btn-lg btn-block" data-toggle="modal" data-target="#deleteProspectView" id="btnDeleteProspectFilterView">
							<i class="fa fa-trash-o" aria-hidden="true"></i>&nbsp;Delete View
						</button>					    
					</div>
				</div>
				
			</div>
		<% Else %>
		
			<div class="row" style="margin-bottom:15px;">
			    <button type="button" class="btn btn-primary btn-lg btn-block" data-toggle="modal" data-target="#saveAsNewProspectFilterView" id="btnSaveAsNewProspectFilterView">
			        <i class="far fa-save"></i>&nbsp;Save This View
			    </button>
			</div>
		
		<% End If %>
	
	<% End If %>
	

	<% 'If UserIsAdmin(Session("UserNo")) AND GetCRMPermissionLevel(Session("UserNo")) <> "READONLY" Then %>
	<% If GetCRMDeleteProspectPermissionLevel(Session("UserNo")) = vbTrue Then %>
		<div class="row">
				<hr class="style7">
				
				<button type="button" class="btn btn-danger btn-lg btn-block" id="deletedSelectedProspects" data-toggle="modal" data-target="#myProspectingDeleteModal" data-tooltip="true" data-title="Delete Prospect(s)" style="display:none; margin-bottom:20px;">
					<i class="fa fa-trash"></i>&nbsp;Delete Prospect(s)
				</button>
		</div>

	<% End If %>
	
    <% If userCanEditCRMOnTheFly(Session("UserNO")) = True AND GetCRMAddEditMenuPermissionLevel(Session("UserNO")) = vbTrue Then %>
		<div class="row">
		

				<button type="button" class="btn btn-success btn-lg btn-block" id="addNotesToProspects" data-toggle="modal" data-target="#myProspectingAddMultipleNotesModal" data-tooltip="true" data-title="Add Note To Prospect(s)" style="display:none; margin-bottom:10px;">
					<i class="fa fa-sticky-note-o"></i>&nbsp;Add Note To Prospect(s)
				</button>
		</div>
	<% End If %>
	
	
    <% If userCanEditCRMOnTheFly(Session("UserNO")) = True AND GetCRMAddEditMenuPermissionLevel(Session("UserNO")) = vbTrue Then %>
		<div class="row">
			<button type="button" class="btn btn-info btn-lg btn-block" id="exportProspects" data-toggle="modal" data-target="#myProspectingExportModal" data-tooltip="true" data-title="Export Prospect(s)" style="display:none; margin-bottom:20px;">
				<i class="fa fa-trash"></i>&nbsp;Export Prospect(s)
			</button>
		</div>
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
<!--#include file="mainModals.asp"-->
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


 