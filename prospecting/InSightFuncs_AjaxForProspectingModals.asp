<!--#include file="InSightFuncs.asp"-->
<!--#include file="InSightFuncs_Users.asp"-->
<!--#include file="InSightFuncs_Routing.asp"-->
<!--#include file="InSightFuncs_Prospecting.asp"-->
<%

'***************************************************
'List of all the AJAX functions & subs
'***************************************************
'Sub DisplayRecyclePoolByPageAndRowCount()
'Sub GetProspectBusinessCardInformationForModal()
'Sub GetProspectOwnerInformationForModal()
'Sub GetProspectCommentsInformationForModal()
'Sub GetProspectOpportunityInformationForModal()
'Sub GetProspectCurrentSupplierInformationForModal()
'Sub GetProspectCompetitorSourceInformationForModal()
'Sub GetProspectActivityInformationForModal()
'Sub GetProspectStageInformationForModal()
'Sub GetProspectDeleteInformationForModal()
'Sub GetProspectAddNotesInformationForModal()
'Sub GetInitialActivityAppmtOrMeeting()
'Sub GetAllowActivityUpdatesToUsersCalendarForModal()
'Sub GetActivityCalendarApptOrMeetingForModal()
'Sub GetMeetingLocationForModal()
'Sub CheckIfSelectedOwnerIsNotCurrentUser()
'Sub CheckIfViewNameExists()
'Sub CheckIfViewNameExistsRecyclePool()
'Sub CheckIfViewNameExistsWonPool()
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

	Case "DisplayRecyclePoolByPageAndRecCount"
		DisplayRecyclePoolByPageAndRecCount()

	Case "GetProspectBusinessCardInformationForModal"
		GetProspectBusinessCardInformationForModal()
		
	Case "GetProspectOwnerInformationForModal"
		GetProspectOwnerInformationForModal()
		
	Case "GetProspectCommentsInformationForModal"
		GetProspectCommentsInformationForModal()
		
	Case "GetProspectOpportunityInformationForModal"
		GetProspectOpportunityInformationForModal()
		
	Case "GetProspectCurrentSupplierInformationForModal"
		GetProspectCurrentSupplierInformationForModal()
		
	Case "GetProspectCompetitorSourceInformationForModal"
		GetProspectCompetitorSourceInformationForModal()
		
	Case "GetProspectActivityInformationForModal" 
		GetProspectActivityInformationForModal()
		
	Case "GetProspectStageInformationForModal"
		GetProspectStageInformationForModal()
		
	Case "GetProspectDeleteInformationForModal"
		GetProspectDeleteInformationForModal()
		
	Case "GetProspectAddNotesInformationForModal"
		GetProspectAddNotesInformationForModal()	
		
	Case "GetAllowActivityUpdatesToUsersCalendarForModal"
		GetAllowActivityUpdatesToUsersCalendarForModal()
		
	Case "GetActivityCalendarApptOrMeetingForModal"
		GetActivityCalendarApptOrMeetingForModal()
		
	Case "GetMeetingLocationForModal"
		GetMeetingLocationForModal()
		
	Case "CheckIfSelectedOwnerIsNotCurrentUser"
		CheckIfSelectedOwnerIsNotCurrentUser()
		
	Case "CheckIfViewNameExists"
		CheckIfViewNameExists()
		
	Case "CheckIfViewNameExistsRecyclePool"
		CheckIfViewNameExistsRecyclePool()
		
	Case "CheckIfViewNameExistsWonPool"
		CheckIfViewNameExistsWonPool()
		
	Case "GetInitialActivityAppmtOrMeeting"
		GetInitialActivityAppmtOrMeeting()
		
End Select


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
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




Sub DisplayRecyclePoolByPageAndRecCount()

	pageToShow = Request.Form("page")
	recsToShow = Request.Form("rows")
		

	dim fldsAll
	set fldsAll = server.createObject("Scripting.Dictionary")
	dim fldsSel
	set fldsSel = server.createObject("Scripting.Dictionary")
	
	
	fldsAll.Add "Company", "Company"
	fldsAll.Add "ActivityRecID", "Last Activity"
	fldsAll.Add "ActivityDueDate", "Last Activity Date"
	fldsAll.Add "Street", "Street Address"
	fldsAll.Add "City", "City"
	fldsAll.Add "State", "State"
	fldsAll.Add "PostalCode", "Postal Code"
	fldsAll.Add "Country", "Country"
	fldsAll.Add "LeadSourceNumber", "Lead Source"
	fldsAll.Add "StageNumber", "Stage"
	fldsAll.Add "StageChangeDate", "Stage Change Date"
	fldsAll.Add "StageChangeReason", "Stage Change Reason"
	fldsAll.Add "IndustryNumber", "Industry"
	fldsAll.Add "EmployeeRangeNumber", "Num Emp"
	fldsAll.Add "OwnerUserNo", "Owner"
	fldsAll.Add "CreatedDate", "Created Date"
	fldsAll.Add "CreatedByUserNo", "Created By"
	fldsAll.Add "TelemarketerUserNo", "Telemarketer"
	fldsAll.Add "NumberOfPantries", "Num Pantries"
	fldsAll.Add "InternalRecordIdentifier", "Prospect ID"
	fldsAll.Add "ProspectRecycle", "Recycle"
	fldsAll.Add "ProspectWatch", "Watch"
	dim key
	
	dim fs,t,userdata
	
	userdata="Company,ActivityRecID,ActivityDueDate,Street,City,State,PostalCode,Country,LeadSourceNumber,StageNumber,StageChangeDate,StageChangeReason,IndustryNumber,EmployeeRangeNumber,OwnerUserNo,CreatedDate,CreatedByUserNo,TelemarketerUserNo,NumberOfPantries,InternalRecordIdentifier,ProspectRecycle,ProspectWatch"
	Dim userdata_arr
	userdata_arr=split(userdata,",")
	
	For Each key In userdata_arr
		If fldsAll.Exists(key) Then
			fldsSel.Add key, fldsAll.item(key)
			fldsAll.Remove(key)
		End If
	Next
	


		If MUV_READ("CRMVIEWSTATE") = "optMyLeads" OR MUV_READ("CRMVIEWSTATE") = "" Then
		
			SQLRecyclePool = "SELECT *, PR_Prospects.InternalRecordIdentifier AS Expr1 FROM PR_Prospects WHERE OwnerUserNo = " & Session("UserNo") & " AND Pool = 'Dead'"
			
		ElseIf MUV_READ("CRMVIEWSTATE") = "optAllLeads" Then
			
			SQLRecyclePool = "SELECT  Top 550 *, PR_Prospects.InternalRecordIdentifier AS Expr1 FROM PR_Prospects WHERE Pool = 'Dead'"
		
		End If

		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.open (Session("ClientCnnString"))
		Set rs8 = Server.CreateObject("ADODB.Recordset")
		rs8.CursorLocation = 3 
		
		SQLRecyclePool = "dbo.DeadPoolSelect @RecsPerPage= " & recsToShow & ", @PgNum=" & pageToShow
		
		Set rs8 = cnn8.Execute(SQLRecyclePool)
		

		If not rs8.EOF Then

			Do While Not rs8.EOF
			
				xx=GetCurrentProspectActivityNumberByProspectNumber(rs8.Fields("Expr1"))
				yy=GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1"))
				
	        	Response.Write("<tr><td></td>")
	        				
				for each key in fldsSel.keys
					If key="Company" Then
						Response.Write("<td><a href='viewProspectDetailRecyclePool.asp?i=" & rs8.Fields("Expr1") & "'>" & rs8("Company") & "</a></td>")
					ElseIf key="ActivityRecID" Then
						%>
						<td><%= GetActivityByNum(GetCurrentProspectActivityNumberByProspectNumber(rs8.Fields("Expr1"))) %></td>  
						<%
					ElseIf key="ActivityDueDate" Then
					
						unformattedActivityTime = timevalue(hour(GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1"))) & ":" & minute(GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1"))))
						
						If hour(GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1"))) > 12 Then
							activityTime = hour(GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1"))) - 12  & ":" & minute(GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1"))) & " " & right(unformattedActivityTime, 2)
						Else
							activityTime = hour(GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1"))) & ":" & minute(GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1"))) & " " & right(unformattedActivityTime, 2)
						End If
								
						If Abs(DateDiff("d",GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1")),Date())) > 0 Then
							Response.Write("<td data-actduedate='" & GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1")) & "'><span class='activitydateoverdue'>" & mmddyy(GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1"))) & " </span><span class='activitytime'>" & activityTime & "</span></td>")
						Else
							Response.Write("<td data-actduedate='" & GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1")) & "'><span class='activityduedate'>" & mmddyy(GetCurrentProspectActivityDueDateByProspectNumber(rs8.Fields("Expr1"))) & " </span><span class='activitytime'>" & activityTime & "</span></td>")
						End If
						
					ElseIf key="LeadSourceNumber" Then
						Response.Write("<td>" & GetLeadSourceByNum(rs8("LeadSourceNumber")) & "</td>")
					ElseIf key="EmployeeRangeNumber" Then
						Response.Write("<td>" & GetEmployeeRangeByNum(rs8("EmployeeRangeNumber")) & "</td>")
					ElseIf key="IndustryNumber" Then
						Response.Write("<td>" & GetIndustryByNum(rs8("IndustryNumber")) & "</td>")
					ElseIf key="StageNumber" Then
						%>
						<td><%= GetStageByNum(GetProspectCurrentStageByProspectNumber(rs8("Expr1"))) %></td>  
						<%
					ElseIf key="StageChangeReason" Then
						Response.Write("<td>" & GetStageReasonByStageIntRecID(GetProspectCurrentStageIntRecIDByProspectNumber(rs8("Expr1"))) & "</td>")
					ElseIf key="StageChangeDate" Then
						Response.Write("<td>" & formatDateTime(GetProspectLastStageChangeDateByProspectNumber(rs8("Expr1")),2) & "</td>")
					ElseIf key="OwnerUserNo" Then
						Response.Write("<td>" & GetUserDisplayNameByUserNo(rs8("OwnerUserNo")) & "</td>")
					ElseIf key="CreatedDate" Then
						Response.Write("<td>" & FormatDateTime(rs8("CreatedDate"),2) & "</td>")
					ElseIf key="TelemarketerUserNo" Then
						IF rs8("TelemarketerUserNo") <> 0 Then Response.Write("<td>" & GetUserDisplayNameByUserNo(rs8("TelemarketerUserNo")) & "</td>") Else Response.Write("<td></td>")
					ElseIf key="CreatedByUserNo" Then
						Response.Write("<td>" & GetUserDisplayNameByUserNo(rs8("CreatedByUserNo")) & "</td>")
					ElseIf key="InternalRecordIdentifier" Then
						Response.Write("<td>" & rs8.Fields("Expr1") & "</td>")
					ElseIf key="ProspectRecycle" Then
						If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then
							%><td>NA</td><%
						ElseIf (cInt(GetProspectOwnerNoByNumber(rs8.Fields("Expr1"))) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then
							%><td><button type="button" class="btn btn-success btn-sm" data-toggle="modal" data-target="#myProspectingRecycleModal" data-tooltip="true" data-title="Recycle This Prospect" data-show="true"><i class="fa fa-recycle" aria-hidden="true"></i></button></td><%	
						ElseIf (cInt(GetProspectOwnerNoByNumber(rs8.Fields("Expr1"))) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then
							%><td><button type="button" class="btn btn-warning btn-sm" data-tooltip="true" data-title="You Do Not Own This Prospect"><i class="fa fa-recycle" aria-hidden="true"></i></button></td><%						
						ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then
							%><td><button type="button" class="btn btn-success btn-sm" data-toggle="modal" data-target="#myProspectingRecycleModal" data-tooltip="true" data-title="Recycle This Prospect" data-show="true"><i class="fa fa-recycle" aria-hidden="true"></i></button></td><%						
						End If							
					ElseIf key="ProspectWatch" Then
						If GetCRMPermissionLevel(Session("UserNo")) = "READONLY" Then
							%><td>NA</td><%
						ElseIf (cInt(GetProspectOwnerNoByNumber(rs8.Fields("Expr1"))) = cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then
							%><td><button type="button" class="btn btn-primary btn-sm" data-toggle="modal" data-target="#myProspectingWatchModal" data-tooltip="true" data-title="Watch This Prospect" data-show="true"><i class="fa fa-eye" aria-hidden="true"></i></button></td><%	
						ElseIf (cInt(GetProspectOwnerNoByNumber(rs8.Fields("Expr1"))) <> cInt(Session("UserNo")) AND GetCRMPermissionLevel(Session("userNo")) = "WRITEOWNED") Then
							%><td><button type="button" class="btn btn-warning btn-sm" data-tooltip="true" data-title="You Do Not Own This Prospect"><i class="fa fa-times" aria-hidden="true"></i></button></td><%						
						ElseIf GetCRMPermissionLevel(Session("userNo")) = "READWRITE" Then
							%><td><button type="button" class="btn btn-primary btn-sm" data-toggle="modal" data-target="#myProspectingWatchModal" data-tooltip="true" data-title="Watch This Prospect" data-show="true"><i class="fa fa-eye" aria-hidden="true"></i></button></td><%						
						End If							
					Else
						Response.Write("<td>" & rs8(key) & "</td>")
					End If
				next
		  		Response.Write("</tr>")
		  		
		  		rs8.movenext
		  		
		  	Loop
		  
		  End If
		  
		  Set rs8 = Nothing
		  cnn8.Close
		  Set cnn8 = Nothing

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub GetProspectBusinessCardInformationForModal()

	ProspectRecID = Request.Form("myProspectID")

	SQLProspect = "SELECT * FROM PR_Prospects where InternalRecordIdentifier = " & ProspectRecID 

	Set cnnProspect = Server.CreateObject("ADODB.Connection")
	cnnProspect.open (Session("ClientCnnString"))
	Set rsProspect = Server.CreateObject("ADODB.Recordset")
	rsProspect.CursorLocation = 3 
	Set rsProspect = cnnProspect.Execute(SQLProspect)

	If not rsProspect.EOF Then
		Company = rsProspect("Company")
		Street= rsProspect("Street")
		City= rsProspect("City")
		State= rsProspect("State")
		PostalCode = rsProspect("PostalCode")
		Country= rsProspect("Country")
		Suite= rsProspect("Floor_Suite_Room__c")							
		Website= rsProspect("Website")								
		IndustryNumber = rsProspect("IndustryNumber")	
		Industry = GetIndustryByNum(IndustryNumber)													
	End If
	set rsProspect = Nothing
	cnnProspect.close
	set cnnProspect = Nothing
	
	SQLContacts = "SELECT * FROM PR_ProspectContacts WHERE ProspectIntRecID = " & ProspectRecID & " AND PrimaryContact = 1"
	
	Set cnnContacts = Server.CreateObject("ADODB.Connection")
	cnnContacts.open (Session("ClientCnnString"))
	Set rsContacts = Server.CreateObject("ADODB.Recordset")
	rsContacts.CursorLocation = 3 
	Set rsContacts = cnnContacts.Execute(SQLContacts)
	
	If not rsContacts.EOF Then
	
	  	primarySuffix = rsContacts("Suffix")
	  	primaryFirstName = rsContacts("FirstName")
		primaryLastName = rsContacts("LastName")	
		primaryTitleNumber = rsContacts("ContactTitleNumber")
		primaryTitle = GetContactTitleByNum(primaryTitleNumber)
		primaryEmail = rsContacts("Email") 
		primaryPhone = rsContacts("Phone")
		primaryPhoneExt = rsContacts("PhoneExt")
		primaryCell = rsContacts("Cell")
		primaryFax = rsContacts("Fax")
				
	End If
	Set rsContacts = Nothing
	cnnContacts.Close
	Set cnnContacts = Nothing
	
	%>

	<script language="JavaScript">
		<!--
		
		$(document).ready(function() {
		   var phones = [{ "mask": "(###) ###-####"}];
		    $('#txtPhoneNumber').inputmask({ 
		        mask: phones, 
		        greedy: false, 
		        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
		    $('#txtCellPhoneNumber').inputmask({ 
		        mask: phones, 
		        greedy: false, 
		        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
		    $('#txtFaxNumber').inputmask({ 
		        mask: phones, 
		        greedy: false, 
		        definitions: { '#': { validator: "[0-9]", cardinality: 1}} });
			
			var curcountry	= '<%=Country%>';
			$('#txtCountry').empty();
				
		$.each(ContactCountries, function (key, ContactCountry) {
			$('#txtCountry').append('<option value="'+ContactCountry.id+'" ' + (curcountry+""==ContactCountry.id+""?'selected':'') + '>'+ContactCountry.title+'</option>');
		});		
		
		        
		});
		
		function isValidPhone(p) {
		  //var phoneRe = /^[2-9]\d{2}[2-9]\d{2}\d{4}$/;
		  //var phoneRe = /^[\+]?[(]?[0-9]{3}[)]?[-\s\.]?[0-9]{3}[-\s\.]?[0-9]{4,6}$/;
		  var phoneRe = /^(1\s|1|)?((\(\d{3}\))|\d{3})(\-|\s)?(\d{3})(\-|\s)?(\d{4})$/;
		  var digits = p.replace(/\D/g, "");
		  return phoneRe.test(digits);
		}
		
		function isValidEmail(email) 
		{
		    var re = /\S+@\S+\.\S+/;
		    return re.test(email);
		}	
	
	   function validateEditProspectBusinessCard()
	    {
	    
	       if (document.frmEditProspectBusinessCardFromModal.txtCompanyName.value == "") {
	            swal("Company name cannot be blank.");
	            return false;
	       }
	       if ((document.frmEditProspectBusinessCardFromModal.txtEmailAddress.value !== "") && (isValidEmail(document.frmEditProspectBusinessCardFromModal.txtEmailAddress.value) == false)) {
	           swal("The email address is invalid. Please enter any format like the following: anystring@anystring.anystring.");
	           return false;
	       }
	       if ((document.frmEditProspectBusinessCardFromModal.txtPhoneNumber.value !== "") && (isValidPhone(document.frmEditProspectBusinessCardFromModal.txtPhoneNumber.value) == false)) {
	           swal("The phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
	           return false;
	       }
	       if ((document.frmEditProspectBusinessCardFromModal.txtPhoneNumberExt.value !== "") && (document.frmEditProspectBusinessCardFromModal.txtPhoneNumber.value == "")) {
	           swal("A phone extension was added with no phone number. Please enter a phone number or clear the extension.");
	           return false;
	       }
	       if ((document.frmEditProspectBusinessCardFromModal.txtCellPhoneNumber.value !== "") && (isValidPhone(document.frmEditProspectBusinessCardFromModal.txtCellPhoneNumber.value) == false)) {
	           swal("The cell phone number is invalid. Please enter the number in the following format: (555) 555-5555.");
	           return false;
	       }
	       if ((document.frmEditProspectBusinessCardFromModal.txtFaxNumber.value !== "") && (isValid(document.frmEditProspectBusinessCardFromModal.txtFaxNumber.value) == false)) {
	           swal("The fax number is invalid. Please enter the number in the following format: (555) 555-5555.");
	           return false;
	       }
	          
	       return true;
	
	    }
	// -->
	</script>   
	
	<style>
		label {
			margin-top:15px;
		}
		
		.input-group .form-control, .input-group-addon, .input-group-btn {
		    display: table-cell;
		    height: 38px;
		}	
		
		.input-group-addon {
		    padding: 6px 12px;
		    font-size: 14px;
		    font-weight: 400;
		    line-height: 1;
		    color: #555;
		    text-align: center;
		    background-color: #eee;
		    border: 1px solid #ccc;
		    border-radius: 4px;
		    height: 38px;
		}
	
	</style>

	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">


              <div class="form-group">
 
	                <div class="col-sm-5">
	                  <label>Salutation, Mr., Mrs., etc.</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon"><i class="fa fa-user"></i></div>
                    		<select data-placeholder="Choose Suffix, Mr., Mrs., etc." class="C_Country_Modal form-control" id="txtSuffix" name="txtSuffix">  
                    			<option value="">Salutation, Mr., Mrs., etc.</option>  
                    			<option value="Mr." <% If primarySuffix = "Mr." Then Response.write("selected") %>>Mr.</option>
								<option value="Mrs." <% If primarySuffix = "Mrs." Then Response.write("selected") %>>Mrs.</option>
								<option value="Miss" <% If primarySuffix = "Miss" Then Response.write("selected") %>>Miss</option>
								<option value="Dr." <% If primarySuffix = "Dr." Then Response.write("selected") %>>Dr.</option>
								<option value="Ms." <% If primarySuffix = "Ms." Then Response.write("selected") %>>Ms.</option>                     
							</select>
	                    	
	                   </div>
	                </div> 
              
               </div>
               
               <br clear="all">

              <div class="form-group">
	                            
	                <div class="col-sm-6">
	                  <label>First Name</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtFirstNameIcon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtFirstName" name="txtFirstName" value="<%= primaryFirstName %>">
	                   </div>
	                </div> 
	                
	                <div class="col-sm-6">
	                  <label>Last Name</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtLastNameIcon"><i class="fa fa-user"></i></div>
	                    	<input type="text" class="form-control" id="txtLastName" name="txtLastName" value="<%= primaryLastName %>">
	                   </div>
	                </div> 
	 
               </div>
               
				<br clear="all">
				
              <div class="form-group">
	                            	                
	                <div class="col-sm-6">
	                  <label>Company Name</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtCompanyNameIcon"><i class="fa fa-briefcase"></i></div>
	                    	<input type="text" class="form-control red-line" id="txtCompanyName" name="txtCompanyName" value="<%= Company %>">
	                   </div>
	                </div> 

	                <div class="col-sm-6">
	                  <label>Job Title</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtTitleIcon"><i class="fa fa-id-card-o"></i></div>
                    		<select data-placeholder="Choose Job Title" class="C_Country_Modal form-control" id="txtTitle" name="txtTitle">  
                    			<option value="">Select Job Title</option>
                                <%If userCanEditCRMOnTheFly(Session("UserNO")) = True AND GetCRMAddEditMenuPermissionLevel(Session("UserNO")) = vbTrue Then%>
                            		<option value="-1" style="font-weight:bold"> -- Add a new Job Title -- </option>
                                <%End If%>                          
								<%
								SQLContactTitles = "SELECT *, InternalRecordIdentifier as id FROM PR_ContactTitles ORDER BY ContactTitle"
								Set cnnContactTitles = Server.CreateObject("ADODB.Connection")
								cnnContactTitles.open (Session("ClientCnnString"))
								Set rsContactTitles = Server.CreateObject("ADODB.Recordset")
								rsContactTitles.CursorLocation = 3 
								Set rsContactTitles = cnnContactTitles.Execute(SQLContactTitles)
								If not rsContactTitles.EOF Then
									Do While Not rsContactTitles.EOF
											%><option value="<%= rsContactTitles("id") %>" <% if rsContactTitles("id") = primaryTitleNumber Then Response.write("selected") %>><%= rsContactTitles("ContactTitle") %></option><%
										rsContactTitles.MoveNext						
									Loop
								End If
								Set rsContactTitles = Nothing
								cnnContactTitles.Close
								Set cnnContactTitles = Nothing
								
								%> 
							</select>
   	                   </div>
	                </div> 
	 
               </div>
               


              <div class="form-group">
	                            
	                <div class="col-sm-7">
	                  <label>Street Address</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtAddressLine1Icon"><i class="fa fa-address-card"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine1" name="txtAddressLine1" value="<%= Street %>">
	                   </div>
	                </div> 
	                
	                <div class="col-sm-5">
	                  <label>Suite, Floor #, etc.</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtAddressLine2Icon"><i class="fa fa-address-card-o"></i></div>
	                    	<input type="text" class="form-control" id="txtAddressLine2" name="txtAddressLine2" value="<%= Suite %>">
	                   </div>
	                </div> 
	                
	           </div>     
           
	           
	                


              <div class="form-group">
	                            
	                <div class="col-sm-5">
	                  <label>City</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtCityIcon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtCity" name="txtCity" value="<%= City %>">
	                   </div>
	                </div> 
	                

	                <div class="col-sm-4">
	                  <label>State</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtStateIcon"><i class="fa fa-address-book"></i></div>
                    		<select data-placeholder="Choose State" class="C_Country_Modal form-control" id="txtState" name="txtState"> 
                    			<option value="">State</option>
								<%
								If State = "NY" Then Response.Write("<option value='NY' selected>" & "New York" & "</option>") Else Response.Write("<option value='NY'>" & "New York" & "</option>")
								If State = "AA" Then Response.Write("<option selected>" & "Armed Forces Americas" & "</option>") Else Response.Write("<option>" & "Armed Forces Americas" & "</option>")
								If State = "AE" Then Response.Write("<option selected>" & "Armed Forces Europe, Middle East, & Canada" & "</option>") Else Response.Write("<option>" & "Armed Forces Europe, Middle East, & Canada" & "</option>")
								If State = "AK" Then Response.Write("<option selected>" & "Alaska" & "</option>") Else Response.Write("<option>" & "Alaska" & "</option>") 
								If State = "AL" Then Response.Write("<option selected>" & "Alabama" & "</option>") Else Response.Write("<option>" & "Alabama" & "</option>") 
								If State = "AP" Then Response.Write("<option selected>" & "Armed Forces Pacific" & "</option>") Else Response.Write("<option>" & "Armed Forces Pacific" & "</option>") 
								If State = "AR" Then Response.Write("<option selected>" & "Arkansas" & "</option>") Else Response.Write("<option>" & "Arkansas" & "</option>") 
								If State = "AS" Then Response.Write("<option selected>" & "American Samoa" & "</option>") Else Response.Write("<option>" & "American Samoa" & "</option>")
								If State = "AZ" Then Response.Write("<option selected>" & "Arizona" & "</option>") Else Response.Write("<option>" & "Arizona" & "</option>")
								If State = "CA" Then Response.Write("<option selected>" & "California" & "</option>") Else Response.Write("<option>" & "California" & "</option>")
								If State = "CO" Then Response.Write("<option selected>" & "Colorado" & "</option>") Else Response.Write("<option>" & "Colorado" & "</option>")
								If State = "CT" Then Response.Write("<option selected>" & "Connecticut" & "</option>") Else Response.Write("<option>" & "Connecticut" & "</option>")
								If State = "DC" Then Response.Write("<option selected>" & "District of Columbia" & "</option>") Else Response.Write("<option>" & "District of Columbia" & "</option>")
								If State = "DE" Then Response.Write("<option selected>" & "Delaware" & "</option>") Else Response.Write("<option>" & "Delaware" & "</option>") 
								If State = "FL" Then Response.Write("<option selected>" & "Florida" & "</option>") Else Response.Write("<option>" & "Florida" & "</option>")
								If State = "FM" Then Response.Write("<option selected>" & "Federated States of Micronesia" & "</option>") Else Response.Write("<option>" & "Federated States of Micronesia" & "</option>")
								If State = "GA" Then Response.Write("<option selected>" & "Georgia" & "</option>") Else Response.Write("<option>" & "Georgia" & "</option>")
								If State = "GU" Then Response.Write("<option selected>" & "Guam" & "</option>") Else Response.Write("<option>" & "Guam" & "</option>") 
								If State = "HI" Then Response.Write("<option selected>" & "Hawaii" & "</option>") Else Response.Write("<option>" & "Hawaii" & "</option>")
								If State = "IA" Then Response.Write("<option selected>" & "Iowa" & "</option>") Else Response.Write("<option>" & "Iowa" & "</option>") 
								If State = "ID" Then Response.Write("<option selected>" & "Idaho" & "</option>") Else Response.Write("<option>" & "Idaho" & "</option>") 
								If State = "IL" Then Response.Write("<option selected>" & "Illinois" & "</option>") Else Response.Write("<option>" & "Illinois" & "</option>") 
								If State = "IN" Then Response.Write("<option selected>" & "Indiana" & "</option>") Else Response.Write("<option>" & "Indiana" & "</option>")
								If State = "KS" Then Response.Write("<option selected>" & "Kansas" & "</option>") Else Response.Write("<option>" & "Kansas" & "</option>") 
								If State = "KY" Then Response.Write("<option selected>" & "Kentucky" & "</option>") Else Response.Write("<option>" & "Kentucky" & "</option>") 
								If State = "LA" Then Response.Write("<option selected>" & "Louisiana" & "</option>") Else Response.Write("<option>" & "Louisiana" & "</option>") 
								If State = "MA" Then Response.Write("<option selected>" & "Massachusetts" & "</option>") Else Response.Write("<option>" & "Massachusetts" & "</option>") 
								If State = "MD" Then Response.Write("<option selected>" & "Maryland" & "</option>") Else Response.Write("<option>" & "Maryland" & "</option>") 
								If State = "ME" Then Response.Write("<option selected>" & "Maine" & "</option>") Else Response.Write("<option>" & "Maine" & "</option>")
								If State = "MH" Then Response.Write("<option selected>" & "Marshall Islands" & "</option>") Else Response.Write("<option>" & "Marshall Islands" & "</option>") 
								If State = "MI" Then Response.Write("<option selected>" & "Michigan" & "</option>") Else Response.Write("<option>" & "Michigan" & "</option>") 
								If State = "MN" Then Response.Write("<option selected>" & "Minnesota" & "</option>") Else Response.Write("<option>" & "Minnesota" & "</option>") 
								If State = "MO" Then Response.Write("<option selected>" & "Missouri" & "</option>") Else Response.Write("<option>" & "Missouri" & "</option>") 
								If State = "MP" Then Response.Write("<option selected>" & "Northern Mariana Islands" & "</option>") Else Response.Write("<option>" & "Northern Mariana Islands" & "</option>")
								If State = "MS" Then Response.Write("<option selected>" & "Mississippi" & "</option>") Else Response.Write("<option>" & "Mississippi" & "</option>") 
								If State = "MT" Then Response.Write("<option selected>" & "Montana" & "</option>") Else Response.Write("<option>" & "Montana" & "</option>") 
								If State = "NC" Then Response.Write("<option selected>" & "North Carolina" & "</option>") Else Response.Write("<option>" & "North Carolina" & "</option>") 
								If State = "ND" Then Response.Write("<option selected>" & "North Dakota" & "</option>") Else Response.Write("<option>" & "North Dakota" & "</option>") 
								If State = "NE" Then Response.Write("<option selected>" & "Nebraska" & "</option>") Else Response.Write("<option>" & "Nebraska" & "</option>") 
								If State = "NH" Then Response.Write("<option selected>" & "New Hampshire" & "</option>") Else Response.Write("<option>" & "New Hampshire" & "</option>") 
								If State = "NJ" Then Response.Write("<option selected>" & "New Jersey" & "</option>") Else Response.Write("<option>" & "New Jersey" & "</option>") 
								If State = "NM" Then Response.Write("<option selected>" & "New Mexico" & "</option>") Else Response.Write("<option>" & "New Mexico" & "</option>") 
								If State = "NV" Then Response.Write("<option selected>" & "Nevada" & "</option>") Else Response.Write("<option>" & "Nevada" & "</option>") 
								If State = "OH" Then Response.Write("<option selected>" & "Ohio" & "</option>") Else Response.Write("<option>" & "Ohio" & "</option>") 
								If State = "OK" Then Response.Write("<option selected>" & "Oklahoma" & "</option>") Else Response.Write("<option>" & "Oklahoma" & "</option>") 
								If State = "OR" Then Response.Write("<option selected>" & "Oregon" & "</option>") Else Response.Write("<option>" & "Oregon" & "</option>") 
								If State = "PA" Then Response.Write("<option selected>" & "Pennsylvania" & "</option>") Else Response.Write("<option>" & "Pennsylvania" & "</option>") 
								If State = "PR" Then Response.Write("<option selected>" & "Puerto Rico" & "</option>") Else Response.Write("<option>" & "Puerto Rico" & "</option>") 
								If State = "PW" Then Response.Write("<option selected>" & "Palau" & "</option>") Else Response.Write("<option>" & "Palau" & "</option>") 
								If State = "RI" Then Response.Write("<option selected>" & "Rhode Island" & "</option>") Else Response.Write("<option>" & "Rhode Island" & "</option>") 
								If State = "SC" Then Response.Write("<option selected>" & "South Carolina" & "</option>") Else Response.Write("<option>" & "South Carolina" & "</option>") 
								If State = "SD" Then Response.Write("<option selected>" & "South Dakota" & "</option>") Else Response.Write("<option>" & "South Dakota" & "</option>") 
								If State = "TN" Then Response.Write("<option selected>" & "Tennessee" & "</option>") Else Response.Write("<option>" & "Tennessee" & "</option>") 
								If State = "TX" Then Response.Write("<option selected>" & "Texas" & "</option>") Else Response.Write("<option>" & "Texas" & "</option>")
								If State = "UT" Then Response.Write("<option selected>" & "Utah" & "</option>") Else Response.Write("<option>" & "Utah" & "</option>") 
								If State = "VA" Then Response.Write("<option selected>" & "Virginia" & "</option>") Else Response.Write("<option>" & "Virginia" & "</option>") 
								If State = "VI" Then Response.Write("<option selected>" & "Virgin Islands" & "</option>") Else Response.Write("<option>" & "Virgin Islands" & "</option>") 
								If State = "VT" Then Response.Write("<option selected>" & "Vermont" & "</option>") Else Response.Write("<option>" & "Vermont" & "</option>") 
								If State = "WA" Then Response.Write("<option selected>" & "Washington" & "</option>") Else Response.Write("<option>" & "Washington" & "</option>") 
								If State = "WV" Then Response.Write("<option selected>" & "West Virginia" & "</option>") Else Response.Write("<option>" & "West Virginia" & "</option>") 
								If State = "WI" Then Response.Write("<option selected>" & "Wisconsin" & "</option>") Else Response.Write("<option>" & "Wisconsin" & "</option>")
								If State = "WY" Then Response.Write("<option selected>" & "Wyoming" & "</option>") Else Response.Write("<option>" & "Wyoming" & "</option>") 
								If State = "AB" Then Response.Write("<option selected>" & "Alberta" & "</option>") Else Response.Write("<option>" & "Alberta" & "</option>") 
								If State = "BC" Then Response.Write("<option selected>" & "British Columbia" & "</option>") Else Response.Write("<option>" & "British Columbia" & "</option>") 
								If State = "MB" Then Response.Write("<option selected>" & "Manitoba" & "</option>") Else Response.Write("<option>" & "Manitoba" & "</option>") 
								If State = "NB" Then Response.Write("<option selected>" & "New Brunswick" & "</option>") Else Response.Write("<option>" & "New Brunswick" & "</option>") 
								If State = "NL" Then Response.Write("<option selected>" & "Newfoundland" & "</option>") Else Response.Write("<option>" & "Newfoundland" & "</option>") 
								If State = "NS" Then Response.Write("<option selected>" & "Nova Scotia" & "</option>") Else Response.Write("<option>" & "Nova Scotia" & "</option>") 
								If State = "NU" Then Response.Write("<option selected>" & "Nunavut" & "</option>") Else Response.Write("<option>" & "Nunavut" & "</option>") 
								If State = "ON" Then Response.Write("<option selected>" & "Ontario" & "</option>") Else Response.Write("<option>" & "Ontario" & "</option>")
								If State = "PE" Then Response.Write("<option selected>" & "Prince Edward Island" & "</option>") Else Response.Write("<option>" & "Prince Edward Island" & "</option>") 
								If State = "QC" Then Response.Write("<option selected>" & "Quebec" & "</option>") Else Response.Write("<option>" & "Quebec" & "</option>") 
								If State = "SK" Then Response.Write("<option selected>" & "Saskatchewan" & "</option>") Else Response.Write("<option>" & "Saskatchewan" & "</option>")
								If State = "NT" Then Response.Write("<option selected>" & "Northwest Territories" & "</option>") Else Response.Write("<option>" & "Northwest Territories" & "</option>") 
								If State = "YT" Then Response.Write("<option selected>" & "Yukon Territory" & "</option>") Else Response.Write("<option>" & "Yukon Territory" & "</option>")
								%>
							</select>	
						</div>
					</div>			

	                <div class="col-sm-3">
	                  <label>Zip Code</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtZipCodeIcon"><i class="fa fa-address-book"></i></div>
	                    	<input type="text" class="form-control" id="txtZipCode" name="txtZipCode" value="<%= PostalCode %>">
	                   </div>
	                </div> 
	 
               </div>
               
                
              <div class="form-group">
	                            	                
	                <div class="col-sm-6">
	                  <label>Choose Country</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtCountryIcon"><i class="fa fa-globe"></i></div>
                    		<select data-placeholder="Choose Country" class="C_Country_Modal form-control" id="txtCountry" name="txtCountry"> 
								<%
								If Country = "US" Then Response.Write("<option selected>" & "United States" & "</option>")  Else Response.Write ("<option>" & "United States" & "</option>")
								If Country = "CA" Then Response.Write("<option selected>" & "Canada" & "</option>")  Else Response.Write ("<option>" & "Canada" & "</option>")
								%>
							</select>

	                   </div>
	                </div> 
	                	 
               </div>
               
               <br clear="all">
              
              <div class="form-group">
	                            
	                <div class="col-sm-6">
	                  <label>Email Address</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtEmailAddressIcon"><i class="fa fa-envelope"></i></div>
	                    	<input type="text" class="form-control" id="txtEmailAddress" name="txtEmailAddress" value="<%= primaryEmail %>">
	                   </div>
	                </div> 
	                <div class="col-sm-6">
	                  <label>Company Website URL</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtWebsiteURLIcon"><i class="fa fa-globe"></i></div>
	                    	<input type="text" class="form-control" id="txtWebsiteURL" name="txtWebsiteURL" value="<%= Website %>">
	                   </div>
	                </div> 
	                
	          </div>    

              <div class="form-group">
              
	                <div class="col-sm-6">
	                  <label>Phone Number</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtPhoneNumberIcon"><i class="fa fa-phone"></i></div>
	                    	<input type="text" class="form-control" id="txtPhoneNumber" name="txtPhoneNumber" value="<%= primaryPhone %>">
	                   </div>
	                </div> 

	                <div class="col-sm-3">
	                  <label>Extension</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtPhoneNumberExtIcon"><i class="fa fa-phone"></i></div>
	                    	<input type="text" class="form-control" id="txtPhoneNumberExt" name="txtPhoneNumberExt" value="<%= primaryPhoneExt %>">
	                   </div>
	                </div> 

		      </div>
		      
              <div class="form-group">
	
	                <div class="col-sm-6">
	                  <label>Cell Phone Number</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtCellPhoneNumberIcon"><i class="fa fa-mobile"></i></div>
	                    	<input type="text" class="form-control" id="txtCellPhoneNumber" name="txtCellPhoneNumber" value="<%= primaryCell %>">
	                   </div>
	                </div> 
	                
	                <div class="col-sm-6">
	                  <label>Fax Number</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtFaxNumberIcon"><i class="fa fa-fax"></i></div>
	                    	<input type="text" class="form-control" id="txtFaxNumber" name="txtFaxNumber" value="<%= primaryFax %>">
	                   </div>
	                </div>  	 

               </div>

               <div class="form-group">
	                            
	                <div class="col-sm-8">
	                  <label>Select Industry</label>
	                  <div class="input-group">
	                    	<div class="input-group-addon" id="txtIndustryIcon"><i class="fa fa-building"></i></div>
                    		<select data-placeholder="Choose Industry" class="C_Country_Modal form-control" id="txtIndustry" name="txtIndustry"> 
                    		
						</select>


	                   </div>
	                </div> 	 
               </div>
              
                    

		</div>
	</div>
</div>

<%

End Sub

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetProspectOwnerInformationForModal()

	ProspectRecID = Request.Form("myProspectID")
	OwnerUserNo = Request.Form("myOwnerUserNo")

%>

	<div class="form-group">
		<div class="col-lg-4" style="padding-left:0px;">
			<label class="control-label" style="padding-left:0px;">Choose an <%= GetTerm("Owner") %> from the list to the right:</label>
		</div>
		<div class="col-lg-8">	
			<input type="hidden" name="txtCurrentProspectOwner"	id="txtCurrentProspectOwner" value="<%= OwnerUserNo %>">
		  	<select class="form-control-modal" name="selProspectEditOwner" id="selProspectEditOwner">
	      	<%'Owner dropdown
	      	 
      	  	SQL = "SELECT UserNo, userFirstName, userLastName, userType FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
      	  	SQL = SQL & "WHERE userArchived <> 1 and userEnabled = 1 "
      	  	SQL = SQL & "ORDER BY userFirstName, userLastName"

			Set cnn8 = Server.CreateObject("ADODB.Connection")
			cnn8.open (Session("ClientCnnString"))
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.CursorLocation = 3 
			Set rs = cnn8.Execute(SQL)
		
			If not rs.EOF Then
				Do
					FullName = rs("userFirstName") & " " & rs("userLastName")
					If cInt(rs("UserNo")) = cInt(OwnerUserNo) Then
						Response.Write("<option value='" & rs("UserNo") & "' selected>" & FullName & "</option>")
					Else
						Response.Write("<option value='" & rs("UserNo") & "'>" & FullName & "</option>")
					End If
					rs.movenext
				Loop until rs.eof
			End If
			set rs = Nothing
			cnn8.close
			set cnn8 = Nothing
			%>
			</select>
		</div>
	</div>

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetProspectCommentsInformationForModal()

	ProspectRecID = Request.Form("myProspectID")

	SQL = "SELECT * FROM PR_Prospects where InternalRecordIdentifier = " & ProspectRecID 

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)

	If not rs.EOF Then
		Comments = rs("Comments")
	End If
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	%>


	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
			  <label for="prospectEditStageNotes"><%= GetTerm("Comments") %> For This Prospect:</label>
			  <textarea class="form-control" rows="5" id="txtProspectEditComments" name="txtProspectEditComments"><%= Comments %></textarea>
			  <input type="hidden" name="txtProspectCurrentComments" id="txtProspectCurrentComments" value="<%= Comments %>">
			</div>
		</div>
	</div>

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetProspectOpportunityInformationForModal()

	ProspectRecID = Request.Form("myProspectID")
	ProspectName = GetProspectNameByNumber(ProspectRecID) 

	SQLcnnCurrentProspectInfo = "SELECT * FROM PR_Prospects where InternalRecordIdentifier = " & ProspectRecID 

	Set cnnCurrentProspectInfo = Server.CreateObject("ADODB.Connection")
	cnnCurrentProspectInfo.open (Session("ClientCnnString"))
	Set rsCurrentProspectInfo = Server.CreateObject("ADODB.Recordset")
	rsCurrentProspectInfo.CursorLocation = 3 
	Set rsCurrentProspectInfo = cnnCurrentProspectInfo.Execute(SQLcnnCurrentProspectInfo)

	If not rsCurrentProspectInfo.EOF Then	

		ProjectedGPSpend= rsCurrentProspectInfo("ProjectedGPSpend")
		NumberOfPantries = rsCurrentProspectInfo("NumberOfPantries")
		EmployeeRangeNumber = rsCurrentProspectInfo("EmployeeRangeNumber")
		NumEmployees = GetEmployeeRangeByNum(EmployeeRangeNumber)
		LeaseExpirationDate = rsCurrentProspectInfo("LeaseExpirationDate")	
		ContractExpirationDate = rsCurrentProspectInfo("ContractExpirationDate")

	End If
	set rsCurrentProspectInfo = Nothing
	cnnCurrentProspectInfo.close
	set cnnCurrentProspectInfo = Nothing

	%>
	<script>
	
	  $("#txtNumEmployees").change(function() {
  
  		if ($("#txtProjectedGPSpend").val() == '' || $("#txtProjectedGPSpend").val() == '0')
  		{
			intRecID = $("#txtNumEmployees").val();
			projGPSpend = $("#" + intRecID).val();
			$("#txtProjectedGPSpend").val(projGPSpend);
		}
	  });

	</script>		
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
			  <label for="prospectEditOpportunity">Projected GP Spend (numbers only)</label>
        	  <input type="text" class="form-control showhim" id="txtProjectedGPSpend" name="txtProjectedGPSpend" value="<%= ProjectedGPSpend %>">
   			</div>
		</div>
	</div>
		
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
			  <label for="prospectEditOpportunity">Select # Employees</label><br>
	    		<select data-placeholder="Select # Employees" class="C_Country_Modal form-control" id="txtNumEmployees" name="txtNumEmployees"> 
	    			
				</select>
  	  			<%
  	  			'Get GP Spend From Employee Range
					SQL9 = "SELECT *, Cast(LEFT(Range,CHARINDEX('-',Range)-1) as int) as Expr1 FROM PR_EmployeeRangeTable "
					SQL9 = SQL9 & "order by Expr1"

					Set cnn9 = Server.CreateObject("ADODB.Connection")
					cnn9.open (Session("ClientCnnString"))
					Set rs9 = Server.CreateObject("ADODB.Recordset")
					rs9.CursorLocation = 3 
					Set rs9 = cnn9.Execute(SQL9)
						
					If not rs9.EOF Then
						Do
							%><input type="hidden" value="<%= rs9("ProjectedGPSpend") %>" id="<%= rs9("InternalRecordIdentifier") %>"><%
							rs9.movenext
						Loop until rs9.eof
					End If
					set rs9 = Nothing
					cnn9.close
					set cnn9 = Nothing
				%>
			</div>
		</div>
	</div>
		
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
			  <label for="prospectEditOpportunity">Select # Pantries</label>
        		<select data-placeholder="Select # Pantries" class="C_Country_Modal form-control" id="txtNumPantries" name="txtNumPantries" value="<%= NumberOfPantries %>"> 
        			<option value="">Select # Pantries</option>
					<% For i = 0 To 50 %>
					  <option value="<%= i %>"  <% If i = NumberOfPantries Then Response.write("selected") %>><%= i %></option>
					<% Next %>							
				</select>
			</div>
		</div>
	</div>
		
		
	<script>

        $('#datetimepickerLeaseExpiresDate').datetimepicker({
        	useCurrent: false,
        	format: 'MM/DD/YYYY',
        	minDate:moment(),
        	ignoreReadonly: true,
        	showClear: true,
		}); 
		

        $('#datetimepickerContractExpireDate').datetimepicker({
        	useCurrent: false,
        	format: 'MM/DD/YYYY',
        	minDate:moment(),
        	ignoreReadonly: true,
        	showClear: true,
		});   
		
		$("#resetLeaseDate").click(function(){
			//$('#datetimepickerLeaseExpiresDate').val("");
			$("#datetimepickerLeaseExpiresDate").datepicker("clearDates");
		});

		$("#resetContractDate").click(function(){
			//$('#datetimepickerContractExpireDate').val("");
			$("#datetimepickerContractExpireDate").datepicker("clearDates");

		});

	</script>
	
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
				<label class="control-label" style="padding-left:0px;">Bldg. Lease Expires Date</label>							  	
                <div class="input-group date" id="datetimepickerLeaseExpiresDate" style="width:250px;">
                    <input type="text" class="form-control" name="txtLeaseExpirationDate" id="txtLeaseExpirationDate" value="<%= LeaseExpirationDate %>" readonly="readonly">
                    <span class="input-group-addon">
                        <span class="glyphicon glyphicon-calendar"></span>
                    </span>
                </div>
			</div>
		</div>
	</div>
	
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
				<label class="control-label" style="padding-left:0px;">Contract Expiration Date</label>							  	
                <div class="input-group date" id="datetimepickerContractExpireDate" style="width:250px;">
                    <input type="text" class="form-control" name="txtContractExpirationDate" id="txtContractExpirationDate" value="<%= ContractExpirationDate %>" readonly="readonly">
                    <span class="input-group-addon">
                        <span class="glyphicon glyphicon-calendar"></span>
                    </span>
                </div>
                <!--<button id="resetContractDate" type="button">Reset</button>-->
			</div>
		</div>
	</div>
<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetProspectCurrentSupplierInformationForModal()

	ProspectRecID = Request.Form("myProspectID")

	SQL = "SELECT * FROM PR_Prospects where InternalRecordIdentifier = " & ProspectRecID 

	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)

	If not rs.EOF Then
		CurrentOffering = rs("CurrentOffering")
	End If
	set rs = Nothing
	cnn8.close
	set cnn8 = Nothing
	%>

	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
			  <label for="prospectEditCurrentOffering"><%= GetTerm("Current Supplier Info") %> For This Prospect:</label>
			  <textarea class="form-control" rows="5" id="txtProspectEditCurrentOffering" name="txtProspectEditCurrentOffering"><%= CurrentOffering %></textarea>
			  <input type="hidden" name="txtProspectCurrentCurrentOffering" id="txtProspectCurrentCurrentOffering" value="<%= CurrentOffering %>">
			</div>
		</div>
	</div>
		

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetProspectCompetitorSourceInformationForModal()

	ProspectRecID = Request.Form("myProspectID")
	ProspectName = GetProspectNameByNumber(ProspectRecID) 
	PrimaryCompetitorID = GetPrimaryCompetitorIDByProspectNumber(ProspectRecID)
					
	If PrimaryCompetitorID <> "" Then
	
		PrimaryCompetitorName = GetCompetitorByNum(PrimaryCompetitorID)
	
		SQLCurrentCompetitors = "SELECT * FROM PR_ProspectCompetitors WHERE CompetitorRecID = " & PrimaryCompetitorID & " AND ProspectRecID = " &  ProspectRecID & " AND PrimaryCompetitor = 1"
		
		Set cnnCurrentCompetitors = Server.CreateObject("ADODB.Connection")
		cnnCurrentCompetitors.open (Session("ClientCnnString"))
		Set rsCurrentCompetitors = Server.CreateObject("ADODB.Recordset")
		rsCurrentCompetitors.CursorLocation = 3 
		Set rsCurrentCompetitors = cnnCurrentCompetitors.Execute(SQLCurrentCompetitors)
		
		If not rsCurrentCompetitors.EOF Then
		
			BottledWater = rsCurrentCompetitors ("BottledWater")
			FilteredWater = rsCurrentCompetitors ("FilteredWater")
			OCS = rsCurrentCompetitors ("OCS")
			OCS_Supply = rsCurrentCompetitors ("OCS_Supply")
			OfficeSupplies = rsCurrentCompetitors ("OfficeSupplies")
			Vending = rsCurrentCompetitors ("Vending")
			Micromarket = rsCurrentCompetitors ("Micromarket")
			Pantry = rsCurrentCompetitors ("Pantry")
							
		End If
		Set rsCurrentCompetitors = Nothing
		cnnCurrentCompetitors.Close
		Set cnnCurrentCompetitors = Nothing
		
		
		If BottledWater = vbTrue Then BottledWater = "Bottled Water" Else BottledWater = ""
		If FilteredWater = vbTrue Then FilteredWater = "Filtered Water" Else FilteredWater = ""
		If OCS = vbTrue Then OCS = "OCS" Else OCS = ""
		If OCS_Supply = vbTrue Then OCS_Supply = "OCS Supply" Else OCS_Supply = ""
		If OfficeSupplies = vbTrue Then OfficeSupplies = "Office Supplies" Else OfficeSupplies = ""
		If Vending = vbTrue Then Vending = "Vending" Else Vending = ""
		If Micromarket = vbTrue Then Micromarket = "Micromarkets" Else Micromarket = ""
		If Pantry = vbTrue Then Pantry = "Pantry" Else Pantry = ""
	Else

		PrimaryCompetitorName = "None Selected"
		BottledWater = ""
		FilteredWater = ""
		OCS = ""
		OCS_Supply = ""
		OfficeSupplies = ""
		Vending = ""
		Micromarket = ""
		Pantry = ""
	
	End If
	
	SQLcnnCurrentProspectInfo = "SELECT * FROM PR_Prospects where InternalRecordIdentifier = " & ProspectRecID 

	Set cnnCurrentProspectInfo = Server.CreateObject("ADODB.Connection")
	cnnCurrentProspectInfo.open (Session("ClientCnnString"))
	Set rsCurrentProspectInfo = Server.CreateObject("ADODB.Recordset")
	rsCurrentProspectInfo.CursorLocation = 3 
	Set rsCurrentProspectInfo = cnnCurrentProspectInfo.Execute(SQLcnnCurrentProspectInfo)

	If not rsCurrentProspectInfo.EOF Then	
		LeadSourceNumber = rsCurrentProspectInfo("LeadSourceNumber")
		LeadSource = GetLeadSourceByNum(LeadSourceNumber)													
		TelemarketerUserNo = rsCurrentProspectInfo("TelemarketerUserNo")
		Telemarketer = GetUserDisplayNameByUserNo(TelemarketerUserNo)
		FormerCustNum = rsCurrentProspectInfo("FormerCustNum")
		CancelDate = rsCurrentProspectInfo("CancelDate")
		
		If IsNull(CancelDate) OR DateDiff("d",CancelDate,"1/1/1900") = 0 Then CancelDate = ""
	End If
	set rsCurrentProspectInfo = Nothing
	cnnCurrentProspectInfo.close
	set cnnCurrentProspectInfo = Nothing

	%>
	
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
			  <label for="prospectEditCompetitorSource">Telemarketer</label>
	    		<select data-placeholder="Choose Telemarketer" class="C_Country_Modal form-control" id="txtTelemarketerUserNo" name="txtTelemarketerUserNo"> 
				<option value="">Select a Telemarketer</option>
		      	<%'Telemarketer dropdown
	
	      	  	SQL = "SELECT UserNo, userFirstName, userLastName, userType FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
	      	  	SQL = SQL & "WHERE userArchived <> 1 AND userEnabled = 1"
	      	  	SQL = SQL & " AND userType = 'Telemarketing' "
	      	  	SQL = SQL & "ORDER BY userFirstName, userLastName"
	
				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
			
				If not rs.EOF Then
					Do
						FullName = rs("userFirstName") & " " & rs("userLastName")
						If rs("UserNo") = TelemarketerUserNo Then
							Response.Write("<option value='" & rs("UserNo") & "' selected>" & FullName & "</option>")
						Else
							Response.Write("<option value='" & rs("UserNo") & "'>" & FullName & "</option>")
						End If
						rs.movenext
					Loop until rs.eof
				End If
				set rs = Nothing
				cnn8.close
				set cnn8 = Nothing
				%>
			</select>
			</div>
		</div>
	</div>
		
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
			  <label for="prospectEditCompetitorSource">Lead Source</label>
        		<select data-placeholder="Choose Lead Source" class="C_Country_Modal form-control" id="txtLeadSource" name="txtLeadSource"> 
            		
				</select>
			</div>
		</div>
	</div>
		
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
			  <label for="prospectEditCompetitorSource">Primary Competitor</label><br>
        		<select data-placeholder="Choose Primary Competitor" class="C_Country_Modal form-control" id="txtPrimaryCompetitor" name="txtPrimaryCompetitor"> 
            		
				</select>
			</div>
		</div>
	</div>
					
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
			  <label for="prospectEditCompetitorSource">Primary Competitor Offerings</label>
			<fieldset class="group"> 
				<ul class="checkbox"> 
				  <li><input type="checkbox" id="chkBottledWater" name="chkBottledWater" <% If BottledWater = "Bottled Water" Then Response.Write("checked") %>><label for="chkBottledWater">Bottled Water</label></li> 
				  <li><input type="checkbox" id="chkFilteredWater" name="chkFilteredWater" <% If FilteredWater = "Filtered Water" Then Response.Write("checked") %>><label for="chkFilteredWater">Filtered Water</label></li> 
				  <li><input type="checkbox" id="chkOCS" name="chkOCS" <% If OCS = "OCS" Then Response.Write("checked") %>><label for="chkOCS">OCS</label></li> 
				  <li><input type="checkbox" id="chkOCS_Supply" name="chkOCS_Supply" <% If OCS_Supply = "OCS Supply" Then Response.Write("checked") %>><label for="chkOCS_Supply">OCS Supply</label></li> 
				  <li><input type="checkbox" id="chkOfficeSupplies" name="chkOfficeSupplies" <% If OfficeSupplies = "Office Supplies" Then Response.Write("checked") %>><label for="chkOfficeSupplies">Office Supplies</label></li> 
				  <li><input type="checkbox" id="chkVending" name="chkVending" <% If Vending = "Vending" Then Response.Write("checked") %>><label for="chkVending">Vending</label></li> 
				  <li><input type="checkbox" id="chkMicroMarket" name="chkMicroMarket" <% If Micromarkets = "Micromarkets" Then Response.Write("checked") %>><label for="chkMicroMarket">Micromarket</label></li>
				  <li><input type="checkbox" id="chkPantry" name="chkPantry" <% If Pantry = "Pantry"  Then Response.Write("checked") %>><label for="chkPantry">Pantry</label></li>
				</ul> 
			</fieldset> 
			</div>
		</div>
	</div>
		
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
			  <label for="prospectEditCompetitorSource">Former Customer #</label>
			  <input type="text" class="form-control" id="txtFormerCustomerNumber" name="txtFormerCustomerNumber" value="<%= FormerCustNum %>">
			</div>
		</div>
	</div>
		
		
	<script>
	 
        $('#datetimepickerCancelDate').datetimepicker({
        	useCurrent: false,
        	format: 'MM/DD/YYYY',
        	maxDate:moment(),
        	ignoreReadonly: true,
        	showClear: true,
		}); 

	</script>
	
	<div class="row">					
		<div class="col-lg-12">	
			<div class="form-group">
				<label class="control-label" style="padding-left:0px;">Cancel Date</label>							  	
                <div class="input-group date" id="datetimepickerCancelDate" style="width:250px;">
                    <input type="text" class="form-control" name="txtFormerCustomerCancelDate" id="txtFormerCustomerCancelDate" value="<%= CancelDate %>" readonly="readonly">
                    <span class="input-group-addon">
                        <span class="glyphicon glyphicon-calendar"></span>
                    </span>
                </div>
			</div>
		</div>
	</div>
	
	<%

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************





'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetProspectActivityInformationForModal() 

	ProspectRecID = Request.Form("myProspectID")
	ActivityRecID  = Request.Form("myActivityRecID")
		
	ProspectName = GetProspectNameByNumber(ProspectRecID) 
	MaxActivityDaysWarning = GetCRMMaxActivityDaysWarning()
	MaxActivityDaysPermitted = GetCRMMaxActivityDaysPermitted()

	
	If ActivityRecID  <> "" Then
		ProspectCurrentActivity = GetCurrentProspectActivityByProspectNumber(ProspectRecID)
		ProspectCurrentActivityDate = GetCurrentProspectActivityDueDateByProspectNumber(ProspectRecID)
		ProspectCurrentActivityDate = FormatDateTime(ProspectCurrentActivityDate,2) & " " & FormatDateTime(ProspectCurrentActivityDate,3)
		
	  	SQLCurrentActivityInfo = "SELECT * FROM PR_ProspectActivities Where ProspectRecID = " & ProspectRecID & " AND " & " ActivityRecID = " & ActivityRecID 
	
		Set cnnCurrentActivityInfo = Server.CreateObject("ADODB.Connection")
		cnnCurrentActivityInfo.open (Session("ClientCnnString"))
		Set rsCurrentActivityInfo = Server.CreateObject("ADODB.Recordset")
		rsCurrentActivityInfo.CursorLocation = 3 
		Set rsCurrentActivityInfo = cnnCurrentActivityInfo.Execute(SQLCurrentActivityInfo)
		If not rsCurrentActivityInfo.EOF Then
			ProspectCurrentStageNotes = rsCurrentActivityInfo("Notes")
		End If
		set rsCurrentActivityInfo = Nothing
		cnnCurrentActivityInfo.close
		set cnnCurrentActivityInfo = Nothing
		
		ProspectCurrentStage = GetCurrentProspectActivityByProspectNumber(ProspectRecID)
		ProspectCurrentStageNotes = ""
	
		%>
		<p><strong>Company:</strong> <%= ProspectName %></p>
		<p><strong>Next Activity:</strong> <%= ProspectCurrentActivity %></p>
		<p><strong>Due Date:</strong> <%= ProspectCurrentActivityDate %></p>
		<p><strong>Notes:</strong> <%= ProspectCurrentActivityNotes %></p>	
		
		<input type="hidden" name="txtCRMMaxActivityDaysWarning" id="txtCRMMaxActivityDaysWarning" value="<%= MaxActivityDaysWarning %>">
		<input type="hidden" name="txtCRMMaxActivityDaysPermitted" id="txtCRMMaxActivityDaysPermitted" value="<%= MaxActivityDaysPermitted %>">
	<% Else %>
		<p><strong>Company:</strong> <%= ProspectName %></p>
		<p><strong>Next Activity:</strong> No Next Activity</p>
		<p><strong>Due Date:</strong> NA</p>
		<p><strong>Notes:</strong> NA</p>	
		
		<input type="hidden" name="txtCRMMaxActivityDaysWarning" id="txtCRMMaxActivityDaysWarning" value="<%= MaxActivityDaysWarning %>">
		<input type="hidden" name="txtCRMMaxActivityDaysPermitted" id="txtCRMMaxActivityDaysPermitted" value="<%= MaxActivityDaysPermitted %>">
	
	<% End If %>

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetInitialActivityAppmtOrMeeting() 

	nextActivityRecID = Request("myActivityRecID")
	
	If Not IsNumeric(nextActivityRecID) OR IsEmpty(nextActivityRecID) Then
  		SQLNextActivity = "SELECT TOP 1 * FROM PR_Activities ORDER BY Activity"
	Else
		SQLNextActivity = "SELECT TOP 1 * FROM PR_Activities WHERE InternalRecordIdentifier=" & nextActivityRecID
	End If

	Set cnnNextActivity = Server.CreateObject("ADODB.Connection")
	cnnNextActivity.open (Session("ClientCnnString"))
	Set rsNextActivity = Server.CreateObject("ADODB.Recordset")
	rsNextActivity.CursorLocation = 3 
	Set rsNextActivity = cnnNextActivity.Execute(SQLNextActivity)
	If not rsNextActivity.EOF Then
		nextActivityRecID = rsNextActivity("InternalRecordIdentifier")
		ActivityCalendarShowApptOrMeeting = GetActivityApptOrMeetingByNum(nextActivityRecID)
		Response.Write(ActivityCalendarShowApptOrMeeting)
	End If
	set rsNextActivity = Nothing
	cnnNextActivity.close
	set cnnNextActivity = Nothing

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************





'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetProspectStageInformationForModal() 

	ProspectRecID = Request.Form("myProspectID")
	StageRecID = Request.Form("myStageRecID")
	
	ProspectName = GetProspectNameByNumber(ProspectRecID) 
	
  	SQLCurrentStageInfo = "SELECT TOP 1 * FROM PR_ProspectStages Where ProspectRecID = " & ProspectRecID & " AND " & " StageRecID = " & StageRecID & " ORDER BY RecordCreationDateTime DESC"

	Set cnnCurrentStageInfo = Server.CreateObject("ADODB.Connection")
	cnnCurrentStageInfo.open (Session("ClientCnnString"))
	Set rsCurrentStageInfo = Server.CreateObject("ADODB.Recordset")
	rsCurrentStageInfo.CursorLocation = 3 
	Set rsCurrentStageInfo = cnnCurrentStageInfo.Execute(SQLCurrentStageInfo)
	If not rsCurrentStageInfo.EOF Then
		ProspectCurrentStageNotes = rsCurrentStageInfo("Notes")
	End If
	set rsCurrentStageInfo = Nothing
	cnnCurrentStageInfo.close
	set cnnCurrentStageInfo = Nothing
	

	%>
	<p><strong>Company:</strong> <%= ProspectName %></p>
	<p><strong>Current Stage:</strong> <%= GetStageByNum(StageRecID) %></p>
	<p><strong>Current Stage Notes:</strong> <%= ProspectCurrentStageNotes %></p>
	<input type="hidden" name="txtCurrentStageNo" value="<%= StageRecID %>">

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetProspectDeleteInformationForModal()

	ProspectIDArray = Split(Request.Form("prospectsArray"),",")
	
	%>

	<input type="hidden" name="prospectsArray" id="prospectsArray" value="<%= Request.Form("prospectsArray") %>">
	<div class="form-group">
		<div class="col-lg-12" style="padding-left:0px; margin-bottom:15px;">
			<label class="control-label" style="padding-left:0px;">You have selected the following prospect(s) to delete:</label>
		</div>
		<div class="col-lg-12">

	<%
	For i = 0 to uBound(ProspectIDArray)

		ProspectIDNumber = cInt(ProspectIDArray(i))
		
		Set rsDelete = Server.CreateObject("ADODB.Recordset")
		rsDelete.CursorLocation = 3 
	
		SQLDelete = "SELECT * FROM PR_Prospects WHERE InternalRecordIdentifier = " & ProspectIDNumber 		
		
		Set cnnDelete = Server.CreateObject("ADODB.Connection")
		cnnDelete.open (Session("ClientCnnString"))
		Set rsDelete = cnnDelete.Execute(SQLDelete)
		
		If NOT rsDelete.EOF Then
			Company = rsDelete("Company")
			Suite= rsDelete("Floor_Suite_Room__c")
			Street= rsDelete("Street")
			City= rsDelete("City")
			State= rsDelete("State")
			PostalCode = rsDelete("PostalCode")
			%><p><strong><%= Company %></strong>,&nbsp;<%= Street %>,&nbsp;<%= City %>&nbsp;<%= State %>,&nbsp;&nbsp;<%= PostalCode %></p><br><%
		End If
		
	Next
		
	Set rsDelete = Nothing
	cnnDelete.Close
	Set cnnDelete = Nothing

%>

		</div>
	</div>
	<div class="col-lg-12" style="padding-left:0px; margin-top:15px;">
		<label class="control-label" style="padding-left:0px;">Click the delete button below to PERMANENTLY DELETE prospect(s). This cannot be undone.</label>
	</div>

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetProspectAddNotesInformationForModal()

	ProspectIDArray = Split(Request.Form("prospectsArray"),",")
	
	%>

	<input type="hidden" name="prospectsArray" id="prospectsArray" value="<%= Request.Form("prospectsArray") %>">
	<div class="form-group">
		<div class="col-lg-12" style="padding-left:0px; margin-bottom:15px;">
			<label class="control-label" style="padding-left:0px;">You have selected the following prospect(s) to add a note:</label>
		</div>
		<div class="col-lg-12">

	<%
	For i = 0 to uBound(ProspectIDArray)

		ProspectIDNumber = cInt(ProspectIDArray(i))
		
		Set rsDelete = Server.CreateObject("ADODB.Recordset")
		rsDelete.CursorLocation = 3 
	
		SQLDelete = "SELECT * FROM PR_Prospects WHERE InternalRecordIdentifier = " & ProspectIDNumber 		
		
		Set cnnDelete = Server.CreateObject("ADODB.Connection")
		cnnDelete.open (Session("ClientCnnString"))
		Set rsDelete = cnnDelete.Execute(SQLDelete)
		
		If NOT rsDelete.EOF Then
			Company = rsDelete("Company")
			Suite= rsDelete("Floor_Suite_Room__c")
			Street= rsDelete("Street")
			City= rsDelete("City")
			State= rsDelete("State")
			PostalCode = rsDelete("PostalCode")
			%><p><strong><%= Company %></strong>,&nbsp;<%= Street %>,&nbsp;<%= City %>&nbsp;<%= State %>,&nbsp;&nbsp;<%= PostalCode %></p><br><%
		End If
		
	Next
		
	Set rsDelete = Nothing
	cnnDelete.Close
	Set cnnDelete = Nothing

%>

		</div>
	</div>
	<div class="col-lg-12" style="padding-left:0px; margin-top:15px;">
		<label class="control-label" style="padding-left:0px;">Enter your note below</label>
	</div>

<%
End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetAllowActivityUpdatesToUsersCalendarForModal() 
	
	AllowCalendarUpdate = AllowUpdatesToUsersCalendar(Session("UserNo"))
	
	Response.Write(AllowCalendarUpdate)

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetActivityCalendarApptOrMeetingForModal() 

	ActivityRecID = Request.Form("myActivityRecID")
	
	ActivityCalendarShowApptOrMeeting = GetActivityApptOrMeetingByNum(ActivityRecID)

	Response.Write(ActivityCalendarShowApptOrMeeting)

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub GetMeetingLocationForModal() 

	ProspectRecID = Request.Form("myProspectID")
	ProspectName = GetProspectNameByNumber(ProspectRecID)

	SQLProspect = "SELECT * FROM PR_Prospects WHERE InternalRecordIdentifier = " & ProspectRecID 

	Set cnnProspect = Server.CreateObject("ADODB.Connection")
	cnnProspect.open (Session("ClientCnnString"))
	Set rsProspect = Server.CreateObject("ADODB.Recordset")
	rsProspect.CursorLocation = 3 
	Set rsProspect = cnnProspect.Execute(SQLProspect)

	If not rsProspect.EOF Then
		Company = rsProspect("Company")
		Street= rsProspect("Street")
		Suite= rsProspect("Floor_Suite_Room__c")
		City= rsProspect("City")
		State= rsProspect("State")
		PostalCode = rsProspect("PostalCode")
																					
	End If
	set rsProspect = Nothing
	cnnProspect.close
	set cnnProspect = Nothing
	
	meetingLocation = Company & " " & Street & " " & Suite & " " & City & ", " & State & " " & PostalCode
	
	If meetingLocation = "" Then
		meetingLocation = ProspectName
	End If

	Response.Write(meetingLocation)

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

Sub CheckIfSelectedOwnerIsNotCurrentUser() 

	showDoNotSendEmailCheckbox = "0"
	
	ProspectRecID = Request.Form("myProspectID")
	SelectedOwnerUserNo = Request.Form("newOwnerUserNo")
	
	If cInt(SelectedOwnerUserNo) <> cInt(Session("UserNo")) Then
		showDoNotSendEmailCheckbox = "1"
	End If
	
	Response.Write(showDoNotSendEmailCheckbox)

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************



'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub CheckIfViewNameExists() 

	viewNameToCheck = Request.Form("newViewName")
	viewNameExists = "False"
	viewNameToCheck = Replace(viewNameToCheck,"'","''")

	SQLProspectViewName = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Live' AND UserReportName = '" & viewNameToCheck & "'"

	Set cnnProspectViewName = Server.CreateObject("ADODB.Connection")
	cnnProspectViewName.open(Session("ClientCnnString"))
	Set rsProspectViewName = Server.CreateObject("ADODB.Recordset")
	rsProspectViewName.CursorLocation = 3 
	Set rsProspectViewName = cnnProspectViewName.Execute(SQLProspectViewName)

	If NOT rsProspectViewName.EOF Then
		viewNameExists = "True"																		
	End If
	
	set rsProspectViewName  = Nothing
	cnnProspectViewName.close
	set cnnProspectViewName  = Nothing
	
	Response.Write(viewNameExists)

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub CheckIfViewNameExistsRecyclePool() 

	viewNameToCheck = Request.Form("newViewName")
	viewNameExists = "False"
	viewNameToCheck = Replace(viewNameToCheck,"'","''")

	SQLProspectViewName = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Dead' AND UserReportName = '" & viewNameToCheck & "'"

	Set cnnProspectViewName = Server.CreateObject("ADODB.Connection")
	cnnProspectViewName.open(Session("ClientCnnString"))
	Set rsProspectViewName = Server.CreateObject("ADODB.Recordset")
	rsProspectViewName.CursorLocation = 3 
	Set rsProspectViewName = cnnProspectViewName.Execute(SQLProspectViewName)

	If NOT rsProspectViewName.EOF Then
		viewNameExists = "True"																		
	End If
	
	set rsProspectViewName  = Nothing
	cnnProspectViewName.close
	set cnnProspectViewName  = Nothing
	
	Response.Write(viewNameExists)

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************




'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
Sub CheckIfViewNameExistsWonPool() 

	viewNameToCheck = Request.Form("newViewName")
	viewNameExists = "False"
	viewNameToCheck = Replace(viewNameToCheck,"'","''")

	SQLProspectViewName = "SELECT * FROM Settings_Reports WHERE ReportNumber = 1400 AND UserNo = " & Session("userNo") & " AND PoolForProspecting = 'Won' AND UserReportName = '" & viewNameToCheck & "'"

	Set cnnProspectViewName = Server.CreateObject("ADODB.Connection")
	cnnProspectViewName.open(Session("ClientCnnString"))
	Set rsProspectViewName = Server.CreateObject("ADODB.Recordset")
	rsProspectViewName.CursorLocation = 3 
	Set rsProspectViewName = cnnProspectViewName.Execute(SQLProspectViewName)

	If NOT rsProspectViewName.EOF Then
		viewNameExists = "True"																		
	End If
	
	set rsProspectViewName  = Nothing
	cnnProspectViewName.close
	set cnnProspectViewName  = Nothing
	
	Response.Write(viewNameExists)

End Sub
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************


'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

'END ALL AJAX MODAL SUBROUTINES AND FUNCTIONS

'********************************************************************************************************************************************************
'********************************************************************************************************************************************************
'********************************************************************************************************************************************************

%>