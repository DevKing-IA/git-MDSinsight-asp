<!--#include file="../../inc/header.asp"-->

<link rel="stylesheet" href="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.css" type="text/css">
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.js"></script>

<% 

InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")


SQL = "SELECT * FROM USER_Teams WHERE InternalRecordIdentifier = " & InternalRecordIdentifier 

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then
	TeamName = rs("TeamName")
	TeamUserNos = rs("TeamUserNos")
End If
set rs = Nothing
cnn8.close
set cnn8 = Nothing

%>

<script language="Javascript">

	$(document).ready(function() {
	
		$('#lstExistingTeamUserIDs').multiselect({
		   buttonTitle: function(options, select) {
			    var selected = '';
			    options.each(function () {
			      selected += $(this).text() + ', ';
			    });
			    return selected.substr(0, selected.length - 2);
			  },
			buttonClass: 'btn-primary btn-lg',
			buttonWidth: '425px',
			maxHeight: 400,
			dropRight:true,
			enableFiltering:true,
			filterPlaceholder:'Search',
			enableCaseInsensitiveFiltering:true,
			// possible options: 'text', 'value', 'both'
			filterBehavior:'text',
			includeFilterClearBtn:true,
			nonSelectedText:'No Team Members Selected',
			numberDisplayed: 20,
		    onChange: function() {
		        var selected = this.$select.val();
		        $("#lstSelectedNewTeamUserIDs").val(selected);
		        console.log(selected);
		        // ...
		    }
    			
	    });	
	    
		//*****************************************************************************
		//Load the bootstrap multiselect box with the current team members preselected
		//*****************************************************************************
		var data= $("#lstSelectedNewTeamUserIDs").val();
		
		if (data) {
			//Make an array
			var dataarray=data.split(",");
			// Set the value
			$("#lstExistingTeamUserIDs").val(dataarray);
			// Then refresh
			$("#lstExistingTeamUserIDs").multiselect("refresh");
		}
		//*****************************************************************************
		
			
	
		$("#txtTeamName").focusout(function() {
						
			var passedNewTeamName = $("#txtTeamName").val();
			var passedCurrTeamName = $("#txtTeamNameOrig").val();
			
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForUsers.asp",
				cache: false,
				data: "action=CheckIfTeamNameAlreadyExists&passedNewTeamName=" + encodeURIComponent(passedNewTeamName) + "&passedCurrTeamName=" + encodeURIComponent(passedCurrTeamName),
				success: function(response)
				 {
	               	 if (response == "TEAMNAMEALREADYEXISTS") {
	               	 	swal("Team Name Already Exists. Please Enter a Unique Team Name");
	               	 	$("#txtTeamName").val('');
	               	 }
				 }		
			});			
		        
		});
		
	});

    function validateTeamForm()
    {
    	
        if (document.frmEditTeam.txtTeamName.value == "") {
            swal("Team name cannot be blank.");
            return false;
        }

         if (document.frmEditTeam.lstSelectedNewTeamUserIDs.value == "") {
            swal("Please select at least 2 users for a team.");
            return false;
        }
        
       return true;
    }
    
		
</script>  


<!-- password strength meter !-->

<style type="text/css">

	.select-line{
		margin-bottom: 15px;
	}
	
	.enable-disable{
		margin-top:20px;
	}
	
	.row-line{
		margin-bottom: 25px;
	}
	
	.table th, tr, td{
		font-weight: normal;
	}
	
	.table>thead>tr>th{
		border: 0px;
	}
	.table thead>tr>th,.table tbody>tr>th,.table tfoot>tr>th,.table thead>tr>td,.table tbody>tr>td,.table tfoot>tr>td{
	border:0px;
	}
	
	.form-control{
		min-width: 100px;
	}
	
	.textarea-box{
		min-width: 260px;
	}
	
	.custom-container{
		max-width:900px;
		margin:0 auto;
	}
	
	.control-label{
		font-size:20px;
		font-weight:normal;
		padding-top:10px;
	}
	.control-label-last{
		padding-top:0px;
	}
	
	.required{
		border-left:3px solid red;
	}

	.multi-select{
		min-height: 400px;
   		min-width: 280px;
   		font-size:16px;
   		line-height:1.2em;
    }

	</style>
<!-- eof password strength meter !-->

<h1 class="page-header"><i class="fas fa-users-class"></i> Edit Existing Team</h1>

<div class="custom-container">

	<form method="POST" action="editTeam_submit.asp" name="frmEditTeam" id="frmEditTeam" onsubmit="return validateTeamForm();">

		<div class="row row-line">

			<div class="form-group col-lg-12">
				<label for="txtCondition" class="col-sm-3 control-label">Team Name</label>	
    			<div class="col-sm-6">
    				<input type="hidden" name="txtTeamNameOrig" id="txtTeamNameOrig" value="<%= TeamName %>">
    				<input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
    				<input type="text" class="form-control required" id="txtTeamName" name="txtTeamName" value="<%= TeamName %>">
    			</div>
			</div>
			
		</div>
		
		<div class="row row-line">		
			<div class="form-group col-lg-12">

				<label for="lstExistingTeamUserIDs" class="col-sm-3 control-label">Use The Dropdown to Modify Your Team</label>
				
				<div class="col-sm-9">
					<input type="hidden" name="lstSelectedNewTeamUserIDs" id="lstSelectedNewTeamUserIDs" value="<%= TeamUserNos %>">
					<select id="lstExistingTeamUserIDs" multiple="multiple" name="lstExistingTeamUserIDs">
						<%	'Get list of all users not currently archived or disabled
							
						Set cnnUserList = Server.CreateObject("ADODB.Connection")
						cnnUserList.open Session("ClientCnnString")
		
						SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 and userEnabled <> 0 ORDER BY userFirstName,userLastName"
						
						Set rsUserList = Server.CreateObject("ADODB.Recordset")
						rsUserList.CursorLocation = 3 
						Set rsUserList = cnnUserList.Execute(SQLUserList)
						
						If Not rsUserList.EOF Then
							Do While Not rsUserList.EOF
							
								FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName") & " (" & rsUserList("userDisplayName") & ")"
								Response.Write("<option value='" & rsUserList("UserNo") & "'>" & FullName & "</option>")
						
								rsUserList.MoveNext
							Loop
						End If
			
						Set rsUserList = Nothing
						cnnUserList.Close
						Set cnnUserList = Nothing
							
						%>
					</select>				
				</div>
				<!-- eof col-sm-9 -->
			</div>
			<!-- eof col-lg-12 -->	
        </div>
		<!-- eof row line -->
		
	    <!-- cancel / submit !-->
		<div class="row row-line pull-right" style="margin-right:25px;">
			<div class="col-lg-12">
				<a href="<%= BaseURL %>admin/teams/main.asp">
    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Teams List</button>
				</a>
				<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
			</div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->
