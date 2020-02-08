<!--#include file="../../inc/header.asp"-->
<link rel="stylesheet" href="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.css" type="text/css">
<script type="text/javascript" src="<%= BaseURL %>js/bootstrap-multiselect/bootstrap-multiselect.js"></script>


<script language="Javascript">

	$(document).ready(function() {
	
		$('#lstNewTeamUserIDs').multiselect({
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
	
		$("#txtTeamName").focusout(function() {
						
			var passedNewTeamName = $("#txtTeamName").val();
			
	    	$.ajax({
				type:"POST",
				url: "../../inc/InSightFuncs_AjaxForUsers.asp",
				cache: false,
				data: "action=CheckIfTeamNameAlreadyExists&passedNewTeamName=" + encodeURIComponent(passedNewTeamName) + "&passedCurrTeamName=''",
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
    	
        if (document.frmAddTeam.txtTeamName.value == "") {
            swal("Team name cannot be blank.");
            return false;
        }

         if (document.frmAddTeam.lstSelectedNewTeamUserIDs.value == "") {
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

<h1 class="page-header"><i class="fad fa-users-medical"></i> Create New Team</h1>

<div class="custom-container">

	<form method="POST" action="addTeam_submit.asp" name="frmAddTeam" id="frmAddTeam" onsubmit="return validateTeamForm();">

		<div class="row row-line">

			<div class="form-group col-lg-12">
				<label for="txtCondition" class="col-sm-3 control-label">Team Name</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtTeamName" name="txtTeamName">
    			</div>
			</div>
			
		</div>
		
		<div class="row row-line">		
			<div class="form-group col-lg-12">

				<label for="lstNewTeamUserIDs" class="col-sm-3 control-label">Use The Dropdown to Build Your Team</label>
				
				<div class="col-sm-9">
					<input type="hidden" name="lstSelectedNewTeamUserIDs" id="lstSelectedNewTeamUserIDs">
					<select id="lstNewTeamUserIDs" multiple="multiple" name="lstNewTeamUserIDs">
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
