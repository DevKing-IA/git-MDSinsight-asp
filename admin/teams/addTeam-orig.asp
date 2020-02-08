<!--#include file="../../inc/header.asp"-->


<link rel="stylesheet" type="text/css" href="<%= BaseURL %>css/tabs.css" />


<script language="Javascript">

	$(document).ready(function() {
	
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
    

	function listbox_addTeamMemberUser() {
	
		var src = document.getElementById("lstNewTeamUserIDs");
		var dest = document.getElementById("lstSelectedNewTeamUserIDs");
	
		for(var count=0; count < src.options.length; count++) {
	
			if(src.options[count].selected == true) {
					var option = src.options[count];
	
					var newOption = document.createElement("option");
					newOption.value = option.value;
					newOption.text = option.text;
					newOption.selected = true;
					try {
							 dest.add(newOption, null); //Standard
							 src.remove(count, null);
					 }catch(error) {
							 dest.add(newOption); // IE only
							 src.remove(count);
					 }
					count--;
			}
		}
		$("#lstSelectedNewTeamUserIDs").sortSelectByText();
	}	
	
	function listbox_removeTeamMemberUser() {
	
		var src = document.getElementById("lstSelectedNewTeamUserIDs");
		var dest = document.getElementById("lstNewTeamUserIDs");
	
		for(var count=0; count < src.options.length; count++) {
	
			if(src.options[count].selected == true) {
					var option = src.options[count];
	
					var newOption = document.createElement("option");
					newOption.value = option.value;
					newOption.text = option.text;
					newOption.selected = true;
					try {
							 dest.add(newOption, null); //Standard
							 src.remove(count, null);
					 }catch(error) {
							 dest.add(newOption); // IE only
							 src.remove(count);
					 }
					count--;
			}
		}
		$("#lstNewTeamUserIDs").sortSelectByText();
	}
		
	function doSubmit() {
	
		$('#lstSelectedNewTeamUserIDs option').prop('selected', true);
		return validateTeamForm();		
		$("#frmAddTeam").submit();
	
	}
	
	//********************************************	
	
	//**************** sort listbox's items
	$.fn.sortSelectByText = function(){
	    this.each(function(){
	        var selected = $(this).val(); 
	        var opts_list = $(this).find('option');
	        opts_list.sort(function(a, b) { return $(a).text() > $(b).text() ? 1 : -1; });
	        $(this).html('').append(opts_list);
	        $(this).val(selected); 
	    })
	    return this;        
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

	<form method="POST" action="addTeam_submit.asp" name="frmAddTeam" id="frmAddTeam">

		<div class="row row-line">

			<div class="form-group col-lg-12">
				<label for="txtCondition" class="col-sm-3 control-label">Team Name</label>	
    			<div class="col-sm-6">
    				<input type="text" class="form-control required" id="txtTeamName" name="txtTeamName">
    			</div>
			</div>
			
		</div>
		
		<div class="row row-line">		
			<div class="col-lg-4 line-full">
				<h4>Master User List</h4>
				<select multiple class="form-control multi-select" id="lstNewTeamUserIDs" name="lstNewTeamUserIDs">
					<%	'Get list of all users not currently archived or disabled
						
					Set cnnUserList = Server.CreateObject("ADODB.Connection")
					cnnUserList.open Session("ClientCnnString")
	
					SQLUserList = "SELECT * FROM tblUsers WHERE userArchived <> 1 ORDER BY userFirstName,userLastName"
					
					Set rsUserList = Server.CreateObject("ADODB.Recordset")
					rsUserList.CursorLocation = 3 
					Set rsUserList = cnnUserList.Execute(SQLUserList)
					
					If Not rsUserList.EOF Then
						Do While Not rsUserList.EOF
						
							FullName = rsUserList("userFirstName") & " " & rsUserList("userLastName")
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
				<!-- eof list of all users -->
	        	
	        	
	            <!-- add / remove -->
	            <div class="col-lg-3 line-full" style="text-align:center">
	                <a href="javascript:void(0)" onclick="javascript:listbox_addTeamMemberUser()"><button type="button" class="btn btn-success" style="margin-bottom:10px;margin-top:40px;margin-left:18px;">Add User To Team <i class="fa fa-arrow-right" aria-hidden="true"></i></button></a>
					<a href="javascript:void(0)" onclick="javascript:listbox_removeTeamMemberUser()"><button type="button" class="btn btn-danger" style="margin-left:3px;"><i class="fa fa-arrow-left" aria-hidden="true"></i> Remove User From Team</button></a>
	            </div>
	            <!-- eof add / remove -->
	            
	            	
				<!-- list of Selected users -->
				<div class="col-lg-4 line-full">
					<h4>Team Members</h4>
					<select multiple class="form-control multi-select" id="lstSelectedNewTeamUserIDs" name="lstSelectedNewTeamUserIDs">
					</select>
				</div>
				<!-- eof list of Selected users-->
            
         <!-- eof line -->
         </div>

		
	    <!-- cancel / submit !-->
		<div class="row row-line pull-right" style="margin-right:25px;">
			<div class="col-lg-12">
				<a href="<%= BaseURL %>admin/teams/main.asp">
    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Teams List</button>
				</a>
				<button type="button" onclick="javascript:doSubmit();" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
			</div>
		</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->
