<!--#include file="../../inc/header.asp"-->

<SCRIPT LANGUAGE="JavaScript">
<!--
    function checkform()
    {

        if (document.getElementById('txtAlertName').value == ''){
 	        swal("Please enter a name for this alert.")
            return false;
        }
	
        if (document.getElementById('selCond').selectedIndex != '0'){
 	        swal("Please select a condition for this alert to occur.")
            return false;
        }
	
	
        if (document.getElementById('selCond').selectedIndex == '3'){

	   		var minut = document.getElementById('selLimitMinutes').value;
			var maxim = document.getElementById('selLimitMaxTimes').value;
	
	       	 if ( minut * maxim > 1200){
		        swal("The combination of [minutes to wait] * [max times to send] cannot exceed 20 hours. Please adjust your entries before saving.")
	            return false;
	         }         
        }

        return true;

    }
    
   
// -->
</SCRIPT>   

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
	
	.when-col{
		width: 10%;
	}
	
	.reference-col{
		width: 45%;
	}
	
	.has-more-col{
		width: 12%;
	}
	
	.form-control{
		min-width: 100px;
	}
	
	.textarea-box{
		min-width: 260px;
	}
	
	.required {
	  position: absolute;
	  top: 3px;
	  right: 16px;
	  font-weight: bold;
	  color: #f44545;
	  font-size: 10px;
	  text-transform: uppercase;
	}

	.alert-checkbox{
		margin-top: 15px;
	}
		
	.when-line{
		margin-top: 20px;
	}
	
	.when-line strong{
		padding-top: 7px;
		display: inline-block;
	}


	.email-alert-line{
		margin-top: 20px;
	}
	
	.incl-message-input {
	    border: 1px solid #ccc;
	    box-shadow: 0px 0px 0px 0px;
	    max-width: 160px;
	}

  .email-multi-select{
	  min-height: 160px;
  }
  
  .limit-alerts{
	  display: inline-block;
	  padding-top: 15px;
  }
		
</style>


<h1 class="page-header"><i class="fa fa-fw fa-exclamation"></i>New Service Ticket Alert (Other Conditions)</h1>


<form method="POST" action="addAlertServiceOtherConditions_submit.asp"  onsubmit="return checkform();" name="frmAddAlert" id="frmAddAlert" class="uk-form bt-flabels js-flabels">


	<div class="row row-line">
	
		<div class="col-lg-2">
			<strong>Alert Name</strong><input type="text" id="txtAlertName" name="txtAlertName" class="form-control">
			<span class="required">Required</span>
		</div>

		<!-- enabled !-->
		<div class="col-lg-1 alert-checkbox">
			<div class="checkbox">
				<label>
					<input type="checkbox" id="chkEnabled" name="chkEnabled" checked="checked"><strong> Enabled</strong>
				</label>
			</div>
		</div>
		<!-- eof enabled !-->
		
		<!-- alert or notification !-->
		<div class="col-lg-1 alert-checkbox">
			<div class="checkbox">
				<label>
					<input type="radio" name="optNotificationType" id="optNotificationType" value="Alert" checked="checked"><strong> Alert</strong><br>
					<input type="radio" name="optNotificationType" id="optNotificationType" value="Notification" ><strong> Notification</strong>

				</label>
			</div>
		</div>
		<!-- eof alert or notification !-->

		<!-- public or private !-->
		<div class="col-lg-1 alert-checkbox">
			<div class="checkbox">
				<label>
					<input type="radio" name="optPublicOrPrivate" id="optPublicOrPrivate" value="Private" checked><strong> Private</strong><br>
					<input type="radio" name="optPublicOrPrivate" id="optPublicOrPrivate" value="Public" ><strong> Public</strong>
				</label>
			</div>
		</div>
		<!-- eof public or private !-->
		
		<!-- when label !-->
		<div class="col-lg-1">
			<p class="when-line" align="right">
				<strong>When</strong>
			</p>
		</div>
		<!-- eof when label !-->
	
		<!-- when select !-->
		<div class="col-lg-4">
			<select class="form-control when-line" name="selCond" id="selCond">
				<!--<option value=""> -- none selected -- </option>-->
				<option value="ContainsServiceNotes" selected="selected">Ticket contains services notes from technician</option>
			</select>
		</div>
		<!-- eof when select !-->		
		
	</div>


	<!-- send email alert !-->
	<div class="row row-line">
	
		<!-- send email !-->
		<div class="col-lg-1">
			<p class="email-alert-line" align="right" ><strong>Send an email alert to</strong></p>
		</div>
		<!-- eof send email !-->

		<!-- none from here !-->
		<div class="col-lg-2">
			<select class="form-control email-alert-line email-multi-select" id="selEmailto" name="selEmailto" multiple>
				<option value="0" selected>--- none from here ---</option>
				<option value="<%=Session("UserNo")%>"><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option>
		      	<%'Users dropdown
		      	 
	      	  	SQL = "SELECT UserNo, userFirstName, userLastName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
	      	  	SQL = SQL & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo")
	      	  	SQL = SQL & " order by  userFirstName, userLastName"

				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
			
				If not rs.EOF Then
					Do
						FullName = rs("userFirstName") & " " & rs("userLastName")
						Response.Write("<option value='" & rs("UserNo") & "'>" & FullName & "</option>")
						rs.movenext
					Loop until rs.eof
				End If
				set rs = Nothing
				cnn8.close
				set cnn8 = Nothing
		      	%>
			</select>
			<strong>Use CTRL and SHIFT to make multiple selections</strong>
		</div>
		<!-- eof none from here !-->
	
		<!-- and !-->
		<div class="col-lg-1">
			<p align="center" class="email-alert-line"><strong>and</strong></p>
		</div>
		<!-- eof and !-->
	
		<!-- separate emails !-->
		<div class="col-lg-3">
			<textarea class="form-control email-alert-line" rows="4" id="txtaAdditionalEmails" name="txtaAdditionalEmails"></textarea>
			<strong>Separate multiple email addresses with a semicolon</strong>
		</div>
		<!-- eof separate emails !-->
	
		<!-- verbiage !-->
		<div class="col-lg-3">
			<strong>Verbiage to include in alert email</strong>
			<textarea class="form-control" rows="4" id="txtaVerbiageEmail" name="txtaVerbiageEmail"></textarea>
	 	</div>
		<!-- eof verbiage !-->
	
	</div>
	<!-- eof  send email alert !-->


	<!-- send text alert !-->
	<div class="row row-line">
		
		<!-- send email !-->
		<div class="col-lg-1">
			<p class="email-alert-line" align="right" ><strong>Send a text alert to</strong></p>
		</div>
		<!-- eof send email !-->
		
		<!-- none from here !-->
		<div class="col-lg-2">
			<select class="form-control email-alert-line email-multi-select"  id="selTextto" name="selTextto" multiple>
				<option value="0" selected>--- none from here ---</option>
				<option value="<%=Session("UserNo")%>"><%=GetUserFirstAndLastNameByUserNo(Session("UserNo"))%></option>
		      	<%'Users dropdown
		      	 
	      	  	SQL = "SELECT UserNo, userFirstName, userLastName FROM " & MUV_Read("SQL_Owner") & ".tblUsers "
	      	  	SQL = SQL & "WHERE userArchived <> 1 and UserNo <> " & Session("UserNo")
	      	  	SQL = SQL & " order by  userFirstName, userLastName"

				Set cnn8 = Server.CreateObject("ADODB.Connection")
				cnn8.open (Session("ClientCnnString"))
				Set rs = Server.CreateObject("ADODB.Recordset")
				rs.CursorLocation = 3 
				Set rs = cnn8.Execute(SQL)
			
				If not rs.EOF Then
					Do
						FullName = rs("userFirstName") & " " & rs("userLastName")
						Response.Write("<option value='" & rs("UserNo") & "'>" & FullName & "</option>")
						rs.movenext
					Loop until rs.eof
				End If
				set rs = Nothing
				cnn8.close
				set cnn8 = Nothing
		      	%>
			</select>
			<strong>Use CTRL and SHIFT to make multiple selections</strong>
		</div>
		<!-- eof none from here !-->
	
	
	<!-- and !-->
	<div class="col-lg-1">
		<p align="center" class="email-alert-line"><strong>and</strong></p>
	</div>
	<!-- eof and !-->
	
	<!-- separate emails !-->
	<div class="col-lg-3">
		<textarea class="form-control email-alert-line" rows="4" id="txtaAdditionalTexts" name="txtaAdditionalTexts"></textarea>
		<strong>Separate multiple phone numbers with a semicolon</strong>
	</div>
	<!-- eof separate emails !-->
	
	<!-- verbiage !--> 
	<div class="col-lg-3">
		<strong>Message to include in text alert</strong><br>
		<input type="text"  value=""  class="form-control" id="txtAlertTextVerbiage" name="txtAlertTextVerbiage">
 	</div>
	<!-- eof verbiage !-->
	
</div>
<!-- eof  send text alert !-->


	<!-- reference line !-->
	<div class="row row-line">
		<div class="col-lg-10">
 			
			<div class="row row-line">
				<div class="col-lg-12 alertbutton">
					<p align="right"><a href="<%= BaseURL %>system/alerts/main.asp#ServiceOtherConditions">
    					<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Alert List</button>
					</a>
					<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button></p>
		    	</div>
			</div>
		</div>
	</div>
</form>


<!--#include file="../../inc/footer-main.asp"-->

