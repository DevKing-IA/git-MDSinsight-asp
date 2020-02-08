	<style type="text/css">
		.the-select{
			min-height: 150px;
			max-width: 50%;
		}
	</style>

		
	<script>
	
	function SelectFieldTechFunc(usrnumber)
	  {   
	
		  var  userno=usrnumber;
				
		   if(userno!='')
		   {
		    $.ajax({
		   type:'post',
		      url:'setDispatchModalOptions.asp',
		          data:{userno: userno},
					success: function(msg){
						window.location = "dispatchCenter_modal.asp";
					}
		 });
		  }
	}
	</script>

	
	<div class="modal-dialog modal-height" id="dispatchCenterModal-<%=rs_Tickets("MemoNumber")%>">
    <div class="modal-content">
      <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
        <h4 class="modal-title" id="myModalLabel" align="center">Dispatch Center</h4>
      </div>
      <div class="modal-header">
        <h3 class="modal-title" id="myModalLabel" align="center"><%=GetUserDisplayNameByUserNo(rs_Users("userNo"))%>&nbsp;-&nbsp;Ticket #: <%=rs_Tickets("MemoNumber")%>&nbsp;-&nbsp;<%=GetTerm("Account")%> #:<%=GetServiceTicketCust(rs_Tickets("MemoNumber"))%></h3>
      </div>

	<form method="post" action="dispatchCenter_modal_SaveValues.asp" name="frmDispatchModal" id="frmDispatchModal">
	
		 <input type='hidden' id='txtServiceTicketNumber' name='txtServiceTicketNumber' value='<%=rs_Tickets("MemoNumber")%>'>
		 <input type='hidden' id='txtAccountNumber' name='txtAccountNumber' value='<%=GetServiceTicketCust(rs_Tickets("MemoNumber"))%>'>		 
	      <!-- insert content in here !-->
	      
	      <div class="modal-body ativa-scroll">
	
	 			    	<!-- field techs !-->
	 			    	<div class="col-lg-12">
	    					<p align="left"><label >Select <%=GetTerm("Field Service Tech")%> to reassign this ticket to</label></p>
							<select name="selFieldTech" id="selFieldTech"  multiple="multiple" class="form-control the-select"> <!-- onchange="SelectFieldTechFunc(this.value)"> !-->
								<%	
								'Fixit
								' cheap fix to let adam henchel see service stuff wihtout being a service manager

								SQLmodal = "SELECT * FROM " & MUV_Read("SQL_Owner") & ".tblUsers WHERE UserNo <> " & rs_Users("userNo") & " AND (userType = 'Field Service' OR userType = 'Service Manager' OR UserNo=56) and userArchived <> 1 Order By UserType,userDisplayName"
								
								Set cnnmodal = Server.CreateObject("ADODB.Connection")
								cnnmodal.open (Session("ClientCnnString"))
								Set rsmodal = Server.CreateObject("ADODB.Recordset")
								rsmodal.CursorLocation = 3 
								Set rsmodal = cnnmodal.Execute(SQLmodal)
		
								If not rsmodal.EOF Then
	
									Do While Not rsmodal.EOF
										userFirstName = rsmodal("userFirstName")
										userLastName = rsmodal("userLastName")
										userDisplayName = rsmodal("userDisplayName")
										userEmail = rsmodal("userEmail")
										userNo = rsmodal("UserNo")
										
										%><option value='<%=userNo%>'><%=userFirstName%>&nbsp;<%=userLastName%></option><%
										
										rsmodal.MoveNext
									Loop
	
								End If
								
								Set rsmodal = Nothing
								cnnmodal.Close
								Set cnnmodal = Nothing
								%>
								<option value='0'>UN-DISPATCH</option>
							</select>
	  					 </div>
	 			    	<!-- eof field techs !-->
	
	<!-- checkboxes !-->
	<div class="col-lg-12">
		<div class="checkbox">
			<label>
   				<input type="checkbox" name="chkSendEmail" id="chkSendEmail"<%If Instr(FSDefaultNotificationMethod(),"Email") <> 0 Then Response.Write(" checked")%>>
   				<strong>Send email</strong>
			</label>
		</div>
		
		<div class="checkbox">
			<label>
			    <input type="checkbox" name="chkSendText" id="chkSendText"<%If Instr(FSDefaultNotificationMethod(),"Text") <> 0 Then Response.Write(" checked")%>>
			    <strong>Send text message</strong>
			</label>
		</div>
	</div>
	<!-- eof checkboxes !-->
	
	 	      	
	  
	      </div>
	      <!-- eof content insertion !-->
	      
	      
	      <div class="modal-footer">
	         <button type="button" class="btn btn-default" data-dismiss="modal">Cancel</button>
   	         <button type="submit"  class="btn btn-primary" >REASSIGN</button>
	      </div>

	</form>
  </div>
</div>