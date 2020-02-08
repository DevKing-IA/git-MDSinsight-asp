<!--#include file="../inc/header-field-service-mobile.asp"-->
<!--#include file="../inc/InsightFuncs_Service.asp"-->

<style type="text/css">
	.fieldservice-heading{
		line-height: 1;
		padding-top: 5px;
		padding-bottom: 5px;
	}
	
	.btn-home{
		color: #fff;
		margin-top: -2px;
		margin-left: 5px;
		float: left;
 	}
	
 ul{
	 color: #666;
	 font-size: 13px;
	 text-transform: uppercase;
	 list-style-type: none;
	     -webkit-margin-before: 0px;
    -webkit-margin-after: 0px;
    -webkit-margin-start: 0px;
    -webkit-margin-end: 0px;
    -webkit-padding-start: 0px;
 }
 
 .enroute{
	 color: green;
 }
 
 .btn-spacing{
	 margin-bottom: 40px;
 }
 
 
 
.btn-block {
    width: auto;
    display: inline-block;
    margin-right:2em;
}
 
 .row{
 	/* flex-wrap: nowrap !important; */
 }

 
 @media (max-width: 767px) {
 	.mob-col{
 		/* width: auto !important;  */
 	}
 }
 
 .driver-menu{
 	margin:5px; 
 	text-align:center
 }
 
 .left-segmented-control{
	-webkit-border-top-left-radius: 4px !important;
	-webkit-border-bottom-left-radius: 4px !important;
	-moz-border-radius-topleft: 4px !important;
	-moz-border-radius-bottomleft: 4px !important;
	border-top-left-radius: 4px !important;
	border-bottom-left-radius: 4px !important;
	}

 .right-segmented-control{
	-webkit-border-top-right-radius: 4px !important;
	-webkit-border-bottom-right-radius: 4px !important;
	-moz-border-radius-topright: 4px !important;
	-moz-border-radius-bottomright: 4px !important;
	border-top-right-radius: 4px !important;
	border-bottom-right-radius: 4px !important;
	}
	
</style>       


 
<h1 class="fieldservice-heading" ><a class="btn-home" href="main_menu.asp" role="button"><i class="fa fa-bars"></i></a> Your Stops</h1>

<div class="container-fluid">


<div class="driver-menu">
	<div class="row">
	  <div class="col-lg-12">
	  	<div class="form-group">
		     <div class="input-group">
		       <div class="input-group-btn">
		         	<button type="button" class="btn btn-default left-segmented-control" tabindex="-1" id="menuUNACK">UNACK (4)</button>
		         	<button type="button" class="btn btn-primary" tabindex="-1" id="menuOPEN">OPEN (5)</button>
		       		<button type="button" class="btn btn-default right-segmented-control" tabindex="-1" id="menuCLOSEDREDO">CLOSED/REDO (0)</button>
		        </div>
		      </div>
		    </div>
	    </div>
	  </div>
	</div><!-- Row -->
</div>

<%
'Now lookup the other info
SQL = "Select CASE MemoStage WHEN 'Dispatched' THEN 0 ELSE 1 END As ACK,* from "
SQL = SQL & "FS_ServiceMemosDetail where serviceDetailRecNumber in "
SQL = SQL & "(Select max(serviceDetailRecNumber) from FS_ServiceMemosDetail where UserNoOfServiceTech = " & Session("UserNo") & " group by memonumber ) "
SQL = SQL & "  AND ClosedorCancelled <> 1 "
SQL = SQL & "Order by ACK, OriginalDispatchDateTime"
 
'Response.Write(SQL)

Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rsCust = Server.CreateObject("ADODB.Recordset")
rsCust.CursorLocation = 3 


Set rs = cnn8.Execute(SQL)
			
	If not rs.EOF Then
		Do While Not rs.EOF
		
			If LastTechUserNo(rs("MemoNumber")) = Session("UserNo") Then ' If we are not the latest tech, it was reassigned & isnt ours anymore
			
				If AwaitingRedispatch(rs("MemoNumber")) <> True Then %>
					
					<!-- client box !-->
					<% If TicketIsUrgent(rs("MemoNumber")) Then 
						Response.Write("<div class='row alert alert-danger'>")
					Else
						Response.Write("<div class='row alert alert-warning'>")
					End If%>	
			
						<!-- client info !-->
						<div class="col-lg-12 mob-col">
							<strong>
							<% If TicketIsUrgent(rs("MemoNumber")) Then Response.Write("<font color='red'>*URGENT*</font><br>")%>
							<%=GetTerm("Account")%>:&nbsp;<%=rs("CustNum")%><br>
							Ticket#: <%= rs("MemoNumber")%>
							</strong><br>
							
							<%'Lookup cust info
							SQL = "Select Name,Addr1,Addr2,CityStateZip,Phone,Contact from AR_Customer where CustNum = '" & rs("CustNum") & "'"
							Set rsCust = cnn8.Execute(SQL)
							If not rsCust.Eof Then
							%>
							<ul><%
								Response.Write("<li>" & rsCust("Name") & "</li>")
								Response.Write("<li>" & rsCust("Addr1") & "</li>")
								Response.Write("<li>" & rsCust("Addr2") & "</li>")
								Response.Write("<li>" & rsCust("CityStateZip") & "</li>")
								Response.Write("<li>" & rsCust("Phone") & "</li>")
								Response.Write("<li>" & rsCust("Contact") & "</li>")
							%></ul>
							<%End If
							TktType = "*Service*"
							If filterChangeModuleOn() = True Then 
								If TicketIsFilterChange(rs("MemoNumber")) Then TktType =  "*Filter Change*" 
							End If
							Response.Write(TktType)%> 
						</div>
					
						<!-- buttons !-->
						<div class="col-lg-12 mob-col">
							<p>&nbsp;</p>
							<form method="post" action="viewTicket.asp" name="frmViewTicket" id="frmViewTicket">
								<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=rs("MemoNumber")%>'>		 
								<button type="submit" class="btn btn-primary btn-block btn-spacing col-xs-6 pull-left" >Details</button>
							</form>

							<%Select Case GetServiceTicketCurrentStage(rs("MemoNumber"))
								Case "Dispatched"
									If FS_TechCanDecline() Then %>
										<a href="tap_decline.asp?t=<%=rs("MemoNumber")%>&c=<%=rs("CustNum")%>&u=<%=Session("Userno")%>" class="btn btn-danger btn-block btn-spacing pull-right" role="button">Tap To Decline</a>
									<% End If %> 
									<a href="tap_ack.asp?t=<%=rs("MemoNumber")%>&c=<%=rs("CustNum")%>&u=<%=Session("Userno")%>" class="btn btn-warning btn-block btn-spacing pull-right" role="button">Tap To Acknowledge</a>
								<%Case "Dispatch Acknowledged"%>
									<a href="tap_enroute.asp?t=<%=rs("MemoNumber")%>&c=<%=rs("CustNum")%>&u=<%=Session("Userno")%>" class="btn btn-primary btn-block btn-spacing col-xs-6 pull-right" role="button">Tap When<br>En Route</a>
								<%Case "En Route"%>
									<a href="tap_onsite.asp?t=<%=rs("MemoNumber")%>&c=<%=rs("CustNum")%>&u=<%=Session("Userno")%>" class="btn btn-primary  btn-block btn-spacing col-xs-6 pull-right" role="button">Tap When<br>On Site</a>
								<%Case Else %>
									<form method="post" action="onSite.asp" name="frmOnSite" id="frmOnSite">
										<input type='hidden' id='txtTicketNumber' name='txtTicketNumber' value='<%=rs("MemoNumber")%>'>		 
										<button type="submit" class="btn btn-primary  btn-block btn-spacing col-xs-6 pull-right" >Tap To Complete<br>This Stop</button>
									</form>
							<%End Select%>
							
							
						</div>
						<!-- eof buttons !-->
					</div>
					<!-- eof client box !-->	
					
					<hr />
		
					<%
					End If
				End If
			rs.movenext
		loop
	Else
		%>No Service calls for you!<%
	End IF

cnn8.close
Set rsCust = Nothing
Set rs = Nothing
Set cnn8 = Nothing				
%></div><!--#include file="../inc/footer-field-service-noTimeout.asp"-->