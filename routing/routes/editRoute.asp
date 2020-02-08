<!--#include file="../../inc/header.asp"-->

<%

InternalRecordIdentifier = Request.QueryString("i") 
If InternalRecordIdentifier = "" Then Response.Redirect("main.asp")


'*******************************************************************
'Lookup the record as it exists now so we can fill in the form fields

SQL = "SELECT * FROM RT_Routes where InternalRecordIdentifier = " & InternalRecordIdentifier
	
Set cnn8 = Server.CreateObject("ADODB.Connection")
cnn8.open (Session("ClientCnnString"))
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn8.Execute(SQL)
	
If not rs.EOF Then	
	RouteID = rs("RouteID")
	RouteDescription = rs("RouteDescription")
	DefaultDriverUserNo = rs("DefaultDriverUserNo")
	ThirdPartyCarrier = rs("ThirdPartyCarrier")
	ShowOnDBoard = rs("ShowOnDBoard")
	ShowInWebApp = rs("ShowInWebApp")
	ShowInPlanner = rs("ShowInPlanner")
	RouteMonday = rs("Monday")
	RouteTuesday = rs("Tuesday")
	RouteWednesday = rs("Wednesday")
	RouteThursday = rs("Thursday")
	RouteFriday = rs("Friday")
	RouteSaturday = rs("Saturday")
	RouteSunday = rs("Sunday")
End If

set rs = Nothing
cnn8.close
set cnn8 = Nothing
'***********************************************************************
'End Lookup the record as it exists now
'***********************************************************************
%>

<SCRIPT LANGUAGE="JavaScript">
	

   function validateRouteForm()
    {
    
       if (document.frmEditRoute.txtRouteID.value == "") {
            swal("Route ID cannot be blank.");
            return false;
       }
       if (document.frmEditRoute.txtRouteDescription.value == "") {
            swal("Route description cannot be blank.");
            return false;
       }
        if (document.frmEditRoute.selDefaultDriverUserNo.value == "") {
            swal("You must select a default driver for this route.");
            return false;
       }
       
       return true;
    }
</script>
       
 

<style type="text/css">
	.input-group {
		margin-top:10px;
	}

	.select-line{
		margin-bottom: 15px;
	}
	
	.enable-disable{
		margin-top:20px;
	}
	
	.row-line{
		margin-bottom: 25px;
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
	
	.control-label{
		font-size:12px;
		font-weight:normal;
		pEditing-top:10px;
	}
	.control-label-last{
		pEditing-top:0px;
	}
	
	.required{
		border-left:3px solid red;
	}
	
	.container {
		margin-bottom: 20px;
		margin-top: 20px;
		margin-left:0px;
		width: 100%;
	}

	.container .row {
		margin-bottom: 20px;
		margin-top: 20px;
	}
	
	.tab-colors-box{
		padding:15px;
		border:2px solid #000;
		margin:0px 0px 15px 0px;
		width:100%;
		display:block;
		float:left;
	}
	
	.tab-colors-title strong{
		width:100%;
		text-align:center;
		display:block;
	}
	
	.tab-colors-title .row{
		margin-bottom:0px;
	}
	
	.line-full{
	 	margin-bottom:20px;
	}
	
	.multi-select{
		min-height:200px;
		min-width:180px;
	}
	
	.custom-select{
		width: auto !important;
		display:inline-block;
	}
	
	.select-large{
		min-width:40% !important;
	}
	
</style>

<h1 class="page-header"> Edit&nbsp;<%= GetTerm("Route") %></h1>

<div class="container">

	<form method="POST" action="EditRoute_Submit.asp" name="frmEditRoute" id="frmEditRoute" onsubmit="return validateRouteForm();">


	<input type="hidden" name="txtInternalRecordIdentifier" id="txtInternalRecordIdentifier" value="<%= InternalRecordIdentifier %>">
	
	 <!-- weekly snapshot report -->
	 <div class="col-lg-4">
	
	    <div class="col-lg-12 tab-colors-title">
			<div class="row">
				<div class="col-lg-12" align="center">
					 <strong>&nbsp;</strong>
				</div>
			</div>
		</div>
    
	
		<div class="col-lg-12">
			<div class="tab-colors-box">
		                 
		        
		         <!-- line -->
		         <div class="row">
		             <div class="col-lg-12 line-full">
						<p><%= GetTerm("Route") %> ID:</p>
						<input type="text"class="form-control" style="width:100%;" name="txtRouteID" id="txtRouteID" value="<%= RouteID %>">
		             </div>
		         </div>
		         <!-- eof line -->
         
		               
		         <!-- line -->
		         <div class="row">
		             <div class="col-lg-12 line-full">
						<p><%= GetTerm("Route") %> Description:</p>
						<input type="text"class="form-control" style="width:100%;" name="txtRouteDescription" id="txtRouteDescription" value="<%= RouteDescription %>">
		             </div>
		         </div>
		         <!-- eof line -->
		         
		         <!-- line -->
		         <div class="row">
		             <div class="col-lg-12 line-full">
						<p><%= GetTerm("Route") %> Default Driver:</p>
				      	<select class="form-control" style="width:100%;" name="selDefaultDriverUserNo" id="selDefaultDriverUserNo">
				      	  	<option value="">-- none --</option>
					      	<% 'Get all users
					      	  	SQL9 = "SELECT * FROM tblUsers WHERE userArchived <> 1 AND userEnabled = 1 ORDER BY userLastName ASC"
								Set cnn9 = Server.CreateObject("ADODB.Connection")
								cnn9.open (Session("ClientCnnString"))
								Set rs9 = Server.CreateObject("ADODB.Recordset")
								rs9.CursorLocation = 3 
								Set rs9 = cnn9.Execute(SQL9)
								If not rs9.EOF Then
									Do
										If cInt(DefaultDriverUserNo) = cInt(rs9("userNo")) Then
											Response.Write("<option value='" & rs9("userNo") & "' selected='selected'>" & rs9("userFirstName") & " " & rs9("userLastName") & "</option>")
										Else
											Response.Write("<option value='" & rs9("userNo") & "'>" & rs9("userFirstName") & " " & rs9("userLastName") & "</option>")
										End If
										rs9.movenext
									Loop until rs9.eof
								End If
								set rs9 = Nothing
								cnn9.close
								set cnn9 = Nothing
					      	%>
						</select>
		             </div>
		         </div>
		         <!-- eof line -->
		         
			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12">
		               	Third Party Carrier <input type="checkbox" name="chkThirdPartyCarrier" id="chkThirdPartyCarrier" <% If ThirdPartyCarrier = 1 Then Response.Write("checked='checked'") %>><br>
		            </div>
		            <!-- eof line -->
		        </div>        		         
		         
         
         
	      	</div>  <!-- eof col-lg-tab-colors-box -->       
	      </div> <!-- eof col-lg-12 -->     

               
   	</div><!-- eof col-lg-4 -->          
 	
 
	 <!-- weekly snapshot report -->
	 <div class="col-lg-4">
	
	    <div class="col-lg-12 tab-colors-title">
			<div class="row">
				<div class="col-lg-12" align="center">
					 <strong><%= GetTerm("Route") %> Visibility</strong>
				</div>
			</div>
		</div>
    
	
		<div class="col-lg-12">
			<div class="tab-colors-box">
		
			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12 line-full">
		               	Show <%= GetTerm("Route") %> on Delivery Board
		               	<input type="checkbox" name="chkShowOnDBoard" id="chkShowOnDBoard" <% If ShowOnDBoard = 1 OR ShowOnDBoard = "" Then Response.Write("checked='checked'") %>>
		            </div>
		            <!-- eof line -->
		         </div>                          
		        
		               
			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12 line-full">
		            	Show <%= GetTerm("Route") %> on Delivery Board Planner
		               	<input type="checkbox" name="chkShowInPlanner" id="chkShowInPlanner" <% If ShowInPlanner = 1 OR ShowInPlanner = "" Then Response.Write("checked='checked'") %>>
		            </div>
		            <!-- eof line -->
		         </div>  
		         

			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12 line-full">
		               	Show <%= GetTerm("Route") %> in Web App 
		               	<input type="checkbox" name="chkShowInWebApp" id="chkShowInWebApp" <% If ShowInWebApp = 1 OR ShowInWebApp = "" Then Response.Write("checked='checked'") %>>
		            </div>
		            <!-- eof line -->
		         </div>  
		                  
         
	      	</div>  <!-- eof col-lg-tab-colors-box -->       
	      </div> <!-- eof col-lg-12 -->     

               
   	</div><!-- eof col-lg-4 -->    
   	
   	
   	
	 <!-- weekly snapshot report -->
	 <div class="col-lg-2">
	
	    <div class="col-lg-12 tab-colors-title">
			<div class="row">
				<div class="col-lg-12" align="center">
					 <strong><%= GetTerm("Route") %> Days</strong>
				</div>
			</div>
		</div>
    
	
		<div class="col-lg-12">
			<div class="tab-colors-box">
		
			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12">
		               	Monday <input type="checkbox" name="chkRouteMonday" id="chkRouteMonday" <% If RouteMonday = 1 OR RouteMonday = "" Then Response.Write("checked='checked'") %>><br>
		            </div>
		            <!-- eof line -->
		        </div>                          
		        
  			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12">
		               	Tuesday <input type="checkbox" name="chkRouteTuesday" id="chkRouteTuesday" <% If RouteTuesday = 1 OR RouteTuesday = "" Then Response.Write("checked='checked'") %>>
		            </div>
		            <!-- eof line -->
		        </div>            
 
		        
  			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12">
		               	Wednesday <input type="checkbox" name="chkRouteWednesday" id="chkRouteWednesday" <% If RouteWednesday = 1 OR RouteWednesday = "" Then Response.Write("checked='checked'") %>>
		            </div>
		            <!-- eof line -->
		        </div>   
		        
		        
  			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12">
		               	Thursday <input type="checkbox" name="chkRouteThursday" id="chkRouteThursday" <% If RouteThursday = 1 OR RouteThursday = "" Then Response.Write("checked='checked'") %>>
		            </div>
		            <!-- eof line -->
		        </div>   
		        

		        
  			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12">
		               	Friday <input type="checkbox" name="chkRouteFriday" id="chkRouteFriday" <% If RouteFriday = 1 OR RouteFriday = "" Then Response.Write("checked='checked'") %>>
		            </div>
		            <!-- eof line -->
		        </div>   
		        
		        
  			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12">
		               	Saturday <input type="checkbox" name="chkRouteSaturday" id="chkRouteSaturday" <% If RouteSaturday = 1 Then Response.Write("checked='checked'") %>>
		            </div>
		            <!-- eof line -->
		        </div>   
		        
		        
  			    <div class="row">
		            <!-- line -->
		            <div class="col-lg-12">
		               	Sunday <input type="checkbox" name="chkRouteSunday" id="chkRouteSunday" <% If RouteSunday = 1 Then Response.Write("checked='checked'") %>>
		            </div>
		            <!-- eof line -->
		        </div>   
		        
		        
         
	      	</div>  <!-- eof col-lg-tab-colors-box -->       
	      </div> <!-- eof col-lg-12 -->     

               
   	</div><!-- eof col-lg-4 -->    				
		
    <!-- cancel / submit !-->
	<div class="row row-line pull-right">
		<div class="col-lg-12 alertbutton">
			<div class="col-lg-12">
				<a href="<%= BaseURL %>routing/routes/main.asp">
    				<button type="button" class="btn btn-default">&lsaquo; Cancel &amp; Go Back To Routes List</button>
				</a>
				<button type="submit" class="btn btn-primary"><i class="far fa-save"></i> Save</button>
			</div>
	    </div>
	</div>
		
	</form>
</div>

<!--#include file="../../inc/footer-main.asp"-->
