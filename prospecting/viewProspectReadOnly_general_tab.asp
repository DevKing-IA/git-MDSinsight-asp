<%'************************
' **** General Info Tab****
'**************************
%>
<div role="tabpanel" class="tab-pane fade" id="general">

<!-- first column !-->
<div class="col-lg-4">

	<table class="table standard-font">
		<tbody>
		  	<tr>
			  	<th class="label-col" ><strong>Company</strong></th>
			  	<th class="input-col" ><input type="text" class="form-control red-line" name="txtCompany" id="txtCompany" value="<%=Company%>"></th>
		  	</tr>
		
		  	<tr>
			  	<th class="label-col"><strong>Street</strong></th>
			  	<th class="input-col"><input type="text" class="form-control red-line" name="txtStreet" id="txtStreet" value="<%=Street%>"></th>
		  	</tr>
		
		  	<tr>
			  	<th class="label-col"><strong>Floor/ Suite/ Room</strong></th>
			  	<th class="input-col"><input type="text" class="form-control" name="txtSuite" id="txtSuite" value="<%=Suite%>"></th>
		
		  	</tr>
		
		  	<tr>
			  	<th class="label-col"><strong>Phone</strong></th>
			  	<th class="input-col"><input type="text" class="form-control" name="txtPhone" id="txtPhone" value="<%=Phone%>"></th>
		  	</tr>
		
		  	<tr>
			  	<th class="label-col"><strong>Extension</strong></th>
			  	<th class="input-col"><input type="text" class="form-control" name="txtPhoneExt" id="txtPhoneExt" value="<%=PhoneExt %>"></th>
		  	</tr>
		  	
		  	<tr>
			  	<th class="label-col"><strong>Fax</strong></th>
			  	<th class="input-col"><input type="text" class="form-control" name="txtFax" id="txtFax" value="<%=Fax%>"></th>
		  	</tr>
		
		  	<tr>
			  	<th class="label-col"><strong>City</strong></th>
			  	<th class="input-col"><input type="text" class="form-control red-line" name="txtCity" id="txtCity" value="<%=City%>"></th>
		  	</tr>
		  	
		  	<tr>
			  	<th class="label-col"><strong>State / Province</strong></th>
			  	<th class="input-col">
			  	<select class="form-control">
			  		<!--#include file="stateList.asp"-->
				</select></th>
		  	</tr>
		
		  	<tr>
			  	<th class="label-col"><strong>Zip / Postal Code</strong></th>
			  	<th class="input-col"><input type="text" class="form-control" name="txtPostalCode" id="txtPostalCode" value="<%=PostalCode%>"></th>
		  	</tr>
		  	
		  	<tr>
			  	<th class="label-col"><strong>Country</strong></th>
				  	<th class="input-col">
				  		<select class="form-control">
							<!--#include file="countryList.asp"-->
						</select>
					</th>
		  	</tr>
		
		  	<tr>
			  	<th class="label-col"><strong>Website</strong></th>
			  	<th class="input-col"><input type="text" class="form-control" name="txtWebsite" id="txtWebsite" value="<%=Website%>"> <a href="#" target="_blank"><i class="fa fa-globe fa-lg" aria-hidden="true"></i></a></th>
		  	</tr> 
		
		  	<tr>
			  	<th class="label-col"><strong>Location Disambiguation</strong></th>
			  	<th class="input-col"><textarea class="form-control" rows="6" name="txtDisambiguation" id="txtDisambiguation"><%=LocationDisambiguation%></textarea></th>
		  	</tr>
		</tbody>
	</table>

</div>
<!-- eof first column !-->	
	
	
<!-- second column !-->
<div class="col-lg-4">
	
<table class="table standard-font">
<tbody>
  	
  	<tr>
	  	<th class="label-col"><strong>Description</strong></th>
	  	<th class="input-col"><textarea class="form-control" rows="6" name="txtDescription " id="txtDescription "><%=Description %></textarea></th>
  	</tr>

  	<tr>
	  	<th class="label-col"><strong># Pantries / Stations</strong></th>
	  	<th class="input-col">
	  		<select class="form-control">
				<option value="-1">-- Not Specified --</option>
				<%
				For x = 1 to 50
				 If x=NumberOfPantries Then
					If x = 0 Then
					 	Response.Write("<option selected value='-1'>-- Not Specified --</option>")
					 Else
					 	Response.Write("<option selected>" & x & "</option>")
					 End If
				 Else
				 	Response.Write("<option>" & x & "</option>")
				 End If
				Next
				%>
			</select>
		</th>
  	</tr>

  	<tr>
	  	<th class="label-col"><strong># Employees</strong></th>
	  	<th class="input-col">
	  		<select class="form-control">
  	  			<option value="0">-- Not Specified --</option>
		      	<% 'Get all emplloyee ranges
					SQL9 = "SELECT *, Cast(LEFT(Range,CHARINDEX('-',Range)-1) as int) as Expr1 FROM PR_EmployeeRangeTable "
					SQL9 = SQL9 & "order by Expr1"

					Set cnn9 = Server.CreateObject("ADODB.Connection")
					cnn9.open (Session("ClientCnnString"))
					Set rs9 = Server.CreateObject("ADODB.Recordset")
					rs9.CursorLocation = 3 
					Set rs9 = cnn9.Execute(SQL9)
						
					If not rs9.EOF Then
						Do
							Response.Write("<option ")
							If EmployeeRangeNumber = rs9("InternalRecordIdentifier") Then
								Response.Write("selected ")
								If StageNumber =  0 Then
									Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>-- Not Specified --</option>")
								Else
									Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Range")& "</option>")
								End If
							Else
								Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Range")& "</option>")
							End If
							rs9.movenext
						Loop until rs9.eof
					End If
					set rs9 = Nothing
					cnn9.close
					set cnn9 = Nothing
				%>
			</select>
			<%	If userCanEditCRMOnTheFly(Session("UserNo")) Then %>
					<a class="plus-button" data-toggle="modal" data-target="#ONTHEFLYmodalEmployeeRange">
						<span data-toggle="tooltip" data-placement="right" title="Add new employee range"><i class="fa fa-plus text-primary" aria-hidden="true" ></i></span>
            		</a>
            <% End If %>
		</th>
  	</tr>

  	<tr>
	  	<th class="label-col"><strong>Industry</strong></th>
	  	<th class="input-col">
	  		<select class="form-control" id="selIndustry" name="selIndustry" >
  	  			<option value="0">-- Not Specified --</option>
  	  			<%
  	  			'Get all industries
		      	  	SQL9 = "SELECT * FROM PR_Industries order by Industry "

					Set cnn9 = Server.CreateObject("ADODB.Connection")
					cnn9.open (Session("ClientCnnString"))
					Set rs9 = Server.CreateObject("ADODB.Recordset")
					rs9.CursorLocation = 3 
					Set rs9 = cnn9.Execute(SQL9)
						
					If not rs9.EOF Then
						Do
							Response.Write("<option ")
							If IndustryNumber = rs9("InternalRecordIdentifier") Then
								Response.Write("selected ")
								If IndustryNumber = 0 Then
									Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>-- Not Specified --</option>")
								Else
									Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Industry")& "</option>")
								End If
							Else
								Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Industry")& "</option>")													
							End If
							rs9.movenext
						Loop until rs9.eof
					End If
					set rs9 = Nothing
					cnn9.close
					set cnn9 = Nothing
				%>
			</select>
			<%	If userCanEditCRMOnTheFly(Session("UserNo")) Then %>
					<a class="plus-button" data-toggle="modal" data-target="#ONTHEFLYmodalIndustry">
						<span data-toggle="tooltip" data-placement="right" title="Add new industry"><i class="fa fa-plus text-primary" aria-hidden="true" ></i></span>
            		</a>
            <% End If %>
		</th>
  	</tr>
  	
  	<tr>
	  	<th class="label-col"><strong>Lead Source</strong></th>
	  	<th class="input-col">
	  		<select class="form-control red-line">
  	  			<option value="0">-- Not Specified --</option>
		      	<% 'Get all lead sources
		      	  	SQL9 = "SELECT * FROM PR_LeadSources order by Leadsource"

					Set cnn9 = Server.CreateObject("ADODB.Connection")
					cnn9.open (Session("ClientCnnString"))
					Set rs9 = Server.CreateObject("ADODB.Recordset")
					rs9.CursorLocation = 3 
					Set rs9 = cnn9.Execute(SQL9)
						
					If not rs9.EOF Then
						Do
							Response.Write("<option ")
							If LeadsourceNumber = rs9("InternalRecordIdentifier") Then
								Response.Write("selected ")
								If LeadsourceNumber =  0 Then
									Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>-- Not Specified --</option>")
								Else
									Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("LeadSource")& "</option>")
								End If
							Else
								Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("LeadSource")& "</option>")
							End If
							rs9.movenext
						Loop until rs9.eof
					End If
					set rs9 = Nothing
					cnn9.close
					set cnn9 = Nothing
				%>
			</select>
			<%	If userCanEditCRMOnTheFly(Session("UserNo")) Then %>
					<a class="plus-button" data-toggle="modal" data-target="#ONTHEFLYmodalLeadSource">
						<span data-toggle="tooltip" data-placement="right" title="Add new lead source"><i class="fa fa-plus text-primary" aria-hidden="true" ></i></span>
            		</a>
            <% End If %>
		</th>
  	</tr>

  	<tr>
	  	<th class="label-col"><strong>Telemarketer</strong></th>
	  		<th class="input-col">
	  			<select class="form-control">
  	  			<option value="0">-- Not Specified --</option>
		      	<% 'Get all telemarketing users
		      	  	SQL9 = "SELECT * FROM tblUsers where UserType='Telemarketing' order by userDisplayName"

					Set cnn9 = Server.CreateObject("ADODB.Connection")
					cnn9.open (Session("ClientCnnString"))
					Set rs9 = Server.CreateObject("ADODB.Recordset")
					rs9.CursorLocation = 3 
					Set rs9 = cnn9.Execute(SQL9)
						
					If not rs9.EOF Then
						Do
							Response.Write("<option ")
							If TelemarketerUserNo = rs9("UserNo") Then
								Response.Write("selected ")
								If TelemarketerUserNo = 0 Then
									Response.Write("value='" & rs9("UserNo") & "'>-- Not Specified --</option>")
								Else
									Response.Write("value='" & rs9("UserNo") & "'>" & rs9("userDisplayName")& "</option>")
								End If
							Else
								Response.Write("value='" & rs9("UserNo") & "'>" & rs9("userDisplayName")& "</option>")
							End If
							rs9.movenext
						Loop until rs9.eof
					End If
					set rs9 = Nothing
					cnn9.close
					set cnn9 = Nothing
				%>

				</select>
			</th>
  	</tr>

  	<tr>
	  	<th class="label-col"><strong>Stage</strong></th>
	  	<th class="input-col">
	  	<select class="form-control red-line">
  	  			<option value="0">-- Not Specified --</option>
		      	<% 'Get all stages
		      	  	SQL9 = "SELECT * FROM PR_Stages order by Stage"

					Set cnn9 = Server.CreateObject("ADODB.Connection")
					cnn9.open (Session("ClientCnnString"))
					Set rs9 = Server.CreateObject("ADODB.Recordset")
					rs9.CursorLocation = 3 
					Set rs9 = cnn9.Execute(SQL9)
						
					If not rs9.EOF Then
						Do
							Response.Write("<option ")
							If StageNumber = rs9("InternalRecordIdentifier") Then
								Response.Write("selected ")
								If StageNumber =  0 Then
									Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>-- Not Specified --</option>")
								Else
									Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Stage")& "</option>")
								End If
							Else
								Response.Write("value='" & rs9("InternalRecordIdentifier") & "'>" & rs9("Stage")& "</option>")
							End If
							rs9.movenext
						Loop until rs9.eof
					End If
					set rs9 = Nothing
					cnn9.close
					set cnn9 = Nothing
				%>
		</select>
			<%	If userCanEditCRMOnTheFly(Session("UserNo")) Then %>
					<a class="plus-button" data-toggle="modal" data-target="#ONTHEFLYmodalStage">
						<span data-toggle="tooltip" data-placement="right" title="Add new stage"><i class="fa fa-plus text-primary" aria-hidden="true" ></i></span>
            		</a>
            <% End If %>

		</th>
  	</tr>

  	<tr>
	  	<th class="label-col"><strong>Probability (%)</strong></th>
	  	<th class="input-col"><div class="progressbarsone" progress="<%= GetPercentForStage(StageNumber)%>%"></div></th>
  	</tr>
</tbody>
</table>
	
					  <table class="table standard-font">
					  	<tbody>
														  		 
						  	
						  	<!-- line !-->
						  	<tr>
							  	<th class="label-col"><strong>Proposal Meeting Date</strong></th>
							  	<th class="input-col"><input type="text" id="txtCloseCancelDate" name="txtCloseCancelDate" value="<%=tmpReturn %>"  class="form-control proposal-meeting-date" data-beatpicker="true"     data-beatpicker-extra="customOptions"  data-beatpicker-format="['MM','DD','YYYY'],separator:'/'"></th>
						  	</tr>
						  	<!-- eof line !-->

					 	</tbody>
					  </table>
                      
</div>

						<!-- Embedded Google Map !-->
					<div class="col-lg-4">
						<%
						'Long & Lat, not using right now
						'MapVar = "<iframe src='https://www.google.com/maps/embed/v1/search?key=AIzaSyBR-NtdHSro_Gd_4ZukBT9NXXjdSJDrwJg&q="
						'MapVar = MapVar & Latitude & "," & Longitude
						'MapVar = MapVar & "' width='100%' height='300' frameborder='0' style='border:0' allowfullscreen></iframe>"
						'Response.Write("<strong>Lat:&nbsp;</strong>" & Latitude & "&nbsp;&nbsp;&nbsp;<strong>Long</strong>:&nbsp;" & Longitude)

						MapVar = "<iframe src='https://www.google.com/maps/embed/v1/search?key=AIzaSyBR-NtdHSro_Gd_4ZukBT9NXXjdSJDrwJg&q="
						MapVar = MapVar & Replace(Street," ","+") & "," & Replace(City," ","+")& "," & Replace(State," ","+")
						MapVar = MapVar & "' width='100%' height='300' frameborder='0' style='border:0' allowfullscreen></iframe>"
						
						Response.Write(MapVar)%>	

											 			

					</div>
					<!-- eof Embedded Google Map !-->

 	
	
</div>
<%'***************************
' **** eof general Tab****
'*****************************
%>