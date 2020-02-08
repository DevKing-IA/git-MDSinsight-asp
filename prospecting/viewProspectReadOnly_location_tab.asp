<%'*************************
' **** Location Info Tab****
'***************************
%>
<div role="tabpanel" class="tab-pane fade" id="location">

<!-- first column !-->
<div class="col-lg-4">

	<table class="table standard-font">
		<tbody>		
		  	<tr>
			  	<th class="label-col"><strong>Location Disambiguation</strong></th>
			  	<th class="input-col"><textarea class="form-control" rows="6" name="txtDisambiguation" id="txtDisambiguation"><%=LocationDisambiguation%></textarea></th>
		  	</tr>
		</tbody>
	</table>

</div>
<!-- eof first column !-->	
	

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
' **** eof Location Tab****
'*****************************
%>