
<%'***********************
' **** Social Media Tab****
'*************************
%>
<div role="tabpanel" class="tab-pane fade" id="socialmedia">


	<div class="table-responsive">
            <table class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="20%">Social Media Platform</th>
				  <th width="80%">Social Media Link</th>
                 </tr>
              </thead>

             <tbody class='searchable-socialmedia'>

				<%
				SQLsocialMedia = "SELECT * FROM PR_ProspectSocialMedia where ProspectIntRecID = " & InternalRecordIdentifier & " ORDER BY SocialMediaPlatform DESC, SocialMediaLink"

				
				Set cnnMedia = Server.CreateObject("ADODB.Connection")
				cnnMedia.open (Session("ClientCnnString"))
				Set rsMedia = Server.CreateObject("ADODB.Recordset")
				rsMedia.CursorLocation = 3 
				Set rsMedia = cnnMedia.Execute(SQLsocialMedia)
				
				If not rsMedia.EOF Then
				
					Do While Not rsMedia.EOF
					
						  	Response.Write("<tr>")
							Response.Write("<td><img src=""../img/socialmedia-icons/"&rsMedia("SocialMediaPlatform")&".png"">&nbsp;" & rsMedia("SocialMediaPlatform") & "</td>")
							Response.Write("<td>" & rsMedia("SocialMediaLink") & "</td>")
			 				Response.Write("</tr>")
  
						rsMedia.MoveNext						
					Loop
				End If
				Set rsMedia = Nothing
				cnnMedia.Close
				Set cnnMedia = Nothing
				%>


			</tbody>
		</table>
	</div>
</div>

<%'***************************
' **** eof Social Media Tab****
'*****************************
%>