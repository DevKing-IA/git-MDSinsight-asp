<div role="tabpanel" class="tab-pane fade in" id="opportunity">
	  

	<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
	    <input id="filter-opportunity" type="text" class="form-control filter-search-width" placeholder="Type here...">
	</div>

	<div class="table-responsive">
            <table class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="10%">Due Date</th>
 				  <th width="25%">Activity</th>
 				  <th width="15%">Status</th>
   				  <th width="40%">Notes</th>
   				  <th width="10%">User</th>
                 </tr>
              </thead>

             <tbody class='searchable-opportunity'>

				<%
				SQLPRActivities = "SELECT * FROM PR_ProspectActivities where ProspectRecID = " & InternalRecordIdentifier & " ORDER BY RecordCreationDateTime Desc"
				
				Set cnnPRActivities = Server.CreateObject("ADODB.Connection")
				cnnPRActivities.open (Session("ClientCnnString"))
				Set rsPRActivities = Server.CreateObject("ADODB.Recordset")
				rsPRActivities.CursorLocation = 3 
				Set rsPRActivities = cnnPRActivities.Execute(SQLPRActivities)
				
				If not rsPRActivities.EOF Then
				
					Do While Not rsPRActivities.EOF
					
						  	Response.Write("<tr>")
							Response.Write("<td>" & FormatDateTime(rsPRActivities("ActivityDueDate"),2) & " " & FormatDateTime(rsPRActivities("ActivityDueDate"),3) & "</td>")
							Response.Write("<td>" & GetActivityByNum(rsPRActivities("ActivityRecID")) & "</td>")
							Response.Write("<td>" & rsPRActivities("Status") & "</td>")
							Response.Write("<td>" & rsPRActivities("Notes") & "</td>")
							If rsPRActivities("StatusChangedByUserNo") <> "" Then
								Response.Write("<td>" & GetUserDisplayNameByUserNo(rsPRActivities("StatusChangedByUserNo")) & "</td>")
							Else
								Response.Write("<td>&nbsp;</td>")	
							End If
			 				Response.Write("</tr>")
  
						rsPRActivities.MoveNext						
					Loop
				End If
				Set rsPRActivities = Nothing
				cnnPRActivities.Close
				Set cnnPRActivities = Nothing
				%>


			</tbody>
		</table>
	</div>
</div>

