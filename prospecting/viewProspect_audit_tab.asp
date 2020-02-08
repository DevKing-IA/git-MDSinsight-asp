
<%'***********************
' **** Audit Trail Tab****
'*************************
%>
<div role="tabpanel" class="tab-pane fade" id="audit">

	<div class="input-group narrow-results"> <span class="input-group-addon">Narrow Results</span>
	    <input id="filter-audit" type="text" class="form-control filter-search-width" placeholder="Type here...">
	</div>

	<div class="table-responsive">
            <table class="table table-striped table-condensed table-hover table-bordered sortable standard-font">
              <thead>
                <tr>
                  <th width="5%">Date</th>
				  <th width="5%">Time</th>
				  <th width="85%">Audit Trail Entry</th>
				  <th width="5%">User</th>
                 </tr>
              </thead>

             <tbody class='searchable-audit'>

				<%
				SQLAudit = "SELECT * FROM PR_Audit where ProspectIntRecID = " & InternalRecordIdentifier & " ORDER BY DateAndTime Desc"
				
				Set cnnAudit = Server.CreateObject("ADODB.Connection")
				cnnAudit.open (Session("ClientCnnString"))
				Set rsAudit = Server.CreateObject("ADODB.Recordset")
				rsAudit.CursorLocation = 3 
				Set rsAudit = cnnAudit.Execute(SQLAudit)
				
				If not rsAudit.EOF Then
				
					Do While Not rsAudit.EOF
					
						  	Response.Write("<tr>")
							Response.Write("<td>" & FormatDateTime(rsAudit("DateAndTime"),2) & "</td>")
							Response.Write("<td>" & FormatDateTime(rsAudit("DateAndTime"),3) & "</td>")
							Response.Write("<td>" & rsAudit("Activity") & "</td>")
							Response.Write("<td>" & GetUserDisplayNameByUserNo(rsAudit("PerformedByUserNo")) & "</td>")
			 				Response.Write("</tr>")
  
						rsAudit.MoveNext						
					Loop
				End If
				Set rsAudit = Nothing
				cnnAudit.Close
				Set cnnAudit = Nothing
				%>


			</tbody>
		</table>
	</div>
</div>

<%'***************************
' **** eof Audit Trail Tab****
'*****************************
%>