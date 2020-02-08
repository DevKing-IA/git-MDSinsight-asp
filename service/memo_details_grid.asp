<%'In order for this to work, you MUST set MDG_MemoNumber variable
'this include file is self contained & will lookup all it's own information



Set cnnMDG = Server.CreateObject("ADODB.Connection")
cnnMDG.open Session("ClientCnnString")
Set rsMDG = Server.CreateObject("ADODB.Recordset")
rsMDG.CursorLocation = 3 


SQLMDG = "SELECT SubmissionDateTime, MemoNumber, RecordSubType AS MemoStage, UserNoOfServiceTech AS UserNoSubmittingRecord, ReleasedNotes AS Remarks "
SQLMDG = SQLMDG & "FROM " & MUV_Read("SQL_Owner") & ".FS_ServiceMemos "
SQLMDG = SQLMDG & "WHERE MemoNumber = '" & MDG_MemoNumber & "' "
SQLMDG = SQLMDG & "UNION "
SQLMDG = SQLMDG & "SELECT SubmissionDateTime, MemoNumber, MemoStage, UserNoSubmittingRecord, Remarks "
SQLMDG = SQLMDG & "FROM " & MUV_Read("SQL_Owner") & ".FS_ServiceMemosDetail "
SQLMDG = SQLMDG & "WHERE MemoNumber = '" & MDG_MemoNumber & "' "
SQLMDG = SQLMDG & "ORDER BY SubmissionDateTime"

'response.write(SQLMDG)

Set rsMDG = cnnMDG.Execute(SQLMDG)

	If not rsMDG.eof then %> 
<h2 id="detail_grid"> </h2>	
		<div class="row">
			<div class="table-responsive col-lg-12">
				<table class="table table-striped"> 
					<thead> 
						<tr> 
							<th class="date-col">Date/Time</th> 
							<th class="stage-col">Status/Stage</th> 
							<th class="user-col">User</th> 
							<th>Notes</th>
						</tr> 
					</thead> 
				
					<tbody> 
						<%
						Do While Not rsMDG.EOF
				
								Response.Write("<tr>")
								Response.Write("<td>" & rsMDG("SubmissionDateTime") & "</td>") 
								Response.Write("<td>" & rsMDG("MemoStage") & "</td>") 
								If rsMDG("MemoStage") = "HOLD" then ' Always show SYSTEM on hold entries
									Response.Write("<td>" & GetUserDisplayNameByUserNo(0) & "</td>")
								Else
									Response.Write("<td>" & GetUserDisplayNameByUserNo(rsMDG("UserNoSubmittingRecord")) & "</td>")
								End If
								If advancedDispatchIsOn() Then
									If rsMDG("MemoStage") <> "HOLD" AND rsMDG("MemoStage") <> "OPEN" AND rsMDG("MemoStage") <> "CLOSE" AND rsMDG("MemoStage") <> "CANCEL" Then
										Response.Write("<td>" & rsMDG("Remarks") & "</td>") 
									ElseIF rsMDG("MemoStage") = "HOLD" Then
										HoldReason = ""
										If rsMDG("MemoStage") = "HOLD" Then
												Set cnnRems = Server.CreateObject("ADODB.Connection")
												cnnRems.open Session("ClientCnnString")
												SQLstatus = "SELECT HoldReason FROM FS_ServiceMemos WHERE MemoNumber = '" & MDG_MemoNumber & "' AND RecordSubType='HOLD'"
												Set rsRems = Server.CreateObject("ADODB.Recordset")
												rsRems.CursorLocation = 3 
												Set rsRems = cnnRems.Execute(SQLstatus )
												If not rsRems.eof then 
													HoldReason = rsRems("HoldReason")		
												End If
												set rsRems = Nothing
												set cnnRems= Nothing
										End If
										Response.Write("<td>" & HoldReason  & "</td>") 
									Else
										If rsMDG("MemoStage") = "CANCEL" or rsMDG("MemoStage") = "CLOSE" or rsMDG("MemoStage") = "OPEN" Then ' Get the notes from a different field
											Set cnnRems = Server.CreateObject("ADODB.Connection")
											cnnRems.open Session("ClientCnnString")
											SQLstatus = "SELECT * FROM FS_ServiceMemos WHERE MemoNumber = '" & MDG_MemoNumber & "' AND RecordSubType='" & rsMDG("MemoStage") & "'"
											Set rsRems = Server.CreateObject("ADODB.Recordset")
											rsRems.CursorLocation = 3 
											Set rsRems = cnnRems.Execute(SQLstatus )
											If not rsRems.eof then 
												IF rsMDG("MemoStage") <> "CLOSE" Then
													result = rsRems("ProblemDescription")
												Else
													result = rsRems("ServiceNotesFromTech")												
												End IF
											End If
											set rsRems = Nothing
											set cnnRems= Nothing
											Response.Write("<td>" & result & "</td>") 
										Else
											Response.Write("<td>&nbsp;</td>") 
										End If
									End If
								Else
									If rsMDG("MemoStage") <> "HOLD" Then 
										Response.Write("<td>" & rsMDG("Remarks") & "</td>") 								
									End If
								End If
								Response.Write("</tr>")				
				
							rsMDG.movenext
						Loop
					End IF	
				
					set rsMDG = Nothing
					cnnMDG.Close
					set cnnMDG = Nothing
					%>
					</tbody> 
		</table>
	</div>
</div>