<%
'Response.Write ("OK, you got here")

	
Set cnnBuildSQL = Server.CreateObject("ADODB.Connection")
cnnBuildSQL.open (Session("ClientCnnString"))
Set rsBuildSQL = Server.CreateObject("ADODB.Recordset")
rsBuildSQL.CursorLocation = 3 
Set rsBuildSQL2 = Server.CreateObject("ADODB.Recordset")
rsBuildSQL2.CursorLocation = 3 
Set cnnBuildSQL2 = Server.CreateObject("ADODB.Connection")
cnnBuildSQL2.open (Session("ClientCnnString"))


SQLBuildSQL = "DELETE FROM PR_DashboardDetailsUQ"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 


'Start by moving everything into the temp table
SQLBuildSQL = "INSERT INTO PR_DashboardDetailsUQ (ProspectRecID, OwnerUserNo, LeadSourceNumber, CreatedDate, CreatedByUserNo) "
SQLBuildSQL = SQLBuildSQL & "SELECT InternalRecordIdentifier, OwnerUserNo, LeadSourceNumber, CreatedDate, CreatedByUserNo FROM PR_Prospects WHERE Pool='Dead'"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 


On error resume next
SQLBuildSQL = "DROP TABLE PR_DashboardDetailsUQ2"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 
On error goto 0

SQLBuildSQL = "SELECT * INTO PR_DashboardDetailsUQ2 FROM PR_DashboardDetailsUQ"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 




'Now fill in all the stages
SQLBuildSQL = "SELECT ProspectRecId FROM PR_DashboardDetailsUQ ORDER BY ProspectRecId"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 

If Not rsBuildSQL.EOF Then

response.write("<br><br><br><br>" & Now()&"<br>")
	Set rsStages = Server.CreateObject("ADODB.Recordset")
	rsStages.CursorLocation = 3 

	Do While NOT rsBuildSQL.EOF
	x=x+1
		'SQLStages = "SELECT * FROM PR_ProspectStages WHERE ProspectRecID = " & rsBuildSQL("ProspectRecId") & " ORDER BY RecordCreationDateTime DESC"
		'Set rsStages = cnnBuildSQL.Execute(SQLStages) 
		
		'If Not rsStages.EOF Then
		
			'SQLBuildSQL2 = "UPDATE PR_DashboardDetailsUQ Set LastStageNumber = " & rsStages("StageRecID")  
			'SQLBuildSQL2 = SQLBuildSQL2 & " ,LastStageIntRecID = " & rsStages("InternalRecordIdentifier") 
			'SQLBuildSQL2 = SQLBuildSQL2 & " WHERE ProspectRecID = " & rsBuildSQL("ProspectRecID")
			
			SQLBuildSQL2 = "UPDATE PR_DashboardDetailsUQ Set LastStageNumber =  0   "
			SQLBuildSQL2 = SQLBuildSQL2 & " ,LastStageIntRecID = 1000 "
			SQLBuildSQL2 = SQLBuildSQL2 & " WHERE ProspectRecID =  300"


			Set rsBuildSQL2 = cnnBuildSQL2.Execute(SQLBuildSQL2) 
		Response.write(SQLBuildSQL2 & "<br>")
	'	End If
	
		rsBuildSQL.MoveNext

		Response.write(SQLStages  & "<br>")
		'if x > 100 then response.end
	Loop
	
	Set rsStages = Nothing
	
End IF

'Get rid of the stages we dont need
SQLBuildSQL2 = "DELETE FROM PR_DashboardDetailsUQ WHERE LastStageNumber <> 0"
Set rsBuildSQL2 = cnnBuildSQL.Execute(SQLBuildSQL2)

response.write("<br><br><br><br>" & Now()&"<br>")
response.end

'Now fill in all the reasons

'Now fill in all the stages
SQLBuildSQL = "SELECT * FROM PR_DashboardDetailsUQ "
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 

If Not rsBuildSQL.EOF Then

	Set rsReasons = Server.CreateObject("ADODB.Recordset")
	rsReasons.CursorLocation = 3 

	Do While NOT rsBuildSQL.EOF
	
		SQLReasons = "SELECT TOP 1 * FROM PR_ProspectStages WHERE ProspectRecID = " & rsBuildSQL("ProspectRecID") & " ORDER BY RecordCreationDateTime DESC"
		Set rsStages = cnnBuildSQL.Execute(SQLStages) 

		If Not rsStages.EOF Then
		
			SQLBuildSQL2 = "UPDATE PR_DashboardDetailsUQ Set LastStageNumber = " & rsStages("StageRecID")  
			SQLBuildSQL2 = SQLBuildSQL2 & " ,LastStageIntRecID = " & rsStages("InternalRecordIdentifier") 
			SQLBuildSQL2 = SQLBuildSQL2 & " WHERE ProspectRecID = " & rsBuildSQL("ProspectRecID")
			Set rsBuildSQL2 = cnnBuildSQL.Execute(SQLBuildSQL2) 
		
		End If
	
		rsBuildSQL.MoveNext
	Loop
	
	Set rsReasons = Nothing
	
End IF




Set rsBuildSQL = Nothing
cnnBuildSQL.Close
Set cnnBuildSQL = Nothing

%>
