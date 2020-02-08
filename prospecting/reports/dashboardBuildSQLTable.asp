<%
'Response.Write ("OK, you got here")


	
Set cnnBuildSQL = Server.CreateObject("ADODB.Connection")
cnnBuildSQL.CursorLocation = adUseClient
cnnBuildSQL.open (Session("ClientCnnString"))

Set rsBuildSQL = Server.CreateObject("ADODB.Recordset")
rsBuildSQL.CursorLocation = 3 
rsBuildSQL.CursorType = 3

Set cnnBuildSQL2 = Server.CreateObject("ADODB.Connection")
cnnBuildSQL2.CursorLocation = adUseClient
cnnBuildSQL2.open (Session("ClientCnnString"))

Set rsBuildSQL2 = Server.CreateObject("ADODB.Recordset")
rsBuildSQL2.CursorLocation = 3 
rsBuildSQL2.CursorType = 3


Set cnnBuildSQL3 = Server.CreateObject("ADODB.Connection")
cnnBuildSQL3.CursorLocation = adUseClient
cnnBuildSQL3.open (Session("ClientCnnString"))

Set cnnBuildSQL4 = Server.CreateObject("ADODB.Connection")
cnnBuildSQL4.CursorLocation = adUseClient
cnnBuildSQL4.open (Session("ClientCnnString"))

Set cnnBuildSQL5 = Server.CreateObject("ADODB.Connection")
cnnBuildSQL5.CursorLocation = adUseClient
cnnBuildSQL5.open (Session("ClientCnnString"))

Set cnnBuildSQL6 = Server.CreateObject("ADODB.Connection")
cnnBuildSQL6.CursorLocation = adUseClient
cnnBuildSQL6.open (Session("ClientCnnString"))



SQLBuildSQL = "DELETE FROM PR_DashboardDetailsUQ"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 


'Start by moving everything into the temp table
SQLBuildSQL = "INSERT INTO PR_DashboardDetailsUQ (ProspectRecID, OwnerUserNo, LeadSourceNumber, CreatedDate, CreatedByUserNo) "
SQLBuildSQL = SQLBuildSQL & "SELECT InternalRecordIdentifier, OwnerUserNo, LeadSourceNumber, CreatedDate, CreatedByUserNo FROM PR_Prospects WHERE Pool='Dead'"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 



'Now fill in all the stages
SQLBuildSQL = "SELECT ProspectRecId FROM PR_DashboardDetailsUQ ORDER BY ProspectRecId"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 

If Not rsBuildSQL.EOF Then

	
	Set rsStages = Server.CreateObject("ADODB.Recordset")
	rsStages.CursorLocation = 3 
	rsStages.CursorType = 3

	Do While NOT rsBuildSQL.EOF
	

		SQLStages = "SELECT TOP 1 * FROM PR_ProspectStages WHERE ProspectRecID = " & rsBuildSQL("ProspectRecId") & " ORDER BY RecordCreationDateTime DESC"
		Set rsStages = cnnBuildSQL2.Execute(SQLStages) 
		
		If Not rsStages.EOF Then
		
			SQLBuildSQL2 = "UPDATE PR_DashboardDetailsUQ Set LastStageNumber = " & rsStages("StageRecID")  
			SQLBuildSQL2 = SQLBuildSQL2 & " ,LastStageIntRecID = " & rsStages("InternalRecordIdentifier") 
			SQLBuildSQL2 = SQLBuildSQL2 & " WHERE ProspectRecID = " & rsBuildSQL("ProspectRecID")

			Set rsBuildSQL2 = cnnBuildSQL3.Execute(SQLBuildSQL2)
			 
		End If
	
		rsBuildSQL.MoveNext
	Loop
	
	Set rsStages = Nothing
	
End IF

'Get rid of the stages we dont need
SQLBuildSQL2 = "DELETE FROM PR_DashboardDetailsUQ WHERE LastStageNumber <> 0 AND  LastStageNumber <> 1"
Set rsBuildSQL2 = cnnBuildSQL.Execute(SQLBuildSQL2)


'Now fill in all the reasons


SQLBuildSQL = "SELECT InternalRecordIdentifier,LastStageIntRecID FROM PR_DashboardDetailsUQ "

Set rsBuildSQL3 = Server.CreateObject("ADODB.Recordset")
rsBuildSQL3.CursorLocation = 3 
rsBuildSQL3.CursorType = 3


Set rsBuildSQL4 = Server.CreateObject("ADODB.Recordset")
rsBuildSQL4.CursorLocation = 3 
rsBuildSQL4.CursorType = 3


Set rsBuildSQL3 = cnnBuildSQL4.Execute(SQLBuildSQL) 

If Not rsBuildSQL3.EOF Then

	Set rsReasons = Server.CreateObject("ADODB.Recordset")
	rsReasons.CursorLocation = 3 
	rsReasons.CursorType = 3

	Do While NOT rsBuildSQL3.EOF
	
		SQLReasons = "SELECT InternalRecordIdentifier,ReasonRecID FROM PR_ProspectReasons WHERE ProspectStagesRecID = " & rsBuildSQL3("LastStageIntRecID")
		Set rsReasons = cnnBuildSQL5.Execute(SQLReasons) 

		If Not rsReasons.EOF Then
		
			SQLBuildSQL2 = "UPDATE PR_DashboardDetailsUQ Set LastReasonNumber = " & rsReasons("ReasonRecID")  
			SQLBuildSQL2 = SQLBuildSQL2 & " WHERE InternalRecordIdentifier = " & rsBuildSQL3("InternalRecordIdentifier")
			Set rsBuildSQL4 = cnnBuildSQL6.Execute(SQLBuildSQL2) 
		
		End If
	
		rsBuildSQL3.MoveNext
	Loop
	
	Set rsReasons = Nothing
	
End IF

'Now work on the summary table
SQLBuildSQL = "DELETE FROM PR_DashboardSummaryByOwnerUQ"
Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 

'Get all reasons
Set rsAllReasons = Server.CreateObject("ADODB.Recordset")
rsAllReasons.CursorLocation = 3 
rsAllReasons.CursorType = 3
SQLAllReasons = "SELECT DISTINCT LastReasonNumber, LastStageNumber FROM PR_DashboardDetailsUQ"
Set rsAllReasons = cnnBuildSQL2.Execute(SQLAllReasons) 

'Get all owners
Set rsAllOwners = Server.CreateObject("ADODB.Recordset")
rsAllOwners.CursorLocation = 3 
rsAllOwners.CursorType = 3
SQLAllOwners = "SELECT DISTINCT OwnerUserNo FROM PR_DashboardDetailsUQ"
Set rsAllOwners = cnnBuildSQL3.Execute(SQLAllOwners) 

Set rsSummary1 = Server.CreateObject("ADODB.Recordset")
rsSummary1.CursorLocation = 3 
rsSummary1.CursorType = 3

If NOT rsAllReasons.EOF Then

	Do While Not rsAllReasons.EOF

		ReasonToDo = rsAllReasons("LastReasonNumber")
		
			Do While Not rsAllOwners.EOF
			
					SQLSummary1 = "INSERT INTO PR_DashboardSummaryByOwnerUQ (OwnerUserNo,ReasonNo,NumberOfProspects,LastStageNumber) VALUES (" & rsAllOwners("OwnerUserNo")
					SQLSummary1 = SQLSummary1 & ", " & ReasonToDo & ",0," & rsAllReasons("LastStageNumber") & ")"
					Set rsSummary = cnnBuildSQL4.Execute(SQLSummary1) 
			
				rsAllOwners.MoveNext
			Loop

			rsAllOwners.MoveFirst
			
		rsAllReasons.MoveNext
	Loop
	
End If

'Now fill in all the actual counts

Set rsTmp1 = Server.CreateObject("ADODB.Recordset")
rsTmp1.CursorLocation = 3 
rsTmp1.CursorType = 3

SQLSummary1 = "SELECT * FROM PR_DashboardSummaryByOwnerUQ"
Set rsSummary = cnnBuildSQL4.Execute(SQLSummary1) 

If NOT rsSummary.EOF Then

	Do While Not rsSummary.EOF

		SQLtmp = "SELECT COUNT(*) AS Expr1 FROM PR_DashboardDetailsUQ WHERE OwnerUserno = " & rsSummary("OwnerUserNo") & " AND LastReasonNumber  = " & rsSummary("ReasonNo") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
		'Response.Write(SQLtmp & "<BR>")
		Set rsTmp1 = cnnBuildSQL4.Execute(SQLtmp)	
		NOP = rsTmp1("Expr1")
		
		SQLtmp = "UPDATE PR_DashboardSummaryByOwnerUQ SET  NumberOfProspects = " & NOP & " WHERE OwnerUserno = " & rsSummary("OwnerUserNo") & " AND ReasonNo = " & rsSummary("ReasonNo") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
		Set rsTmp1 = cnnBuildSQL4.Execute(SQLtmp)	
			
		rsSummary.MoveNext
	Loop
	
End If

'************************************
'Rebuild PR_DashboardSummaryByLSourceUQ
'************************************
Set cnnSummary = Server.CreateObject("ADODB.Connection")
cnnSummary.CursorLocation = adUseClient
cnnSummary.open (Session("ClientCnnString"))

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")
objRecordSet.CursorLocation = 3
objConnection.Open (Session("ClientCnnString"))
objSQL = "SELECT * FROM sysobjects Where Name= 'PR_DashboardSummaryByLSourceUQ' AND xType= 'U'"
objRecordSet.Open objSQL,objConnection

If objRecordset.RecordCount = 0  Then
	'Not there, we must create it
	SQLSummary = "CREATE TABLE PR_DashboardSummaryByLSourceUQ "
	SQLSummary = SQLSummary & "	( "
	SQLSummary = SQLSummary & "     [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
	SQLSummary = SQLSummary & "     [RecordCreationDateTime] [datetime] NULL, "
	SQLSummary = SQLSummary & "     [LeadSourceNumber] [int] NULL, "
	SQLSummary = SQLSummary & "     [ReasonNo] [int] NULL, "
	SQLSummary = SQLSummary & "     [NumberOfProspects] [int] NULL, "
	SQLSummary = SQLSummary & "     [LastStageNumber] [int] NULL "
	SQLSummary = SQLSummary & "     ) ON [PRIMARY] "
	Set rsSummary = cnnSummary.Execute(SQLSummary)
Else
	'It is there, so just empty it
	SQLSummary = "DELETE FROM PR_DashboardSummaryByLSourceUQ"
	Set rsSummary = cnnSummary.Execute(SQLSummary)
End If
Set objRecordSet = Nothing
objConnection.Close
Set objConnection = Nothing

'Get all reasons
Set rsAllReasons = Server.CreateObject("ADODB.Recordset")
rsAllReasons.CursorLocation = 3 
rsAllReasons.CursorType = 3
SQLAllReasons = "SELECT DISTINCT LastReasonNumber, LastStageNumber FROM PR_DashboardDetailsUQ"
Set rsAllReasons = cnnBuildSQL2.Execute(SQLAllReasons) 

'Get all lead sources
Set rsAllLSource = Server.CreateObject("ADODB.Recordset")
rsAllLSource.CursorLocation = 3 
rsAllLSource.CursorType = 3
SQLAllLSource = "SELECT DISTINCT LeadSourceNumber FROM PR_DashboardDetailsUQ"
Set rsAllLSource = cnnBuildSQL3.Execute(SQLAllLSource) 

'Initial population of the table
Set rsSummary1 = Server.CreateObject("ADODB.Recordset")
rsSummary1.CursorLocation = 3 
rsSummary1.CursorType = 3

If NOT rsAllReasons.EOF Then

	Do While Not rsAllReasons.EOF

		ReasonToDo = rsAllReasons("LastReasonNumber")
		
			Do While Not rsAllLSource.EOF
			
					SQLSummary1 = "INSERT INTO PR_DashboardSummaryByLSourceUQ (LeadSourceNumber,ReasonNo,NumberOfProspects,LastStageNumber) VALUES (" & rsAllLSource("LeadSourceNumber")
					SQLSummary1 = SQLSummary1 & ", " & ReasonToDo & ",0," & rsAllReasons("LastStageNumber") & ")"
					Set rsSummary = cnnSummary.Execute(SQLSummary1) 
			
				rsAllLSource.MoveNext
			Loop

			rsAllLSource.MoveFirst
			
		rsAllReasons.MoveNext
	Loop
	
End If
Set rsSummary = Nothing

'Now fill in all the actual counts

Set rsTmp1 = Server.CreateObject("ADODB.Recordset")
rsTmp1.CursorLocation = 3 
rsTmp1.CursorType = 3

SQLSummary1 = "SELECT * FROM PR_DashboardSummaryByLSourceUQ"
Set rsSummary = cnnBuildSQL4.Execute(SQLSummary1) 

If NOT rsSummary.EOF Then

	Do While Not rsSummary.EOF

		SQLtmp = "SELECT COUNT(*) AS Expr1 FROM PR_DashboardDetailsUQ WHERE LeadSourceNumber = " & rsSummary("LeadSourceNumber") & " AND LastReasonNumber  = " & rsSummary("ReasonNo") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
		Set rsTmp1 = cnnSummary.Execute(SQLtmp)	
		NOP = rsTmp1("Expr1")
		
		SQLtmp = "UPDATE PR_DashboardSummaryByLSourceUQ SET  NumberOfProspects = " & NOP & " WHERE LeadSourceNumber = " & rsSummary("LeadSourceNumber") & " AND ReasonNo = " & rsSummary("ReasonNo") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
		Set rsTmp1 = cnnSummary.Execute(SQLtmp)	
			
		rsSummary.MoveNext
	Loop
	
End If

cnnSummary.Close
Set cnnSummary = Nothing
%>
