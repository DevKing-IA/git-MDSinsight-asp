<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<script type="text/javascript">
	function closeme() {
		window.open('', '_parent', '');
		window.close();  }
</script>
<%
Server.ScriptTimeout = 2500

mondayOfThisWeek = DateAdd("d", -((Weekday(date()) + 7 - 2) Mod 7), date())
mondayOfLastWeek = DateAdd("ww",-1,mondayOfThisWeek)
yesterday = DateAdd("d",-1, date())


'Designed to be launched via a scheduled process (Win Task Scheduler)
'Self contained page will build the dashboard tables for Prospecting
'Usage = "http://{xxx}.{domain}.com/directLaunch/prospecting/BuildDashboardTables.asp?runlevel=run_now

'The runlevel parameter is inconsequential to the operation 
'of the page. It is only used so that the page will not run
'if it is loaded via an unexpected method (spiders, etc)
If Request.QueryString("runlevel") <> "run_now" then response.end

'baseURL should alwats have a trailing /slash, just in case, handle either way
If right(baseURL,1)="/" Then maildomain = Left(right(baseURL,len(baseURL)-7),len(right(baseURL,len(baseURL)-7))-1) Else maildomain = right(baseURL,len(baseURL)-7)

'This single page loops through and handles alerts for ALL clients
SQL = "SELECT * FROM tblServerInfo WHERE Active = 1"

Set TopConnection = Server.CreateObject("ADODB.Connection")
Set TopRecordset = Server.CreateObject("ADODB.Recordset")
TopConnection.Open InsightCnnString
'Open the recordset object executing the SQL statement and return records
TopRecordset.Open SQL,TopConnection,3,3

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and exit
Response.Write("Top of loop<br>")

If Not TopRecordset.Eof Then

	Do While Not TopRecordset.EOF

		PassPhrase = TopRecordset.Fields("directLaunchPassphrase")
		ClientKey = TopRecordset.Fields("clientkey")

		Response.Write("******** Processing " & ClientKey  & "************<br>")
		Call SetClientCnnString
		
		Session("ClientCnnString") = MUV_READ("ClientCnnString") ' Until session vars are gone, then delete this
		
		If MUV_READ("cnnStatus") = "OK" Then 'AND Instr(Ucase(ClientKey),"D") = 0 Then ' else it loops and excludes dev client keys
		
			If MUV_READ("ProspectingModule") = "Enabled" Then 
		
				'Each dashboard build is handled individually for each client
				
				'****************************************
				'Begin Build Prospecting Dashboard Tables
				'****************************************
				 Response.Write("Begin Build Prospecting Dashboard Tables<br>")
				'******************************************
					
					Response.Write ("OK, you got here")
					
					Set cnnTechInfo = Server.CreateObject("ADODB.Connection")
					cnnTechInfo.open (MUV_Read("ClientCnnString"))
					Set rsTechInfo = Server.CreateObject("ADODB.Recordset")
					rsTechInfo.CursorLocation = 3 
		
					'See if these fields are in SC_TechInfo & add them if not there
					SQL_TechInfo = "SELECT COL_LENGTH('SC_TechInfo', 'PRDashboardRebuild_Start') AS IsItThere"
					Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
					If IsNull(rsTechInfo("IsItThere")) Then
						Response.Write("<br><strong><font color='blue'>The column PRDashboardRebuild_Start was not defined in SC_TechInfo - Adding the column for clientID " & ClientKey & "</font></strong><br>")
						SQL_TechInfo = "ALTER TABLE SC_TechInfo ADD PRDashboardRebuild_Start datetime NULL"
						Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
					End If
					
					SQL_TechInfo = "SELECT COL_LENGTH('SC_TechInfo', 'PRDashboardRebuild_Finish') AS IsItThere"
					Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
					If IsNull(rsTechInfo("IsItThere")) Then
						Response.Write("<br><strong><font color='blue'>The column PRDashboardRebuild_Finish was not defined in SC_TechInfo - Adding the column for clientID " & ClientKey & "</font></strong><br>")
						SQL_TechInfo = "ALTER TABLE SC_TechInfo ADD PRDashboardRebuild_Finish datetime NULL"
						Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
					End If
	
					SQL_TechInfo = "SELECT COL_LENGTH('SC_TechInfo', 'PRDashboardRebuild_LastStatus') AS IsItThere"
					Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
					If IsNull(rsTechInfo("IsItThere")) Then
						Response.Write("<br><strong><font color='blue'>The column PRDashboardRebuild_LastStatus was not defined in SC_TechInfo - Adding the column for clientID " & ClientKey & "</font></strong><br>")
						SQL_TechInfo = "ALTER TABLE SC_TechInfo ADD PRDashboardRebuild_LastStatus varchar(255) NULL"
						Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
					End If
	
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
					
					'See if the tables exist for this client
					SQLBuildSQL  = "SELECT * FROM sysobjects Where Name= 'PR_DashboardDetailsUQ_LastWeek' AND xType= 'U'"
					rsBuildSQL.Open SQLBuildSQL,cnnBuildSQL
					
					SkipThisClient = True
					
					If rsBuildSQL.RecordCount = 0 Then 
						SkipThisClient = True
						Response.Write("<br>The table PR_DashboardDetailsUQ_LastWeek does not exist for " & ClientKey & " skipping this client<br>")
					Else
						SkipThisClient = False
					End If
						
					rsBuildSQL.Close	
	
					If SkipThisClient <> True Then 
					
						Response.Write("<br><strong><font color='green'>Begin rebuild for " & ClientKey & "</font></strong><br>")
						
						SQL_TechInfo = "UPDATE SC_TechInfo SET PRDashboardRebuild_Start = getdate()"
						Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
	
						SQL_TechInfo = "UPDATE SC_TechInfo SET PRDashboardRebuild_LastStatus = 'Started'"
						Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
					
						SQLBuildSQL = "DELETE FROM PR_DashboardDetailsUQ_LastWeek"
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 
						
						
						'Start by moving everything into the temp table
						SQLBuildSQL = "INSERT INTO PR_DashboardDetailsUQ_LastWeek (ProspectRecID, OwnerUserNo, LeadSourceNumber, CreatedDate, CreatedByUserNo) "
						SQLBuildSQL = SQLBuildSQL & "SELECT InternalRecordIdentifier, OwnerUserNo, LeadSourceNumber, CreatedDate, CreatedByUserNo FROM PR_Prospects WHERE Pool='Dead'"
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 
						
						
						'Now fill in all the stages
						SQLBuildSQL = "SELECT ProspectRecId FROM PR_DashboardDetailsUQ_LastWeek ORDER BY ProspectRecId"
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 
						
						If Not rsBuildSQL.EOF Then
						
							
							Set rsStages = Server.CreateObject("ADODB.Recordset")
							rsStages.CursorLocation = 3 
							rsStages.CursorType = 3
						
							Do While NOT rsBuildSQL.EOF
							
						
								SQLStages = "SELECT TOP 1 * FROM PR_ProspectStages WHERE ProspectRecID = " & rsBuildSQL("ProspectRecId") & " ORDER BY RecordCreationDateTime DESC"
								Set rsStages = cnnBuildSQL2.Execute(SQLStages) 
								
								If Not rsStages.EOF Then
								
									'If the last stage change was not within our one week period, delete it
									'Otherwise it is ok to leave it there and move on
									
									If rsStages("RecordCreationDateTime") >= mondayOfLastWeek AND rsStages("RecordCreationDateTime") <= mondayOfThisWeek Then
									
										SQLBuildSQL2 = "UPDATE PR_DashboardDetailsUQ_LastWeek Set LastStageNumber = " & rsStages("StageRecID")  
										SQLBuildSQL2 = SQLBuildSQL2 & " ,LastStageIntRecID = " & rsStages("InternalRecordIdentifier") 
										SQLBuildSQL2 = SQLBuildSQL2 & " WHERE ProspectRecID = " & rsBuildSQL("ProspectRecID")
							
										Set rsBuildSQL2 = cnnBuildSQL3.Execute(SQLBuildSQL2)
										 
									Else
									
										SQLBuildSQL2 = "DELETE FROM PR_DashboardDetailsUQ_LastWeek "
										SQLBuildSQL2 = SQLBuildSQL2 & " WHERE ProspectRecID = " & rsBuildSQL("ProspectRecID")
							
										Set rsBuildSQL2 = cnnBuildSQL3.Execute(SQLBuildSQL2)
	 
									End If
							
								End If
								
								rsBuildSQL.MoveNext
							Loop
							
							Set rsStages = Nothing
							
						End IF
						
						'Get rid of the stages we dont need
						SQLBuildSQL2 = "DELETE FROM PR_DashboardDetailsUQ_LastWeek WHERE LastStageNumber <> 0 AND  LastStageNumber <> 1"
						Set rsBuildSQL2 = cnnBuildSQL.Execute(SQLBuildSQL2)
						
						
						'Now fill in all the reasons
						
						
						SQLBuildSQL = "SELECT InternalRecordIdentifier,LastStageIntRecID FROM PR_DashboardDetailsUQ_LastWeek "
						
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
								
									SQLBuildSQL2 = "UPDATE PR_DashboardDetailsUQ_LastWeek Set LastReasonNumber = " & rsReasons("ReasonRecID")  
									SQLBuildSQL2 = SQLBuildSQL2 & " WHERE InternalRecordIdentifier = " & rsBuildSQL3("InternalRecordIdentifier")
									Set rsBuildSQL4 = cnnBuildSQL6.Execute(SQLBuildSQL2) 
								
								End If
							
								rsBuildSQL3.MoveNext
							Loop
							
							Set rsReasons = Nothing
							
						End IF
						
						'Now work on the summary table
						SQLBuildSQL = "DELETE FROM PR_DashboardSummaryByOwnerUQ_LastWeek"
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 
						
						'Get all reasons
						Set rsAllReasons = Server.CreateObject("ADODB.Recordset")
						rsAllReasons.CursorLocation = 3 
						rsAllReasons.CursorType = 3
						SQLAllReasons = "SELECT DISTINCT LastReasonNumber, LastStageNumber FROM PR_DashboardDetailsUQ_LastWeek"
						Set rsAllReasons = cnnBuildSQL2.Execute(SQLAllReasons) 
						
						'Get all owners
						Set rsAllOwners = Server.CreateObject("ADODB.Recordset")
						rsAllOwners.CursorLocation = 3 
						rsAllOwners.CursorType = 3
						SQLAllOwners = "SELECT DISTINCT OwnerUserNo FROM PR_DashboardDetailsUQ_LastWeek"
						Set rsAllOwners = cnnBuildSQL3.Execute(SQLAllOwners) 
						
						Set rsSummary1 = Server.CreateObject("ADODB.Recordset")
						rsSummary1.CursorLocation = 3 
						rsSummary1.CursorType = 3
						
						If NOT rsAllReasons.EOF Then
						
							Do While Not rsAllReasons.EOF
						
								ReasonToDo = rsAllReasons("LastReasonNumber")
								
									Do While Not rsAllOwners.EOF
									
											SQLSummary1 = "INSERT INTO PR_DashboardSummaryByOwnerUQ_LastWeek (OwnerUserNo,ReasonNo,NumberOfProspects,LastStageNumber) VALUES (" & rsAllOwners("OwnerUserNo")
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
						
						SQLSummary1 = "SELECT * FROM PR_DashboardSummaryByOwnerUQ_LastWeek"
						Set rsSummary = cnnBuildSQL4.Execute(SQLSummary1) 
						
						If NOT rsSummary.EOF Then
						
							Do While Not rsSummary.EOF
						
								SQLtmp = "SELECT COUNT(*) AS Expr1 FROM PR_DashboardDetailsUQ_LastWeek WHERE OwnerUserno = " & rsSummary("OwnerUserNo") & " AND LastReasonNumber  = " & rsSummary("ReasonNo") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
								'Response.Write(SQLtmp & "<BR>")
								Set rsTmp1 = cnnBuildSQL4.Execute(SQLtmp)	
								NOP = rsTmp1("Expr1")
								
								SQLtmp = "UPDATE PR_DashboardSummaryByOwnerUQ_LastWeek SET  NumberOfProspects = " & NOP & " WHERE OwnerUserno = " & rsSummary("OwnerUserNo") & " AND ReasonNo = " & rsSummary("ReasonNo") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
								Set rsTmp1 = cnnBuildSQL4.Execute(SQLtmp)	
									
								rsSummary.MoveNext
							Loop
							
						End If
						
						'************************************
						'Rebuild PR_DashboardSummaryByLSourceUQ_LastWeek
						'************************************
						Set cnnSummary = Server.CreateObject("ADODB.Connection")
						cnnSummary.CursorLocation = adUseClient
						cnnSummary.open (Session("ClientCnnString"))
						
						Set objConnection = CreateObject("ADODB.Connection")
						Set objRecordSet = CreateObject("ADODB.Recordset")
						objRecordSet.CursorLocation = 3
						objConnection.Open (Session("ClientCnnString"))
						objSQL = "SELECT * FROM sysobjects Where Name= 'PR_DashboardSummaryByLSourceUQ_LastWeek' AND xType= 'U'"
						objRecordSet.Open objSQL,objConnection
						
						If objRecordset.RecordCount = 0  Then
							'Not there, we must create it
							SQLSummary = "CREATE TABLE PR_DashboardSummaryByLSourceUQ_LastWeek "
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
							SQLSummary = "DELETE FROM PR_DashboardSummaryByLSourceUQ_LastWeek"
							Set rsSummary = cnnSummary.Execute(SQLSummary)
						End If
						Set objRecordSet = Nothing
						objConnection.Close
						Set objConnection = Nothing
						
						'Get all reasons
						Set rsAllReasons = Server.CreateObject("ADODB.Recordset")
						rsAllReasons.CursorLocation = 3 
						rsAllReasons.CursorType = 3
						SQLAllReasons = "SELECT DISTINCT LastReasonNumber, LastStageNumber FROM PR_DashboardDetailsUQ_LastWeek"
						Set rsAllReasons = cnnBuildSQL2.Execute(SQLAllReasons) 
						
						'Get all lead sources
						Set rsAllLSource = Server.CreateObject("ADODB.Recordset")
						rsAllLSource.CursorLocation = 3 
						rsAllLSource.CursorType = 3
						SQLAllLSource = "SELECT DISTINCT LeadSourceNumber FROM PR_DashboardDetailsUQ_LastWeek"
						Set rsAllLSource = cnnBuildSQL3.Execute(SQLAllLSource) 
						
						'Initial population of the table
						Set rsSummary1 = Server.CreateObject("ADODB.Recordset")
						rsSummary1.CursorLocation = 3 
						rsSummary1.CursorType = 3
						
						If NOT rsAllReasons.EOF Then
						
							Do While Not rsAllReasons.EOF
						
								ReasonToDo = rsAllReasons("LastReasonNumber")
								
									Do While Not rsAllLSource.EOF
									
											SQLSummary1 = "INSERT INTO PR_DashboardSummaryByLSourceUQ_LastWeek (LeadSourceNumber,ReasonNo,NumberOfProspects,LastStageNumber) VALUES (" & rsAllLSource("LeadSourceNumber")
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
						
						SQLSummary1 = "SELECT * FROM PR_DashboardSummaryByLSourceUQ_LastWeek"
						Set rsSummary = cnnBuildSQL4.Execute(SQLSummary1) 
						
						If NOT rsSummary.EOF Then
						
							Do While Not rsSummary.EOF
						
								SQLtmp = "SELECT COUNT(*) AS Expr1 FROM PR_DashboardDetailsUQ_LastWeek WHERE LeadSourceNumber = " & rsSummary("LeadSourceNumber") & " AND LastReasonNumber  = " & rsSummary("ReasonNo") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
								Set rsTmp1 = cnnSummary.Execute(SQLtmp)	
								NOP = rsTmp1("Expr1")
								
								SQLtmp = "UPDATE PR_DashboardSummaryByLSourceUQ_LastWeek SET  NumberOfProspects = " & NOP & " WHERE LeadSourceNumber = " & rsSummary("LeadSourceNumber") & " AND ReasonNo = " & rsSummary("ReasonNo") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
								Set rsTmp1 = cnnSummary.Execute(SQLtmp)	
									
								rsSummary.MoveNext
							Loop
							
						End If
						
						cnnSummary.Close
						Set cnnSummary = Nothing

'**********NEW CODE FOR QUALIFIED TABLES

Response.Write("<b>**********NEW CODE FOR QUALIFIED TABLES</b><br><br>")

						SQLBuildSQL = "DELETE FROM PR_DashboardDetailsQ_LastWeek"
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 
						
						'Start by moving everything into the temp table
						SQLBuildSQL = "INSERT INTO PR_DashboardDetailsQ_LastWeek (ProspectRecID, OwnerUserNo, LeadSourceNumber, CreatedDate, CreatedByUserNo) "
						SQLBuildSQL = SQLBuildSQL & "SELECT InternalRecordIdentifier, OwnerUserNo, LeadSourceNumber, CreatedDate, CreatedByUserNo FROM PR_Prospects WHERE Pool='Live'"
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 
						
						'Now fill in all the stages
						SQLBuildSQL = "SELECT ProspectRecId FROM PR_DashboardDetailsQ_LastWeek ORDER BY ProspectRecId"
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 
						
						If Not rsBuildSQL.EOF Then
						
							Set rsStages = Server.CreateObject("ADODB.Recordset")
							rsStages.CursorLocation = 3 
							rsStages.CursorType = 3
						
							Do While NOT rsBuildSQL.EOF
							
						
								SQLStages = "SELECT TOP 1 * FROM PR_ProspectStages WHERE ProspectRecID = " & rsBuildSQL("ProspectRecId") & " ORDER BY RecordCreationDateTime DESC"
								Set rsStages = cnnBuildSQL2.Execute(SQLStages) 
								
								If Not rsStages.EOF Then
								
									'If the last stage change was not within our one week period, delete it
									'Otherwise it is ok to leave it there and move on
									
									If rsStages("RecordCreationDateTime") >= mondayOfLastWeek AND rsStages("RecordCreationDateTime") <= mondayOfThisWeek Then

										SQLBuildSQL2 = "UPDATE PR_DashboardDetailsQ_LastWeek Set LastStageNumber = " & rsStages("StageRecID")  
										SQLBuildSQL2 = SQLBuildSQL2 & " ,LastStageIntRecID = " & rsStages("InternalRecordIdentifier") 
										SQLBuildSQL2 = SQLBuildSQL2 & " WHERE ProspectRecID = " & rsBuildSQL("ProspectRecID")
							
										Set rsBuildSQL2 = cnnBuildSQL3.Execute(SQLBuildSQL2)
										
									Else
									
 										SQLBuildSQL2 = "DELETE FROM PR_DashboardDetailsQ_LastWeek "
										SQLBuildSQL2 = SQLBuildSQL2 & " WHERE ProspectRecID = " & rsBuildSQL("ProspectRecID")
							
										Set rsBuildSQL2 = cnnBuildSQL3.Execute(SQLBuildSQL2)
										 
									End If
								
								End If
							
								rsBuildSQL.MoveNext
							Loop
							
							Set rsStages = Nothing
							
						End IF
						
						'Get rid of the stages we dont need - only keep stages that Secondary
						SQLBuildSQL2 = "DELETE FROM PR_DashboardDetailsQ_LastWeek WHERE LastStageNumber = 2 OR LastStageNumber = 1"
						Set rsBuildSQL2 = cnnBuildSQL.Execute(SQLBuildSQL2)
					
						
						'Now work on the summary table
						SQLBuildSQL = "DELETE FROM PR_DashboardSummaryByOwnerQ_LastWeek"
						Set rsBuildSQL = cnnBuildSQL.Execute(SQLBuildSQL) 
						
						
						'Get all owners
						Set rsAllOwners = Server.CreateObject("ADODB.Recordset")
						rsAllOwners.CursorLocation = 3 
						rsAllOwners.CursorType = 3
						SQLAllOwners = "SELECT DISTINCT OwnerUserNo,* FROM PR_DashboardDetailsQ_LastWeek"
						Set rsAllOwners = cnnBuildSQL3.Execute(SQLAllOwners) 
						
						Set rsSummary1 = Server.CreateObject("ADODB.Recordset")
						rsSummary1.CursorLocation = 3 
						rsSummary1.CursorType = 3
						
						If NOT rsAllOwners.EOF Then		
							Do While Not rsAllOwners.EOF
							
									SQLSummary1 = "INSERT INTO PR_DashboardSummaryByOwnerQ_LastWeek (OwnerUserNo,NumberOfProspects,LastStageNumber) VALUES (" & rsAllOwners("OwnerUserNo")
									SQLSummary1 = SQLSummary1 & ",0," & rsAllOwners("LastStageNumber") & ")"
									Set rsSummary = cnnBuildSQL4.Execute(SQLSummary1) 
							
								rsAllOwners.MoveNext
							Loop
				
							rsAllOwners.MoveFirst
						End If		
									
						'Now fill in all the actual counts
						
						Set rsTmp1 = Server.CreateObject("ADODB.Recordset")
						rsTmp1.CursorLocation = 3 
						rsTmp1.CursorType = 3
						
						SQLSummary1 = "SELECT * FROM PR_DashboardSummaryByOwnerQ_LastWeek"
						Set rsSummary = cnnBuildSQL4.Execute(SQLSummary1) 
						
						If NOT rsSummary.EOF Then
						
							Do While Not rsSummary.EOF
						
								SQLtmp = "SELECT COUNT(*) AS Expr1 FROM PR_DashboardDetailsQ_LastWeek WHERE OwnerUserno = " & rsSummary("OwnerUserNo") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
								'Response.Write(SQLtmp & "<BR>")
								Set rsTmp1 = cnnBuildSQL4.Execute(SQLtmp)	
								NOP = rsTmp1("Expr1")
								
								SQLtmp = "UPDATE PR_DashboardSummaryByOwnerQ_LastWeek SET  NumberOfProspects = " & NOP & " WHERE OwnerUserno = " & rsSummary("OwnerUserNo") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
								Set rsTmp1 = cnnBuildSQL4.Execute(SQLtmp)	
									
								rsSummary.MoveNext
							Loop
							
						End If
						
						'************************************
						'Rebuild PR_DashboardSummaryByLSourceQ_LastWeek
						'************************************
						Set cnnSummary = Server.CreateObject("ADODB.Connection")
						cnnSummary.CursorLocation = adUseClient
						cnnSummary.open (Session("ClientCnnString"))
						
						Set objConnection = CreateObject("ADODB.Connection")
						Set objRecordSet = CreateObject("ADODB.Recordset")
						objRecordSet.CursorLocation = 3
						objConnection.Open (Session("ClientCnnString"))
						objSQL = "SELECT * FROM sysobjects Where Name= 'PR_DashboardSummaryByLSourceQ_LastWeek' AND xType= 'U'"
						objRecordSet.Open objSQL,objConnection
						
						If objRecordset.RecordCount = 0  Then
							'Not there, we must create it
							SQLSummary = "CREATE TABLE PR_DashboardSummaryByLSourceQ_LastWeek "
							SQLSummary = SQLSummary & "	( "
							SQLSummary = SQLSummary & "     [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
							SQLSummary = SQLSummary & "     [RecordCreationDateTime] [datetime] NULL, "
							SQLSummary = SQLSummary & "     [LeadSourceNumber] [int] NULL, "
							SQLSummary = SQLSummary & "     [NumberOfProspects] [int] NULL, "
							SQLSummary = SQLSummary & "     [LastStageNumber] [int] NULL "
							SQLSummary = SQLSummary & "     ) ON [PRIMARY] "
							Set rsSummary = cnnSummary.Execute(SQLSummary)
						Else
							'It is there, so just empty it
							SQLSummary = "DELETE FROM PR_DashboardSummaryByLSourceQ_LastWeek"
							Set rsSummary = cnnSummary.Execute(SQLSummary)
						End If
						Set objRecordSet = Nothing
						objConnection.Close
						Set objConnection = Nothing
						
						'Get all lead sources
						Set rsAllLSource = Server.CreateObject("ADODB.Recordset")
						rsAllLSource.CursorLocation = 3 
						rsAllLSource.CursorType = 3
						SQLAllLSource = "SELECT DISTINCT LeadSourceNumber,* FROM PR_DashboardDetailsQ_LastWeek"
						Set rsAllLSource = cnnBuildSQL3.Execute(SQLAllLSource) 
						
						'Initial population of the table
						Set rsSummary1 = Server.CreateObject("ADODB.Recordset")
						rsSummary1.CursorLocation = 3 
						rsSummary1.CursorType = 3
						
						IF Not rsAllLSource.EOF Then		
							Do While Not rsAllLSource.EOF
							
									SQLSummary1 = "INSERT INTO PR_DashboardSummaryByLSourceQ_LastWeek (LeadSourceNumber,NumberOfProspects,LastStageNumber) VALUES (" & rsAllLSource("LeadSourceNumber")
									SQLSummary1 = SQLSummary1 & ",0," & rsAllLSource("LastStageNumber") & ")"
									Set rsSummary = cnnSummary.Execute(SQLSummary1) 
							
								rsAllLSource.MoveNext
							Loop
				
							rsAllLSource.MoveFirst
						End If

						Set rsSummary = Nothing
						
						'Now fill in all the actual counts
						
						Set rsTmp1 = Server.CreateObject("ADODB.Recordset")
						rsTmp1.CursorLocation = 3 
						rsTmp1.CursorType = 3
						
						SQLSummary1 = "SELECT * FROM PR_DashboardSummaryByLSourceQ_LastWeek"
						Set rsSummary = cnnBuildSQL4.Execute(SQLSummary1) 
						
						If NOT rsSummary.EOF Then
						
							Do While Not rsSummary.EOF
						
								SQLtmp = "SELECT COUNT(*) AS Expr1 FROM PR_DashboardDetailsQ_LastWeek WHERE LeadSourceNumber = " & rsSummary("LeadSourceNumber") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
								Set rsTmp1 = cnnSummary.Execute(SQLtmp)	
								NOP = rsTmp1("Expr1")
								
								SQLtmp = "UPDATE PR_DashboardSummaryByLSourceQ_LastWeek SET  NumberOfProspects = " & NOP & " WHERE LeadSourceNumber = " & rsSummary("LeadSourceNumber") & " AND LastStageNumber  = " & rsSummary("LastStageNumber")
								Set rsTmp1 = cnnSummary.Execute(SQLtmp)	
									
								rsSummary.MoveNext
							Loop
							
						End If
						
						cnnSummary.Close
						Set cnnSummary = Nothing



'******END NEW CODE FOR QUALIFIED TABLES
					
						SQL_TechInfo = "UPDATE SC_TechInfo SET PRDashboardRebuild_Finish = getdate()"
						Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
						SQL_TechInfo = "UPDATE SC_TechInfo SET PRDashboardRebuild_LastStatus = 'Finished'"
						Set rsTechInfo = cnnTechInfo.Execute(SQL_TechInfo)
		
					End If ' Endif for skipthisclient
	
	
				'****************************************
				'End Build Prospecting Dashboard Tables
				'****************************************
				 Response.Write("End Build Prospecting Dashboard Tables<br>")
				'******************************************
				
			Else
			
				Response.Write("The proepcting module is not enabled for client id " & ClientKey & " skipping the build<br>")
				
			End If ' Endif for Prospecting module enabled
			
			'******************************************	
			Response.Write("******** DONE Processing " & ClientKey  & "************<br>")
	End If	
				
	TopRecordset.movenext
	
	Loop

	TopRecordset.Close
	Set TopRecordset = Nothing
	TopConnection.Close
	Set TopConnection = Nothing
	
End If

Response.write("<script type=""text/javascript"">closeme();</script>")	
'Response.End
'*************************
'*************************
'Subs and funcs begin here


Sub SetClientCnnString

	dummy=MUV_WRITE("cnnStatus","")

	SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open InsightCnnString
	
	'Open the recordset object executing the SQL statement and return records
	Recordset.Open SQL,Connection,3,3

	
	'First lookup the ClientKey in tblServerInfo
	'If there is no record with the entered client key, close connection
	'and exit
	If Recordset.recordcount <= 0 then
		Recordset.close
		Connection.close
		set Recordset=nothing
		set Connection=nothing
	Else
		ClientCnnString = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
		ClientCnnString = ClientCnnString & ";Database=" & Recordset.Fields("dbCatalog")
		ClientCnnString = ClientCnnString & ";Uid=" & Recordset.Fields("dbLogin")
		ClientCnnString = ClientCnnString & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
		dummy = MUV_Write("ClientCnnString",ClientCnnString)
		dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
		dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
		dummy = MUV_Write("ProspectingModule",Recordset.Fields("ProspectingModule"))
		Recordset.close
		Connection.close
		dummy=MUV_WRITE("cnnStatus","OK")
	End If
End Sub



%>