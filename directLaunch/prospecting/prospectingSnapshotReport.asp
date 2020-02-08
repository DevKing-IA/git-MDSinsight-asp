<% @ Language = VBScript %>
<%
Response.Buffer = True
Response.Expires = 0
Response.Clear

FontSizeVar = 9
PageNum = 0
NoBreak = False
PageWidth = 1450

Server.ScriptTimeout = 25000


%>
<!--#include file="../../inc/SubsAndFuncs.asp"-->
<!--#include file="../../inc/InsightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Prospecting.asp"-->
<%dummy=MUV_Write("ClientID","") 'Need this here

ClientKey = Request.QueryString("c")

SQL = "SELECT * FROM tblServerInfo where clientKey='"& ClientKey &"'"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open InsightCnnString
'Response.Write("InsightCnnString:" & InsightCnnString & "<br>")

'Open the recordset object executing the SQL statement and return records
Recordset.Open SQL,Connection,3,3
'Response.Write("SQL:" & SQL& "<br>")

'First lookup the ClientKey in tblServerInfo
'If there is no record with the entered client key, close connection
'and go back to login with QueryString
If Recordset.recordcount <= 0 then
	Recordset.close
	Connection.close
	set Recordset=nothing
	set Connection=nothing
	%>MDS Insight: Unable to connect to SQL database. The server is not available or the credentials specified are incorrect. - Prospecting weekly snapshot report<%
	Response.End
Else
	Session("ClientCnnString") = "Driver={SQL Server};Server=" & Recordset.Fields("dbServer")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Database=" & Recordset.Fields("dbCatalog")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Uid=" & Recordset.Fields("dbLogin")
	Session("ClientCnnString") = Session("ClientCnnString") & ";Pwd=" & Recordset.Fields("dbPassword") & ";"
	dummy = MUV_Write("SQL_Owner",Recordset.Fields("dbLogin"))
	dummy = MUV_Write("ClientID",Recordset.Fields("clientkey"))
	Recordset.close
	Connection.close	
End If	



'This is here so we only open it once for the whole page
Set cnn_Settings_Global = Server.CreateObject("ADODB.Connection")
cnn_Settings_Global.open (Session("ClientCnnString"))
Set rs_Settings_Global = Server.CreateObject("ADODB.Recordset")
rs_Settings_Global.CursorLocation = 3 
SQL_Settings_Global = "SELECT * FROM Settings_Global"
Set rs_Settings_Global = cnn_Settings_Global.Execute(SQL_Settings_Global)
If not rs_Settings_Global.EOF Then
	ProspSnapshotOnOff = rs_Settings_Global("ProspSnapshotOnOff")
	ProspSnapshotInsideSales = rs_Settings_Global("ProspSnapshotInsideSales")
	ProspSnapshotOutsideSales = rs_Settings_Global("ProspSnapshotOutsideSales")
	ProspSnapshotUserNos = rs_Settings_Global("ProspSnapshotUserNos")
	ProspSnapshotAdditionalEmails = rs_Settings_Global("ProspSnapshotAdditionalEmails")
	ProspSnapshotEmailSubject = rs_Settings_Global("ProspSnapshotEmailSubject")
	ProspSnapshotSalesRepDisplayUserNos = rs_Settings_Global("ProspSnapshotSalesRepDisplayUserNos")
Else
	ProspSnapshotOnOff = vbFalse
End If
Set rs_Settings_Global = Nothing
cnn_Settings_Global.Close
Set cnn_Settings_Global = Nothing


%>


<body bgcolor="#FFFFFF" text="#000000" link="#000080" vlink="#000080" alink="#000080" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">


<table border="0" width="<%=PageWidth%>" align="center">

	<tr>
		<td width="100%" align="center">
		
		

		<%
			
	 mondayOfThisWeek = DateAdd("d", -((Weekday(date()) + 7 - 2) Mod 7), date())
	 mondayOfLastWeek = DateAdd("ww",-1,mondayOfThisWeek)
	 sundayOfThisWeek = DateAdd("d",-1, mondayOfThisWeek) 
	 

	 %>
	<table border="0" width="<%=PageWidth%>" align="center">

	<tr>
		<td width="100%" align="center">
 <%
	 
	'*******************************************************
	'*** This section is the first page which prints all the
	'*** prospecting summary info
	'*******************************************************

	Call PageHeader

	LinesPerPage = 42
	
	SQL = "SELECT * FROM PR_Prospects WHERE Pool = 'LIVE'"

	Set cnnWeeklySnapshotSummary = Server.CreateObject("ADODB.Connection")
	cnnWeeklySnapshotSummary.open(Session("ClientCnnString"))
	Set rsWeeklySnapshotSummary  = Server.CreateObject("ADODB.Recordset")
	rsWeeklySnapshotSummary.CursorLocation = 3 
	rsWeeklySnapshotSummary.Open SQL, cnnWeeklySnapshotSummary 
				
	'Response.Write(SQL & "<br>")
	
	If Not rsWeeklySnapshotSummary.EOF AND ProspSnapshotSalesRepDisplayUserNos <> "" Then
	
		
		%>
		
		<!------- TEST AREA ---------------->

		
		<br><br><br>
		<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<tr>
				<td colspan="8">
				<font face="Consolas">
				<hr>
				<center><h2>Prospecting Snapshot <%= mondayOfLastWeek %> - <%= sundayOfThisWeek %></h2></center>
				<hr>
				</font>
				</td>
			</tr>
			<tr>
				<td colspan="8">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="8">&nbsp;</td>
			</tr>


			<tr>
				<%
				TotalNumberOfProspectsPreexisting = TotalNumberOfPreexistingProspectsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos)
				TotalNumberOfProspectsCreated = TotalNumberOfCreatedProspectsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos)
				TotalNumberOfWonProspects = TotalNumberOfWonProspectsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos)
				TotalNumberOfLostProspects = TotalNumberOfLostProspectsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos)
				TotalNumberOfUnqualifiedProspects = TotalNumberOfUnqualifiedProspectsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos)
	
				TotalProspectsWeeklySnapshot = TotalNumberOfProspectsPreexisting + TotalNumberOfProspectsCreated - TotalNumberOfWonProspects - TotalNumberOfLostProspects  - TotalNumberOfUnqualifiedProspects	
				%>
			
				<td colspan="2" valign="top" width="25%">
					<table width="100%">
						<tr>
							<td colspan="2" style="border-bottom: 1px solid black;">
								<font face="Consolas" style="font-size: 14pt"><strong>PROSPECTS<br>OVERVIEW</strong></font>
							</td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfProspectsPreexisting %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;Pre-Existing</font></td>
						</tr>	
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfProspectsCreated %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;Created last week</font></td>
						</tr>
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfWonProspects %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;Won last week</font></td>
						</tr>
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfLostProspects %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;Lost last week</font></td>
						</tr>
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfUnqualifiedProspects %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;Unqualified last week</font></td>
						</tr>
					
						<tr>
							<td colspan="2" align="right">
								<font face="Consolas" style="font-size: 14pt"><hr width="80%" align="left"></font>
							</td>					
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalProspectsWeeklySnapshot %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;TOTAL</font></td>
						</tr>
					</table>
				</td>	
					
					
				<td colspan="2" valign="top" width="25%">
				
					<%
					TotalNumberOfAppointmentsCreated = TotalNumberOfCreatedAppointmentsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos)
					TotalNumberOfAppointmentsCompleted = TotalNumberOfCompletedAppointmentsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos)	
					TotalNumberOfAppointmentsRescheduled = TotalNumberOfRescheduledAppointmentsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos)
					TotalNumberOfAppointmentsCancelled = TotalNumberOfCancelledAppointmentsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos)
					TotalNumberOfAppointmentsExpired = TotalNumberOfExpiredAppointmentsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos)
					%>
				
					<table width="100%">
						<tr>
							<td colspan="2" style="border-bottom: 1px solid black;">
								<font face="Consolas" style="font-size: 14pt"><strong>APPOINTMENTS DUE<br><%= mondayOfLastWeek %> - <%= sundayOfThisWeek %></strong></font>
							</td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsCreated %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;New Appts. Created</font></td>
						</tr>	
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsCompleted %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;Completed (Went On)</font></td>
						</tr>
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsRescheduled %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;Rescheduled</font></td>
						</tr>
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsCancelled %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;Cancelled</font></td>
						</tr>
						<tr>
							<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsExpired %></font></td>
							<td><font face="Consolas" style="font-size: 14pt">&nbsp;Expired</font></td>
						</tr>
					</table>
				</td>
				

				<td colspan="2" valign="top" width="25%">
					<table width="100%">
						<tr>
							<td colspan="2" style="border-bottom: 1px solid black;">
								<font face="Consolas" style="font-size: 14pt"><strong>PROSPECTS<br>LOST BY REASON</strong></font>
							</td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
					
						<%
					
						SQLProspectsLostByReason = "SELECT ReasonRecID, COUNT(ReasonRecID) AS Expr1"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " FROM  (SELECT InternalRecordIdentifier, RecordCreationDateTime, ProspectRecID, StageRecID, ReasonRecID, ReasonChangedByUserNo, ProspectStagesRecID"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " FROM  PR_ProspectReasons WHERE StageRecID = 1 "
						SQLProspectsLostByReason = SQLProspectsLostByReason & " AND CAST(RecordCreationDateTime AS DATE) >= '" & mondayOfLastWeek & "' AND CAST(RecordCreationDateTime AS DATE) <= '" & sundayOfThisWeek & "' "
						SQLProspectsLostByReason = SQLProspectsLostByReason & " AND (InternalrecordIdentifier IN"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " (SELECT MAX(InternalrecordIdentifier) AS CurrentReasonIntRecID"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " FROM  (SELECT InternalrecordIdentifier, RecordCreationDateTime, ProspectRecID, ReasonRecID, ReasonChangedByUserNo, ProspectStagesRecID"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " FROM  PR_ProspectReasons AS PR_ProspectReasons_1"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " WHERE (ProspectRecID IN"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " (SELECT InternalRecordIdentifier"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " FROM PR_Prospects"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " WHERE (CAST(CreatedDate AS DATE) <= '" & sundayOfThisWeek & "') AND "
						SQLProspectsLostByReason = SQLProspectsLostByReason & " (Pool = 'Dead') AND (OwnerUserNo IN (" & ProspSnapshotSalesRepDisplayUserNos & ")))))  AS derivedtbl_2"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " GROUP BY ProspectRecID)))  AS derivedtbl_1"
						SQLProspectsLostByReason = SQLProspectsLostByReason & " GROUP BY ReasonRecID "
						
						'Response.Write(SQLProspectsLostByReason & "<br>")
					
						Set cnnWeeklySnapshotProspectsLostByReason = Server.CreateObject("ADODB.Connection")
						cnnWeeklySnapshotProspectsLostByReason.open(Session("ClientCnnString"))
						Set rsWeeklySnapshotProspectsLostByReason  = Server.CreateObject("ADODB.Recordset")
						rsWeeklySnapshotProspectsLostByReason.CursorLocation = 3 
						rsWeeklySnapshotProspectsLostByReason.Open SQLProspectsLostByReason, cnnWeeklySnapshotProspectsLostByReason
									
						
						
						If Not rsWeeklySnapshotProspectsLostByReason.EOF Then
						
							TotalProspectsLostByReasonWeeklySnapshot = 0
							TotalNumberOfProspectsByReason = 0
							
							Do While Not rsWeeklySnapshotProspectsLostByReason.EOF
							
								ReasonIntRecID = rsWeeklySnapshotProspectsLostByReason("ReasonRecID")
								ReasonName = GetReasonByNum(ReasonIntRecID)
								TotalNumberOfProspectByReason = rsWeeklySnapshotProspectsLostByReason("Expr1")
								
								TotalProspectsLostByReasonWeeklySnapshot = TotalProspectsLostByReasonWeeklySnapshot + TotalNumberOfProspectByReason	
								
								%>									
								<tr>
									<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfProspectByReason %></font></td>
									<td><font face="Consolas" style="font-size: 14pt">&nbsp;<%= ReasonName %></font></td>
								</tr>	
								<% 
							
							rsWeeklySnapshotProspectsLostByReason.MoveNext
							Loop
							%>
							<tr>
								<td colspan="2" align="right">
									<font face="Consolas" style="font-size: 14pt"><hr width="80%" align="left"></font>
								</td>					
							<tr>
								<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalProspectsLostByReasonWeeklySnapshot %></font></td>
								<td><font face="Consolas" style="font-size: 14pt">&nbsp;TOTAL</font></td>
							</tr>
							<%
						Else
						%>
							<tr>
								<td colspan="2"><font face="Consolas" style="font-size: 14pt">No Lost Prospects The Week of <%= mondayOfLastWeek %> - <%= sundayOfThisWeek %></font></td>
							</tr>
						<%
						End If		
						%>

					</table>
				</td>
				

				<td colspan="2" valign="top" width="25%">
					<table width="100%">
						<tr>
							<td colspan="2" style="border-bottom: 1px solid black;">
								<font face="Consolas" style="font-size: 14pt"><strong>PROSPECTS<br>UNQUALIFIED BY REASON</strong></font>
							</td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						
						<%
					
						SQLProspectsUnqualifiedByReason = "SELECT ReasonRecID, COUNT(ReasonRecID) AS Expr1"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " FROM  (SELECT InternalRecordIdentifier, RecordCreationDateTime, ProspectRecID, StageRecID, ReasonRecID, ReasonChangedByUserNo, ProspectStagesRecID"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " FROM  PR_ProspectReasons WHERE StageRecID = 0 "
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " AND CAST(RecordCreationDateTime AS DATE) >= '" & mondayOfLastWeek & "' AND CAST(RecordCreationDateTime AS DATE) <= '" & sundayOfThisWeek & "' "
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " AND (InternalrecordIdentifier IN"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " (SELECT MAX(InternalrecordIdentifier) AS CurrentReasonIntRecID"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " FROM  (SELECT InternalrecordIdentifier, RecordCreationDateTime, ProspectRecID, ReasonRecID, ReasonChangedByUserNo, ProspectStagesRecID"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " FROM  PR_ProspectReasons AS PR_ProspectReasons_1"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " WHERE (ProspectRecID IN"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " (SELECT InternalRecordIdentifier"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " FROM PR_Prospects"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " WHERE (CAST(CreatedDate AS DATE) <= '" & sundayOfThisWeek & "') AND "
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " (Pool = 'Dead') AND (OwnerUserNo IN (" & ProspSnapshotSalesRepDisplayUserNos & ")))))  AS derivedtbl_2"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " GROUP BY ProspectRecID)))  AS derivedtbl_1"
						SQLProspectsUnqualifiedByReason = SQLProspectsUnqualifiedByReason & " GROUP BY ReasonRecID "
						
						'Response.Write(SQLProspectsUnqualifiedByReason & "<br>")
					
						Set cnnWeeklySnapshotProspectsUnqualifiedByReason = Server.CreateObject("ADODB.Connection")
						cnnWeeklySnapshotProspectsUnqualifiedByReason.open(Session("ClientCnnString"))
						Set rsWeeklySnapshotProspectsUnqualifiedByReason  = Server.CreateObject("ADODB.Recordset")
						rsWeeklySnapshotProspectsUnqualifiedByReason.CursorLocation = 3 
						rsWeeklySnapshotProspectsUnqualifiedByReason.Open SQLProspectsUnqualifiedByReason, cnnWeeklySnapshotProspectsUnqualifiedByReason
									
						
						
						If Not rsWeeklySnapshotProspectsUnqualifiedByReason.EOF Then
						
							TotalProspectsUnqualifiedByReasonWeeklySnapshot = 0
							TotalNumberOfProspectsByReason = 0
							
							Do While Not rsWeeklySnapshotProspectsUnqualifiedByReason.EOF
							
								ReasonIntRecID = rsWeeklySnapshotProspectsUnqualifiedByReason("ReasonRecID")
								ReasonName = GetReasonByNum(ReasonIntRecID)
								TotalNumberOfProspectByReason = rsWeeklySnapshotProspectsUnqualifiedByReason("Expr1")
								
								TotalProspectsUnqualifiedByReasonWeeklySnapshot = TotalProspectsUnqualifiedByReasonWeeklySnapshot + TotalNumberOfProspectByReason	
								
								%>									
								<tr>
									<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfProspectByReason %></font></td>
									<td><font face="Consolas" style="font-size: 14pt">&nbsp;<%= ReasonName %></font></td>
								</tr>	
								<% 
							
							rsWeeklySnapshotProspectsUnqualifiedByReason.MoveNext
							Loop
							%>
							<tr>
								<td colspan="2" align="right">
									<font face="Consolas" style="font-size: 14pt"><hr width="80%" align="left"></font>
								</td>					
							<tr>
								<td align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalProspectsUnqualifiedByReasonWeeklySnapshot %></font></td>
								<td><font face="Consolas" style="font-size: 14pt">&nbsp;TOTAL</font></td>
							</tr>
							<%
						Else
						%>
							<tr>
								<td colspan="2"><font face="Consolas" style="font-size: 14pt">No Unqualified Prospects The Week of <%= mondayOfLastWeek %> - <%= sundayOfThisWeek %></font></td>
							</tr>
						<%
						End If		
						%>
					</table>
				</td>
				
			</tr>
			<tr>
				<td colspan="8">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="8">&nbsp;</td>
			</tr>
		
		<!------- END TEST AREA ---------------->
						
			<% RowCount = 15 %>
							
		<% Else %>
			<tr>
				<td colspan="8">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="8">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="8">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="8">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="8">
					<font face="Consolas" style="font-size: 14pt">
					<hr>
					<center>There has been no <%= GetTerm("prospecting") %> activity for <%= mondayOfLastWeek %> - <%= sundayOfThisWeek %>.</center></font>
					
					<% NoBreak = True %>
				</td>
			</tr>
			<% RowCount = 0

		End If %>

		</table>
		
		<%
	
		SQLPipeline = "SELECT StageRecID, COUNT(StageRecID) AS Expr1"
		SQLPipeline = SQLPipeline & " FROM  (SELECT Distinct(ProspectRecID), InternalrecordIdentifier, RecordCreationDateTime,  StageRecID, Notes, StageChangedByUserNo"
		SQLPipeline = SQLPipeline & " FROM  PR_ProspectStages"
		SQLPipeline = SQLPipeline & " WHERE (InternalrecordIdentifier IN"
		SQLPipeline = SQLPipeline & " (SELECT MAX(InternalrecordIdentifier) AS CurrentStageIntRecID"
		SQLPipeline = SQLPipeline & " FROM  (SELECT distinct(ProspectRecID), InternalrecordIdentifier, RecordCreationDateTime,  StageRecID, Notes, StageChangedByUserNo"
		SQLPipeline = SQLPipeline & " FROM  PR_ProspectStages AS PR_ProspectStages_1"
		SQLPipeline = SQLPipeline & " WHERE (ProspectRecID IN"
		SQLPipeline = SQLPipeline & " (SELECT InternalRecordIdentifier"
		SQLPipeline = SQLPipeline & " FROM PR_Prospects"
		SQLPipeline = SQLPipeline & " WHERE (CAST(CreatedDate AS DATE) <= '" & mondayOfLastWeek & "') AND "
		'SQLPipeline = SQLPipeline & " WHERE (CAST(CreatedDate AS DATE) <= '" & sundayOfThisWeek & "') AND "
		SQLPipeline = SQLPipeline & " (Pool = 'LIVE') AND (OwnerUserNo IN (" & ProspSnapshotSalesRepDisplayUserNos & ")))))  AS derivedtbl_2"
		SQLPipeline = SQLPipeline & " GROUP BY ProspectRecID)))  AS derivedtbl_1"
		SQLPipeline = SQLPipeline & " GROUP BY StageRecID "
		
		'Response.Write(SQLPipeline & "<br>")
	
		Set cnnWeeklySnapshotSummaryPipeline = Server.CreateObject("ADODB.Connection")
		cnnWeeklySnapshotSummaryPipeline.open(Session("ClientCnnString"))
		Set rsWeeklySnapshotSummaryPipeline  = Server.CreateObject("ADODB.Recordset")
		rsWeeklySnapshotSummaryPipeline.CursorLocation = 3 
		rsWeeklySnapshotSummaryPipeline.Open SQLPipeline, cnnWeeklySnapshotSummaryPipeline
					
		
		
		If Not rsWeeklySnapshotSummaryPipeline.EOF Then
		
			TotalProspectsByStageWeeklySnapshot = 0
		%>
		
		<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<tr>
				<td colspan="5">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2" style="border-bottom: 1px solid black;" width="30%">
					<font face="Consolas" style="font-size: 14pt"><strong>PIPELINE BY STAGE</strong></font>
				</td>
				<td style="border-bottom: 1px solid white;" width="10%">&nbsp;</td>
				<td style="border-bottom: 1px solid white;" colspan="2" width="40%">&nbsp;</td>							
			</tr>
			
			<tr>
				<td colspan="5">&nbsp;</td>
			</tr>	
			
			<%
	
			Do While Not rsWeeklySnapshotSummaryPipeline.EOF
			
				StageIntRecID = rsWeeklySnapshotSummaryPipeline("StageRecID")
				StageName = GetStageByNum (StageIntRecID)
				TotalNumberOfProspectByStage = rsWeeklySnapshotSummaryPipeline("Expr1")
				
				TotalProspectsByStageWeeklySnapshot = TotalProspectsByStageWeeklySnapshot + TotalNumberOfProspectByStage 	
				
				%>					
				
				<tr>
					<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfProspectByStage %></font>&nbsp;</td>
					<td width="42%"><font face="Consolas" style="font-size: 14pt"><%= StageName %></font></td>
					<td width="5%">&nbsp;</td>
					<td width="47%" colspan="2">&nbsp;</td>
				</tr>
			<% 
			
			rsWeeklySnapshotSummaryPipeline.MoveNext
			Loop
			
			RowCount = Rowcount + 15 %>
			
			
			<tr>
				<td colspan="2">
					<font face="Consolas" style="font-size: 14pt"><hr width="50%" align="left"></font>
				</td>
				<td width="5%">&nbsp;</td>
				<td width="47%" colspan="2">&nbsp;</td>
			</tr>
			
			<tr>
				<td width="5%" align="right"><font face="Consolas" style="font-size: 14pt"><%= TotalProspectsByStageWeeklySnapshot %></font>&nbsp;</td>
				<td width="42%"><font face="Consolas" style="font-size: 14pt">TOTAL</font></td>
				<td width="5%">&nbsp;</td>
				<td width="47%" colspan="2">&nbsp;</td>						
			</tr>
			
			<tr>
				<td colspan="5">&nbsp;</td>
			</tr>
			
			</table>
			
		<% End If %>

		<br/><br/>
		</table>
		
		</td>
	</tr>
	
	<% 

	Call Footer

		
	'*******************************************************
	'*** END END END END END END END END END END END END END 
	'*** This section is the first page which prints all the
	'*** prospecting weekly summary info
	'*******************************************************

	%>
	
	</table>
	
	<br><br><br>
	
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<tr>
				<td colspan="11">
		
<%

	'*******************************************************
	'*** This section is SUMMARY BY SALES REP
	'*******************************************************
	
	SQL = "SELECT * FROM PR_Prospects WHERE Pool = 'LIVE' AND OwnerUserNo IN (" & ProspSnapshotSalesRepDisplayUserNos  & ")"

	Set cnnWeeklySnapshotSummary = Server.CreateObject("ADODB.Connection")
	cnnWeeklySnapshotSummary.open(Session("ClientCnnString"))
	Set rsWeeklySnapshotSummary  = Server.CreateObject("ADODB.Recordset")
	rsWeeklySnapshotSummary.CursorLocation = 3 
	rsWeeklySnapshotSummary.Open SQL, cnnWeeklySnapshotSummary 
				
	'Response.Write(SQL & "<br>")
	
	If Not rsWeeklySnapshotSummary.EOF AND ProspSnapshotSalesRepDisplayUserNos <> "" Then
		
		Call PageHeader
	
		%>
		<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<tr>
				<td colspan="11">
				<font face="Consolas">
				<hr>
				<center><h2>Summary By Rep <%= mondayOfLastWeek %> - <%= sundayOfThisWeek %></h2></center>
				<hr>
				</font>
				</td>
			</tr>
		<%
		
		FontSizeVar = 9
		LinesPerPage = 42
		RowCount = 0
		
		RunningTotalProspectsPreexistingByRep = 0
		RunningTotalProspectsCreatedByRep = 0
		RunningTotalLostUnqualifiedProspectsByRep = 0
		RunningTotalWonProspectsByRep = 0
		RunningTotalAppmtsCreatedByRep = 0
		RunningTotalAppmtsCompletedByRep = 0
		RunningTotalAppmtsRescheduledByRep = 0
		RunningTotalAppmtsCancelledByRep = 0
		RunningTotalAppmtsNotUpdatedByRep = 0
		RunningTotalExpiredActivitiesByRep = 0
		
		Call SubHeaderSalesRep
		
		SalesRepDisplayUserNosArray = Split(ProspSnapshotSalesRepDisplayUserNos,",")
	
		For i = 0 to Ubound(SalesRepDisplayUserNosArray)

			SalesRepDisplayName = GetUserDisplayNameByUserNo(SalesRepDisplayUserNosArray(i))
			SalesRepUserNo = SalesRepDisplayUserNosArray(i)
			
			TotalNumberOfPreexistingProspectsByRep = TotalNumberOfPreexistingProspectsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)
			TotalNumberOfProspectsCreatedByRep = TotalNumberOfCreatedProspectsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)
			TotalNumberOfLostProspectsByRep = TotalNumberOfLostProspectsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)
			TotalNumberOfUnqualifiedProspectsByRep = TotalNumberOfUnqualifiedProspectsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)
			TotalNumberOfWonProspectsByRep = TotalNumberOfWonProspectsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)
			
			TotalNumberOfAppointmentsCreatedByRep = TotalNumberOfCreatedAppointmentsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)
			TotalNumberOfAppointmentsCompletedByRep = TotalNumberOfCompletedAppointmentsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)	
			TotalNumberOfAppointmentsRescheduledByRep = TotalNumberOfRescheduledAppointmentsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)
			TotalNumberOfAppointmentsCancelledByRep = TotalNumberOfCancelledAppointmentsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)
			TotalNumberOfAppointmentsNotUpdatedByRep = TotalNumberOfNotUpdatedAppointmentsWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)
			
			TotalNumberOfExpiredActivitiesByRep = TotalNumberOfExpiredActivitiesWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,SalesRepUserNo)
			
			RunningTotalProspectsPreexistingByRep = RunningTotalProspectsPreexistingByRep + TotalNumberOfPreexistingProspectsByRep
			RunningTotalProspectsCreatedByRep = RunningTotalProspectsCreatedByRep + TotalNumberOfProspectsCreatedByRep
			RunningTotalLostProspectsByRep = RunningTotalLostProspectsByRep + TotalNumberOfLostProspectsByRep
			RunningTotalUnqualifiedProspectsByRep = RunningTotalUnqualifiedProspectsByRep + TotalNumberOfUnqualifiedProspectsByRep
			RunningTotalWonProspectsByRep = RunningTotalWonProspectsByRep + TotalNumberOfWonProspectsByRep
			RunningTotalAppmtsCreatedByRep = RunningTotalAppmtsCreatedByRep + TotalNumberOfAppointmentsCreatedByRep
			RunningTotalAppmtsCompletedByRep = RunningTotalAppmtsCompletedByRep + TotalNumberOfAppointmentsCompletedByRep
			RunningTotalAppmtsRescheduledByRep = RunningTotalAppmtsRescheduledByRep + TotalNumberOfAppointmentsRescheduledByRep
			RunningTotalAppmtsCancelledByRep = RunningTotalAppmtsCancelledByRep + TotalNumberOfAppointmentsCancelledByRep
			RunningTotalAppmtsNotUpdatedByRep = RunningTotalAppmtsNotUpdatedByRep + TotalNumberOfAppointmentsNotUpdatedByRep
			RunningTotalExpiredActivitiesByRep = RunningTotalExpiredActivitiesByRep + TotalNumberOfExpiredActivitiesByRep

			%>
			<tr>
				<td width="15%"><font face="Consolas" style="font-size: 14pt"><%= SalesRepDisplayName %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfPreexistingProspectsByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfProspectsCreatedByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfLostProspectsByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfUnqualifiedProspectsByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfWonProspectsByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsCreatedByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsCompletedByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsRescheduledByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsCancelledByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfExpiredActivitiesByRep %></font></td>
			</tr>
			<%
			RowCount = RowCount + 1.5

			If RowCount > LinesPerPage Then
				%></table><%
				Call Footer
				%><br><br><br><%
				Call PageHeader
				Call SubHeaderSalesRep
			End If
			
		Next
	
		%>
			<tr>
				<td colspan="11"><hr></td>
			</tr>
			
			<tr>
				<td width="15%"><font face="Consolas" style="font-size: 14pt">TOTAL</font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalProspectsPreexistingByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalProspectsCreatedByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalLostProspectsByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalUnqualifiedProspectsByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalWonProspectsByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalAppmtsCreatedByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalAppmtsCompletedByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalAppmtsRescheduledByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalAppmtsCancelledByRep %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalExpiredActivitiesByRep %></font></td>
			</tr>
			<%
	
	Else
			
		Call PageHeader
	
		%>
		<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<tr>
				<td colspan="11">
				<font face="Consolas">
				<hr>
				<center><h2>Summary By Rep <%= mondayOfLastWeek %> - <%= sundayOfThisWeek %></h2></center>
				<hr>
				</font>
				</td>
			</tr>
		<%
		
		FontSizeVar = 9
		LinesPerPage = 42
		RowCount = 0
	%>
		<tr>
			<td colspan="11">
				<font face="Consolas">
					<center><h2>No Sales Rep Data for Last Week.</h2></center>
				</font>
			</td>
		</tr>
		
		</table>
	<%
	End If
	
	
	Call Footer


	'*******************************************************
	'*** END END END END END END END END END END END END END 
	'*** This section is SUMMARY BY SALES REP
	'*******************************************************
	
	%>
	
	</table>
	
	<br><br><br>
	
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<tr>
				<td colspan="10">
		
<%

	'*******************************************************
	'*** This section is SUMMARY BY LEAD SOURCE
	'*******************************************************

	SQLProspectingLeadSource = "SELECT PR_LeadSources.InternalRecordIdentifier AS LeadSourceNum FROM PR_LeadSources ORDER BY InternalRecordIdentifier "

	Set cnnWeeklySnapshotLeadSourceSummary = Server.CreateObject("ADODB.Connection")
	cnnWeeklySnapshotLeadSourceSummary.open(Session("ClientCnnString"))
	Set rsWeeklySnapshotLeadSourceSummary = Server.CreateObject("ADODB.Recordset")
	rsWeeklySnapshotLeadSourceSummary.CursorLocation = 3 
	rsWeeklySnapshotLeadSourceSummary.Open SQLProspectingLeadSource, cnnWeeklySnapshotLeadSourceSummary 
				
'	Response.Write(SQLProspectingLeadSource & "<br>")


	If Not rsWeeklySnapshotLeadSourceSummary.EOF Then
		
		Call PageHeader
	
		%>
		<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<tr>
				<td colspan="10">
				<font face="Consolas">
				<hr>
				<center><h2>Summary By Lead Source <%= mondayOfLastWeek %> - <%= sundayOfThisWeek %></h2></center>
				<hr>
				</font>
				</td>
			</tr>
		<%
		
		FontSizeVar = 9
		LinesPerPage = 42
		RowCount = 0
		
		RunningTotalProspectsCreatedByLeadSource = 0
		RunningTotalLostProspectsByLeadSource = 0
		RunningTotalUnqualifiedProspectsByLeadSource = 0
		RunningTotalWonProspectsByLeadSource = 0
		RunningTotalAppmtsCreatedByLeadSource = 0
		RunningTotalAppmtsCompletedByLeadSource = 0
		RunningTotalAppmtsRescheduledByLeadSource = 0
		RunningTotalAppmtsCancelledByLeadSource = 0
		RunningTotalAppmtsNotUpdatedByLeadSource = 0
		RunningTotalExpiredActivitiesByLeadSource = 0	
				
		Call SubHeader
		
		Do While NOT rsWeeklySnapshotLeadSourceSummary.EOF
		
			LeadIntRecID = rsWeeklySnapshotLeadSourceSummary("LeadSourceNum")
			LeadSourceDisplayName = GetLeadSourceByNum(LeadIntRecID)
			TotalForThisLeadSource = 0
			
			TotalNumberOfCreatedProspectsByLeadSource = TotalNumberOfCreatedProspectsByLeadSourceWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos,LeadIntRecID)
			TotalNumberOfLostProspectsByLeadSource = TotalNumberOfLostProspectsByLeadSourceWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos,LeadIntRecID)
			TotalNumberOfUnqualifiedProspectsByLeadSource = TotalNumberOfUnqualifiedProspectsByLeadSourceWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos,LeadIntRecID)
			TotalNumberOfWonProspectsByLeadSource = TotalNumberOfWonProspectsByLeadSourceWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos,LeadIntRecID)
			
			TotalNumberOfAppointmentsCreatedByLeadSource = TotalNumberOfCreatedAppointmentsByLeadSourceWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos,LeadIntRecID)
			TotalNumberOfAppointmentsCompletedByLeadSource = TotalNumberOfCompletedAppointmentsByLeadSourceWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos,LeadIntRecID)	
			TotalNumberOfAppointmentsRescheduledByLeadSource = TotalNumberOfRescheduledAppointmentsByLeadSourceWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos,LeadIntRecID)
			TotalNumberOfAppointmentsCancelledByLeadSource = TotalNumberOfCancelledAppointmentsByLeadSourceWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos,LeadIntRecID)
			TotalNumberOfAppointmentsNotUpdatedByLeadSource = TotalNumberOfNotUpdatedAppointmentsByLeadSourceWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos,LeadIntRecID)
			
			TotalNumberOfExpiredActivitiesByLeadSource = TotalNumberOfExpiredActivitiesByLeadSourceWeeklySnapshot(mondayOfLastWeek,sundayOfThisWeek,ProspSnapshotSalesRepDisplayUserNos,LeadIntRecID)

			RunningTotalProspectsCreatedByLeadSource = RunningTotalProspectsCreatedByLeadSource + TotalNumberOfCreatedProspectsByLeadSource
			RunningTotalLostProspectsByLeadSource = RunningTotalLostProspectsByLeadSource + TotalNumberOfLostProspectsByLeadSource 
			RunningTotalUnqualifiedProspectsByLeadSource = RunningTotalUnqualifiedProspectsByLeadSource + TotalNumberOfUnqualifiedProspectsByLeadSource
			RunningTotalWonProspectsByLeadSource = RunningTotalWonProspectsByLeadSource + TotalNumberOfWonProspectsByLeadSource
			RunningTotalAppmtsCreatedByLeadSource = RunningTotalAppmtsCreatedByLeadSource + TotalNumberOfAppointmentsCreatedByLeadSource
			RunningTotalAppmtsCompletedByLeadSource = RunningTotalAppmtsCompletedByLeadSource + TotalNumberOfAppointmentsCompletedByLeadSource
			RunningTotalAppmtsRescheduledByLeadSource = RunningTotalAppmtsRescheduledByLeadSource + TotalNumberOfAppointmentsRescheduledByLeadSource
			RunningTotalAppmtsCancelledByLeadSource = RunningTotalAppmtsCancelledByLeadSource + TotalNumberOfAppointmentsCancelledByLeadSource
			RunningTotalAppmtsNotUpdatedByLeadSource = RunningTotalAppmtsNotUpdatedByLeadSource + TotalNumberOfAppointmentsNotUpdatedByLeadSource
			RunningTotalExpiredActivitiesByLeadSource = RunningTotalExpiredActivitiesByLeadSource + TotalNumberOfExpiredActivitiesByLeadSource
			
			
			TotalForThisLeadSource = TotalNumberOfCreatedProspectsByLeadSource + TotalNumberOfLostUnqualifiedProspectsByLeadSource + TotalNumberOfWonProspectsByLeadSource + _
				TotalNumberOfAppointmentsCreatedByLeadSource + TotalNumberOfAppointmentsCompletedByLeadSource + TotalNumberOfAppointmentsRescheduledByLeadSource + _
				TotalNumberOfAppointmentsCancelledByLeadSource + TotalNumberOfExpiredActivitiesByLeadSource
				
			If TotalForThisLeadSource > 0 Then
				%>
				<tr>
					<td width="15%"><font face="Consolas" style="font-size: 14pt"><%= LeadSourceDisplayName %></font></td>
					<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfCreatedProspectsByLeadSource %></font></td>
					<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfLostProspectsByLeadSource %></font></td>
					<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfUnqualifiedProspectsByLeadSource %></font></td>
					<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfWonProspectsByLeadSource %></font></td>
					<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsCreatedByLeadSource %></font></td>
					<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsCompletedByLeadSource %></font></td>
					<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsRescheduledByLeadSource %></font></td>
					<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfAppointmentsCancelledByLeadSource %></font></td>
					<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= TotalNumberOfExpiredActivitiesByLeadSource %></font></td>
				</tr>
				<%
				RowCount = RowCount + 1.5
			End If
			

			If RowCount > LinesPerPage Then
				%></table><%
				Call Footer
				%><br><br><br><%
				Call PageHeader
				Call SubHeader
			End If
			
		rsWeeklySnapshotLeadSourceSummary.MoveNext

		Loop
		
		%>
			<tr>
				<td colspan="10"><hr></td>
			</tr>
			
			<tr>
				<td width="15%"><font face="Consolas" style="font-size: 14pt">TOTAL</font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalProspectsCreatedByLeadSource %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalLostProspectsByLeadSource %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalUnqualifiedProspectsByLeadSource %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalWonProspectsByLeadSource %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalAppmtsCreatedByLeadSource %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalAppmtsCompletedByLeadSource %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalAppmtsRescheduledByLeadSource %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalAppmtsCancelledByLeadSource %></font></td>
				<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt"><%= RunningTotalExpiredActivitiesByLeadSource %></font></td>
			</tr>
			<%
	
	Else
		
		Call PageHeader
	
		%>
		<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
			<tr>
				<td colspan="10">
				<font face="Consolas">
				<hr>
				<center><h2>Summary By Lead Source <%= mondayOfLastWeek %> - <%= sundayOfThisWeek %></h2></center>
				<hr>
				</font>
				</td>
			</tr>
		<%
		
		FontSizeVar = 9
		LinesPerPage = 42
		RowCount = 0
			
	%>
		<tr>
			<td colspan="10">
				<font face="Consolas">
					<center><h2>No Lead Source Data for Last Week.</h2></center>
				</font>
			</td>
		</tr>
		
		</table>
	<%
	End If
	
	
	Call Footer


	'*******************************************************
	'*** END END END END END END END END END END END END END 
	'*** This section is SUMMARY BY LEAD SOURCE
	'*******************************************************
	
	%>
	
	</table>
	
</body>
</html>


<%
Sub PageHeader

	RowCount = 0
	%>

	<table border="0" width="100%">
		<tr>
			<td width="60%"><img src="<%= BaseURL %>clientfiles/<%= MUV_Read("ClientID") %>/logos/logo.png" style="height:55px;"></td>
			<td width="40%">
				<p align="center"><b><font face="Consolas" size="4">MDS Insight Weekly Prospecting Snapshot Report</font></b></p>
				<p align="center"><font face="Consolas" size="2">Report Generated: <%= WeekDayName(WeekDay(DateValue(Now()))) %>&nbsp;<%= Now() %><br></font></p>
			</td>
		</tr>
		<tr>
			<td width="20%" height="16">
				<p align="right"><font face="Consolas" size="1">&nbsp;</font></p>
			</td>
		</tr>
	</table>
	<%
	PageNum = PageNum + 1
End Sub

Sub SubHeader
	%> 
		<tr>
			<td colspan="10">&nbsp;</td>
		</tr>
		<tr>
			<td width="15%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Prosp.<br>Created</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Prosp.<br>Lost</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Prosp.<br>Unqualified</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Prosp.<br>Won</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Appt.<br>Created</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Appt.<br>Completed</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Appt.<br>Rescheduled</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Appt.<br>Cancelled</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Activity<br>Expired</font></td>
		</tr>
		<tr>
			<td colspan="10"><hr></td>
		</tr>
		<tr>
			<td colspan="10">&nbsp;</td>
		</tr>
	<%
End Sub


Sub SubHeaderSalesRep
	%> 
		<tr>
			<td colspan="11">&nbsp;</td>
		</tr>
		<tr>
			<td width="15%"><font face="Consolas" style="font-size: 14pt">&nbsp;</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Pre-<br>Existing</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Prosp.<br>Created</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Prosp.<br>Lost</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Prosp.<br>Unqualified</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Prosp.<br>Won</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Appt.<br>Created</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Appt.<br>Completed</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Appt.<br>Rescheduled</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Appt.<br>Cancelled</font></td>
			<td width="9%" align="center"><font face="Consolas" style="font-size: 14pt">Activity<br>Expired</font></td>
		</tr>
		<tr>
			<td colspan="11"><hr></td>
		</tr>
		<tr>
			<td colspan="11">&nbsp;</td>
		</tr>
	<%
End Sub

Sub Footer

	'Now get us to the next page
	Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'><table>")
	For x = 1 to LinesPerPage - RowCount
		Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'><tr><td border='1'>&nbsp;</td></tr>")
	Next
	Response.Write("<font face='Consolas' style='font-size: " & FontSizeVar & "pt'></table>")
	%>
	<table border="0" width="<%=PageWidth%>" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table32" align="center">	
		<tr>
			<td colspan="3">
				<hr>
			</td>
		</tr>
		<tr>
			<td width="33%">
				<font face="Consolas" style="font-size: 9pt">directlaunch/prospecting/prospectingSnapshotReport.asp</font>
			</td>
			<td width="33%" align="center">
				<font face="Consolas" style="font-size: 12pt">Page:&nbsp;<%=PageNum%></font>
			</td>
			<td width="33%">
				<font face="Consolas" style="font-size: 12pt">&nbsp;</font>
			</td>
		</tr>
	</table>
	<% If NoBreak <> True Then %>
		<BR style="page-break-after: always">
	<% End If

End Sub

%>