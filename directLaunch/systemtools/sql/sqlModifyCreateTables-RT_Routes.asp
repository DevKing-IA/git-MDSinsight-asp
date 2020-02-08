<%	
	Set cnnRT_Routes = Server.CreateObject("ADODB.Connection")
	cnnRT_Routes.open (Session("ClientCnnString"))
	Set rsRT_Routes = Server.CreateObject("ADODB.Recordset")
	rsRT_Routes.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsRT_Routes = cnnRT_Routes.Execute("SELECT TOP 1 * FROM RT_Routes ORDER BY InternalRecordIdentifier DESC")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLRT_Routes = "CREATE TABLE [RT_Routes]( "
			SQLRT_Routes = SQLRT_Routes & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL, "
			SQLRT_Routes = SQLRT_Routes & " [RecordCreationDateTime] [datetime] NULL, "
			SQLRT_Routes = SQLRT_Routes & " [RecordSource] [varchar](255) NULL, "
			SQLRT_Routes = SQLRT_Routes & " [RouteID] [varchar](255) NULL, "
			SQLRT_Routes = SQLRT_Routes & " [RouteDescription] [varchar](255) NULL, "
			SQLRT_Routes = SQLRT_Routes & " [ShowOnDBoard] [int] NULL, "
			SQLRT_Routes = SQLRT_Routes & " [ShowInWebApp] [int] NULL, "
			SQLRT_Routes = SQLRT_Routes & " [ShowInPlanner] [int] NULL, "
			SQLRT_Routes = SQLRT_Routes & " [ThirdPartyCarrier] [int] NULL, "
			SQLRT_Routes = SQLRT_Routes & " [DefaultDriverUserNo] [int] NULL "
			SQLRT_Routes = SQLRT_Routes & ") ON [PRIMARY]"
			Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
		End If
		
		
		SQLRT_Routes = "CREATE CLUSTERED INDEX [IX_RT_Routes] ON [RT_Routes] ("
		SQLRT_Routes = SQLRT_Routes & " [RouteID] ASC "
		SQLRT_Routes = SQLRT_Routes & " )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, "
		SQLRT_Routes = SQLRT_Routes & " SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"

		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)

	
		SQLRT_Routes = "ALTER TABLE [RT_Routes] ADD  CONSTRAINT [DF_RT_Routes_RecordCreationDateTime]  DEFAULT (getdate()) FOR [RecordCreationDateTime]"

		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)

	
		SQLRT_Routes = "ALTER TABLE [RT_Routes] ADD  CONSTRAINT [DF_RT_Routes_RecordSource]  DEFAULT ('Insight') FOR [RecordSource]"

		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)

	
		SQLRT_Routes = "ALTER TABLE [RT_Routes] ADD  CONSTRAINT [DF_RT_Routes_ShowOnDBoard]  DEFAULT ((1)) FOR [ShowOnDBoard]"

		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)

	
		SQLRT_Routes = "ALTER TABLE [RT_Routes] ADD  CONSTRAINT [DF_RT_Routes_ShowInWebApp]  DEFAULT ((1)) FOR [ShowInWebApp]"

		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)

	
		SQLRT_Routes = "ALTER TABLE [RT_Routes] ADD  CONSTRAINT [DF_RT_Routes_ShowInPlanner]  DEFAULT ((1)) FOR [ShowInPlanner]"

		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)

	
		SQLRT_Routes = "ALTER TABLE [RT_Routes] ADD  CONSTRAINT [DF_RT_Routes_ThirdPartyCarrier]  DEFAULT ((0)) FOR [ThirdPartyCarrier]"

		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)

	
		SQLRT_Routes = "ALTER TABLE [RT_Routes] ADD  CONSTRAINT [DF_RT_Routes_DefaultDriverUserNo]  DEFAULT ((0)) FOR [DefaultDriverUserNo]"

		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
		
	End If
	
	SQLRT_Routes = "SELECT COL_LENGTH('RT_Routes', 'Monday') AS IsItThere"
	Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	If IsNull(rsRT_Routes("IsItThere")) Then
		SQLRT_Routes = "ALTER TABLE RT_Routes ADD Monday int NULL"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
		SQLRT_Routes = "UPDATE RT_Routes SET Monday = 1"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	End If
	
	SQLRT_Routes = "SELECT COL_LENGTH('RT_Routes', 'Tuesday') AS IsItThere"
	Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	If IsNull(rsRT_Routes("IsItThere")) Then
		SQLRT_Routes = "ALTER TABLE RT_Routes ADD Tuesday int NULL"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
		SQLRT_Routes = "UPDATE RT_Routes SET Tuesday = 1"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	End If
	
	SQLRT_Routes = "SELECT COL_LENGTH('RT_Routes', 'Wednesday') AS IsItThere"
	Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	If IsNull(rsRT_Routes("IsItThere")) Then
		SQLRT_Routes = "ALTER TABLE RT_Routes ADD Wednesday int NULL"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
		SQLRT_Routes = "UPDATE RT_Routes SET Wednesday = 1"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	End If

	SQLRT_Routes = "SELECT COL_LENGTH('RT_Routes', 'Thursday') AS IsItThere"
	Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	If IsNull(rsRT_Routes("IsItThere")) Then
		SQLRT_Routes = "ALTER TABLE RT_Routes ADD Thursdayint NULL"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
		SQLRT_Routes = "UPDATE RT_Routes SET Thursday= 1"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	End If

	SQLRT_Routes = "SELECT COL_LENGTH('RT_Routes', 'Thrusday') AS IsItThere"
	Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	If NOT IsNull(rsRT_Routes("IsItThere")) Then
		SQLRT_Routes = "ALTER TABLE RT_Routes DROP COLUMN Thrusday"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	End If
	
	SQLRT_Routes = "SELECT COL_LENGTH('RT_Routes', 'Friday') AS IsItThere"
	Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	If IsNull(rsRT_Routes("IsItThere")) Then
		SQLRT_Routes = "ALTER TABLE RT_Routes ADD Friday int NULL"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
		SQLRT_Routes = "UPDATE RT_Routes SET Friday = 1"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	End If
	
	SQLRT_Routes = "SELECT COL_LENGTH('RT_Routes', 'Saturday') AS IsItThere"
	Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	If IsNull(rsRT_Routes("IsItThere")) Then
		SQLRT_Routes = "ALTER TABLE RT_Routes ADD Saturday int NULL"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
		SQLRT_Routes = "UPDATE RT_Routes SET Saturday = 0"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	End If

	SQLRT_Routes = "SELECT COL_LENGTH('RT_Routes', 'Sunday') AS IsItThere"
	Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	If IsNull(rsRT_Routes("IsItThere")) Then
		SQLRT_Routes = "ALTER TABLE RT_Routes ADD Sunday int NULL"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
		SQLRT_Routes = "UPDATE RT_Routes SET Sunday = 0"
		Set rsRT_Routes = cnnRT_Routes.Execute(SQLRT_Routes)
	End If

				
	set rsRT_Routes = nothing
	cnnRT_Routes.close
	set cnnRT_Routes = nothing
				
%>