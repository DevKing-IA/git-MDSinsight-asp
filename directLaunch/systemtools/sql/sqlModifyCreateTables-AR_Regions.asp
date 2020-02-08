<%	
	Response.Write("sqlModifyCreateTables-AR_Regions.asp" & "<br>")
	On Error Goto 0

	Set cnnCheckARRegions = Server.CreateObject("ADODB.Connection")
	cnnCheckARRegions.open (Session("ClientCnnString"))
	Set rsCheckARRegions = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckARRegions = cnnCheckARRegions.Execute("SELECT TOP 1 * FROM AR_Regions")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckARRegions = "CREATE TABLE [AR_Regions]("
			SQLCheckARRegions = SQLCheckARRegions & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckARRegions = SQLCheckARRegions & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_Regions_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckARRegions = SQLCheckARRegions & " [Region] [varchar](255) NULL,"
			SQLCheckARRegions = SQLCheckARRegions & " [Cities1] [varchar](8000) NULL,"
			SQLCheckARRegions = SQLCheckARRegions & " [Cities2] [varchar](8000) NULL,"
			SQLCheckARRegions = SQLCheckARRegions & " [Cities3] [varchar](8000) NULL,"
			SQLCheckARRegions = SQLCheckARRegions & " [StatesOrProvinces] [varchar](8000) NULL,"
			SQLCheckARRegions = SQLCheckARRegions & " [ZipOrPostalCodes1] [varchar](8000) NULL,"
			SQLCheckARRegions = SQLCheckARRegions & " [ZipOrPostalCodes2] [varchar](8000) NULL,"
			SQLCheckARRegions = SQLCheckARRegions & " [AutoFilterPercentage] [float] NULL CONSTRAINT [DF_AR_Regions_AutoFilterPercentage]  DEFAULT (0) "
			SQLCheckARRegions = SQLCheckARRegions & " ) ON [PRIMARY]"      

		   Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
		   
		End If
	End If

	On error goto 0
	
	'Special for the parts  file
	'Make sure code 0 is there
	SQLCheckARRegions = "SELECT * FROM AR_Regions WHERE InternalRecordIdentifier = 0"
	Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	If rsCheckARRegions.EOF Then 
	
		SQLCheckARRegions = "SET IDENTITY_INSERT AR_Regions ON;"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)

		SQLCheckARRegions = SQLCheckARRegions & "INSERT INTO AR_Regions (InternalRecordIdentifier,Region) "
		SQLCheckARRegions = SQLCheckARRegions & " VALUES (0,'Undefined')"
		Response.Write(SQLCheckARRegions)
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
		
		SQLCheckARRegions = "SET IDENTITY_INSERT AR_Regions OFF;"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
		
	End If

	SQLCheckARRegions = "SELECT COL_LENGTH('AR_Regions', 'IncludeInAutoFilterTickets') AS IsItThere"
	Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	If IsNull(rsCheckARRegions("IsItThere")) Then
		SQLCheckARRegions = "ALTER TABLE AR_Regions ADD IncludeInAutoFilterTickets int NULL"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	End If
	
	SQLCheckARRegions = "SELECT COL_LENGTH('AR_Regions', 'IncludeInAutoFilterTickets') AS IsItThere"
	Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	If IsNull(rsCheckARRegions("IsItThere")) Then
		SQLCheckARRegions = "ALTER TABLE AR_Regions ADD IncludeInAutoFilterTickets int NULL"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	End If
	
	SQLCheckARRegions = "SELECT COL_LENGTH('AR_Regions', 'IncludeInSuggestedFilterTickets') AS IsItThere"
	Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	If IsNull(rsCheckARRegions("IsItThere")) Then
		SQLCheckARRegions = "ALTER TABLE AR_Regions ADD IncludeInSuggestedFilterTickets int NULL"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	End If
	
	SQLCheckARRegions = "SELECT COL_LENGTH('AR_Regions', 'AutoFilterChangeMaxNumTicketsPerDay') AS IsItThere"
	Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	If IsNull(rsCheckARRegions("IsItThere")) Then
		SQLCheckARRegions = "ALTER TABLE AR_Regions ADD AutoFilterChangeMaxNumTicketsPerDay int NULL"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
		SQLCheckARRegions = "UPDATE AR_Regions SET AutoFilterChangeMaxNumTicketsPerDay = 25"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	End If
	
	SQLCheckARRegions = "SELECT COL_LENGTH('AR_Regions', 'SuggestedFilterChangeMaxNumTicketsPerDay') AS IsItThere"
	Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	If IsNull(rsCheckARRegions("IsItThere")) Then
		SQLCheckARRegions = "ALTER TABLE AR_Regions ADD SuggestedFilterChangeMaxNumTicketsPerDay int NULL"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
		SQLCheckARRegions = "UPDATE AR_Regions SET SuggestedFilterChangeMaxNumTicketsPerDay = 25"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	End If

	' This one is a drop
	SQLCheckARRegions = "SELECT COL_LENGTH('AR_Regions', 'AutoFilterPercentage') AS IsItThere"
	Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	If NOT IsNull(rsCheckARRegions("IsItThere")) Then
		SQLCheckARRegions = "ALTER TABLE AR_Regions DROP COLUMN AutoFilterPercentage"
'		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	End If
		
	SQLCheckARRegions = "SELECT COL_LENGTH('AR_Regions', 'StateForCities') AS IsItThere"
	Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	If IsNull(rsCheckARRegions("IsItThere")) Then
		SQLCheckARRegions = "ALTER TABLE AR_Regions ADD StateForCities varchar(255) NULL"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	End If

	SQLCheckARRegions = "SELECT COL_LENGTH('AR_Regions', 'CatchAllRegionIntRecIDs') AS IsItThere"
	Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	If IsNull(rsCheckARRegions("IsItThere")) Then
		SQLCheckARRegions = "ALTER TABLE AR_Regions ADD CatchAllRegionIntRecIDs varchar(255) NULL"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	End If

	SQLCheckARRegions = "SELECT COL_LENGTH('AR_Regions', 'UseRegionForServiceTickets') AS IsItThere"
	Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	If IsNull(rsCheckARRegions("IsItThere")) Then
		SQLCheckARRegions = "ALTER TABLE AR_Regions ADD UseRegionForServiceTickets int NULL"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
		SQLCheckARRegions = "UPDATE AR_Regions SET UseRegionForServiceTickets = 0"
		Set rsCheckARRegions = cnnCheckARRegions.Execute(SQLCheckARRegions)
	End If

	set rsCheckARRegions = nothing
	cnnCheckARRegions.close
	set cnnCheckARRegions = nothing
				
%>