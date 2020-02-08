<%

	Set cnnEquipment = Server.CreateObject("ADODB.Connection")
	cnnEquipment.open (Session("ClientCnnString"))
	Set rsEquipment = Server.CreateObject("ADODB.Recordset")
	rsEquipment.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsEquipment = cnnEquipment.Execute("SELECT * FROM EQ_MovementCodes")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE EQ_MovementCodes ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_EQ_MovementCodes_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "movementCode [varchar](255) NULL, "
			SQLBuild = SQLBuild & "movementDesc [varchar](255) NULL, "
			SQLBuild = SQLBuild & "RecordSource [varchar](50) NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			Set rsEquipment = cnnEquipment.Execute(SQLBuild)
		End If
	End If
	On Error Goto 0

	Err.Clear
	on error resume next
	Set rsEquipment = cnnEquipment.Execute("SELECT * FROM EQ_AcquisitionCodes")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0
			'The table is not there, we need to create it
			SQLBuild = "CREATE TABLE EQ_AcquisitionCodes ("
			SQLBuild = SQLBuild & "InternalRecordIdentifier [int] IDENTITY(1,1) NOT NULL, "
			SQLBuild = SQLBuild & "RecordCreationDateTime [datetime] NULL CONSTRAINT [DF_EQ_AcquisitionCodes_RecordCreationDateTime] DEFAULT (getdate()), "
			SQLBuild = SQLBuild & "acquisitionCode [varchar](255) NULL, "
			SQLBuild = SQLBuild & "acquisitionDesc [varchar](255) NULL, "
			SQLBuild = SQLBuild & "RecordSource [varchar](50) NULL "
			SQLBuild = SQLBuild & ") ON [PRIMARY]"
			Set rsEquipment = cnnEquipment.Execute(SQLBuild)
		End If
	End If
	On Error Goto 0
	
	Set cnnCheckEquipment = Server.CreateObject("ADODB.Connection")
	cnnCheckEquipment.open (Session("ClientCnnString"))
	Set rsCheckEquipment = Server.CreateObject("ADODB.Recordset")
	rsCheckEquipment.CursorLocation = 3 


	SQL_CheckEquipment = "SELECT COL_LENGTH('EQ_Equipment', 'MovementCodeIntRecID') AS IsItThere"
	Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	If IsNull(rsCheckEquipment("IsItThere")) Then
		SQL_CheckEquipment = "ALTER TABLE EQ_Equipment ADD MovementCodeIntRecID int NULL"
		Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	End If


	SQL_CheckEquipment = "SELECT COL_LENGTH('EQ_Equipment', 'AcquisitionCodeIntRecID') AS IsItThere"
	Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	If IsNull(rsCheckEquipment("IsItThere")) Then
		SQL_CheckEquipment = "ALTER TABLE EQ_Equipment ADD AcquisitionCodeIntRecID int NULL"
		Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	End If


	SQL_CheckEquipment = "SELECT COL_LENGTH('EQ_Equipment', 'PurchaseOrAquisitionType') AS IsItThere"
	Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	If NOT IsNull(rsCheckEquipment("IsItThere")) Then
		SQL_CheckEquipment = "ALTER TABLE EQ_Equipment DROP COLUMN PurchaseOrAquisitionType"
		Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	End If

	
	SQL_CheckEquipment = "SELECT COL_LENGTH('EQ_AcquisitionCodes', 'RecordSource') AS IsItThere"
	Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	If IsNull(rsCheckEquipment("IsItThere")) Then
		SQL_CheckEquipment = "ALTER TABLE EQ_AcquisitionCodes ADD RecordSource varchar(50) NULL"
		Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	End If
	
	
	SQL_CheckEquipment = "SELECT COL_LENGTH('EQ_MovementCodes', 'RecordSource') AS IsItThere"
	Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	If IsNull(rsCheckEquipment("IsItThere")) Then
		SQL_CheckEquipment = "ALTER TABLE EQ_MovementCodes ADD RecordSource varchar(50) NULL"
		Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	End If
	
	
	SQL_CheckEquipment = "SELECT COL_LENGTH('EQ_Models', 'InsightAssetTagPrefix') AS IsItThere"
	Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	If IsNull(rsCheckEquipment("IsItThere")) Then
		SQL_CheckEquipment = "ALTER TABLE EQ_Models ADD InsightAssetTagPrefix varchar(50) NULL"
		Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	End If

	
	SQL_CheckEquipment = "SELECT COL_LENGTH('EQ_Brands', 'InsightAssetTagPrefix') AS IsItThere"
	Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	If IsNull(rsCheckEquipment("IsItThere")) Then
		SQL_CheckEquipment = "ALTER TABLE EQ_Brands ADD InsightAssetTagPrefix varchar(50) NULL"
		Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	End If
	
	
	SQL_CheckEquipment = "SELECT COL_LENGTH('EQ_Classes', 'InsightAssetTagPrefix') AS IsItThere"
	Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	If IsNull(rsCheckEquipment("IsItThere")) Then
		SQL_CheckEquipment = "ALTER TABLE EQ_Classes ADD InsightAssetTagPrefix varchar(50) NULL"
		Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	End If
	
	
	SQL_CheckEquipment = "SELECT COL_LENGTH('EQ_Manufacturers', 'InsightAssetTagPrefix') AS IsItThere"
	Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	If IsNull(rsCheckEquipment("IsItThere")) Then
		SQL_CheckEquipment = "ALTER TABLE EQ_Manufacturers ADD InsightAssetTagPrefix varchar(50) NULL"
		Set rsCheckEquipment = cnnCheckEquipment.Execute(SQL_CheckEquipment)
	End If


	set rsEquipment = nothing
	cnnEquipment.close
	set cnnEquipment = nothing
	

	Set rsCheckEquipment = Nothing
	cnnCheckEquipment.Close
	Set cnnCheckEquipment = Nothing
				
	
			
%>