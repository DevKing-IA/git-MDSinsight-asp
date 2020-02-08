<%	
	Response.Write("sqlModifyCreateTables-AR_CustomerContacts.asp" & "<br>")
	
	Set cnnCheckCustomerContacts = Server.CreateObject("ADODB.Connection")
	cnnCheckCustomerContacts.open (Session("ClientCnnString"))
	Set rsCheckCustomerContacts = Server.CreateObject("ADODB.Recordset")
	rsCheckCustomerContacts.CursorLocation = 3 

	Err.Clear
	on error resume next
	Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute("SELECT TOP 1 * FROM AR_CustomerContacts")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			SQLCheckCustomerContacts = "CREATE TABLE [AR_CustomerContacts]( "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & " [InternalRecordIdentifier] [int] IDENTITY(2,1) NOT NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[RecordCreationDateTime] [datetime] NULL, "			
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[CustomerIntRecID] [int] NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[Suffix] [varchar](50) NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[FirstName] [varchar](255) NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[LastName] [varchar](255) NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[ContactTitleNumber] [int] NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[Email] [varchar](255) NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[Phone] [varchar](255) NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[PhoneExt] [varchar](255) NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[Cell] [varchar](255) NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[Fax] [varchar](255) NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[DecisionMaker] [bit] NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[PrimaryContact] [bit] NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[Notes] [varchar](8000) NULL, "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & "[DoNotEmail] [bit] NULL "
			SQLCheckCustomerContacts = SQLCheckCustomerContacts & ") ON [PRIMARY]"
			
			Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)

		
		End If
	End If
	
	On Error Goto 0
	
	Set cnnCheckCustomerContacts = Server.CreateObject("ADODB.Connection")
	cnnCheckCustomerContacts.open (Session("ClientCnnString"))
	Set rsCheckCustomerContacts = Server.CreateObject("ADODB.Recordset")

		
' This one is a DROP
	SQLCheckCustomerContacts  = "SELECT COL_LENGTH('AR_CustomerContacts', 'Address1') AS IsItThere"
	Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	If NOT IsNull(rsCheckCustomerContacts("IsItThere")) Then
		SQLCheckCustomerContacts = "ALTER TABLE AR_CustomerContacts DROP COLUMN Address1"
		Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	End If
		
' This one is a DROP
	SQLCheckCustomerContacts  = "SELECT COL_LENGTH('AR_CustomerContacts', 'Latitude') AS IsItThere"
	Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	If NOT IsNull(rsCheckCustomerContacts("IsItThere")) Then
		SQLCheckCustomerContacts = "ALTER TABLE AR_CustomerContacts DROP COLUMN Latitude"
		Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	End If
		
' This one is a DROP
	SQLCheckCustomerContacts  = "SELECT COL_LENGTH('AR_CustomerContacts', 'Longitude') AS IsItThere"
	Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	If NOT IsNull(rsCheckCustomerContacts("IsItThere")) Then
		SQLCheckCustomerContacts = "ALTER TABLE AR_CustomerContacts DROP COLUMN Longitude"
		Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	End If
		
' This one is a DROP
	SQLCheckCustomerContacts  = "SELECT COL_LENGTH('AR_CustomerContacts', 'Country') AS IsItThere"
	Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	If NOT IsNull(rsCheckCustomerContacts("IsItThere")) Then
		SQLCheckCustomerContacts = "ALTER TABLE AR_CustomerContacts DROP COLUMN Country"
		Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	End If
		
' This one is a DROP
	SQLCheckCustomerContacts  = "SELECT COL_LENGTH('AR_CustomerContacts', 'PostalCode') AS IsItThere"
	Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	If NOT IsNull(rsCheckCustomerContacts("IsItThere")) Then
		SQLCheckCustomerContacts = "ALTER TABLE AR_CustomerContacts DROP COLUMN PostalCode"
		Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	End If
		
' This one is a DROP
	SQLCheckCustomerContacts  = "SELECT COL_LENGTH('AR_CustomerContacts', 'State') AS IsItThere"
	Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	If NOT IsNull(rsCheckCustomerContacts("IsItThere")) Then
		SQLCheckCustomerContacts = "ALTER TABLE AR_CustomerContacts DROP COLUMN State"
		Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	End If
		
' This one is a DROP
	SQLCheckCustomerContacts  = "SELECT COL_LENGTH('AR_CustomerContacts', 'City') AS IsItThere"
	Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	If NOT IsNull(rsCheckCustomerContacts("IsItThere")) Then
		SQLCheckCustomerContacts = "ALTER TABLE AR_CustomerContacts DROP COLUMN City"
		Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	End If
		
' This one is a DROP
	SQLCheckCustomerContacts  = "SELECT COL_LENGTH('AR_CustomerContacts', 'Address2') AS IsItThere"
	Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	If NOT IsNull(rsCheckCustomerContacts("IsItThere")) Then
		SQLCheckCustomerContacts = "ALTER TABLE AR_CustomerContacts DROP COLUMN Address2"
		Set rsCheckCustomerContacts = cnnCheckCustomerContacts.Execute(SQLCheckCustomerContacts)
	End If

	set rsCheckCustomerContacts = nothing
	cnnCheckCustomerContacts.close
	set cnnCheckCustomerContacts = nothing
%>