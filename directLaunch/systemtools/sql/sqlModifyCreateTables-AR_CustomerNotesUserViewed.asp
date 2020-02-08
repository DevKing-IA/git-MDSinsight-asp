<%	
	Set cnnCheckARCustNotesUserView = Server.CreateObject("ADODB.Connection")
	cnnCheckARCustNotesUserView.open (Session("ClientCnnString"))
	Set rsCheckARCustNotesUserView = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckARCustNotesUserView = cnnCheckARCustNotesUserView.Execute("SELECT TOP 1 * FROM AR_CustomerNotesUserViewed")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckARCustNotesUserView = "CREATE TABLE [AR_CustomerNotesUserViewed]("
			SQLCheckARCustNotesUserView = SQLCheckARCustNotesUserView & "	[UserNo] [int] NULL, "
			SQLCheckARCustNotesUserView = SQLCheckARCustNotesUserView & "	[CustID] [varchar](50) NULL, "
			SQLCheckARCustNotesUserView = SQLCheckARCustNotesUserView & "	[Category] [int] NULL, "
			SQLCheckARCustNotesUserView = SQLCheckARCustNotesUserView & "	[DateLastViewed] [datetime] NULL "
			SQLCheckARCustNotesUserView = SQLCheckARCustNotesUserView & " ) ON [PRIMARY]"      

		   	Set rsCheckARCustNotesUserView = cnnCheckARCustNotesUserView.Execute(SQLCheckARCustNotesUserView)
		   
		End If
	End If


	SQLCheckARCustNotesUserView  = "SELECT COL_LENGTH('AR_CustomerNotesUserViewed', 'NoteTypeIntRecID') AS IsItThere"
	Set rsCheckARCustNotesUserView = cnnCheckARCustNotesUserView.Execute(SQLCheckARCustNotesUserView )
	If IsNull(rsCheckARCustNotesUserView ("IsItThere")) Then
		SQLCheckARCustNotesUserView = "ALTER TABLE AR_CustomerNotesUserViewed ADD NoteTypeIntRecID  int NULL"
		Set rsCheckARCustNotesUserView = cnnCheckARCustNotesUserView.Execute(SQLCheckARCustNotesUserView )
	End If

		
	SQLCheckARCustNotesUserView = "UPDATE AR_CustomerNotesUserViewed SET NoteTypeIntRecID = 0 WHERE NoteTypeIntRecID IS NULL"
	Set rsCheckARCustNotesUserView = cnnCheckARCustNotesUserView.Execute(SQLCheckARCustNotesUserView )
		
	set rsCheckARCustNotesUserView = nothing
	cnnCheckARCustNotesUserView.close
	set cnnCheckARCustNotesUserView = nothing
				
%>