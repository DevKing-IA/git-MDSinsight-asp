<%	
	Set cnnCheckARCustPOS = Server.CreateObject("ADODB.Connection")
	cnnCheckARCustPOS.open (Session("ClientCnnString"))
	Set rsCheckARCustPOS = Server.CreateObject("ADODB.Recordset")

	Err.Clear
	on error resume next
	Set rsCheckARCustPOS = cnnCheckARCustPOS.Execute("SELECT TOP 1 * FROM AR_CustomerPOS")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it
			
		    SQLCheckARCustPOS = "CREATE TABLE [AR_CustomerPOS]("
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [InternalRecordIdentifier] [int] IDENTITY(1,1) NOT NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [RecordCreationDateTime] [datetime] NULL CONSTRAINT [DF_AR_CustomerPOS_RecordCreationDateTime]  DEFAULT (getdate()),"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [CustIntRecID] int NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [ShipToIntRecID] int NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [PosId] [varchar](255) NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [PosName] [varchar](255) NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [LocationAddr1] [varchar](255) NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [LocationCity] [varchar](255) NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [LocationState] [varchar](50) NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [LocationZip] [varchar](50) NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [Contact] [varchar](255) NULL," 
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [ContactFirstName] [varchar](255) NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [ContactLastName] [varchar](255) NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [InstallDate] [datetime] NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [PosStatus] [varchar](50) NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [PrimarySalesperson] int NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [SecondarySalesperson] int NULL,"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " [Route] [varchar](255) NULL"
			SQLCheckARCustPOS = SQLCheckARCustPOS & " ) ON [PRIMARY]"      

		   Set rsCheckARCustPOS = cnnCheckARCustPOS.Execute(SQLCheckARCustPOS)
		   
		End If
	End If

on error goto 0

	SQLCheckARCustPOS  = "SELECT COL_LENGTH('AR_CustomerPOS', 'BillToIntRecID') AS IsItThere"
	Set rsCheckARCustPOS = cnnCheckARCustPOS.Execute(SQLCheckARCustPOS)
	If IsNull(rsCheckARCustPOS("IsItThere")) Then
		SQLCheckARCustPOS = "ALTER TABLE AR_CustomerPOS ADD [BillToIntRecID] int NULL"
		Set rsCheckARCustPOS = cnnCheckARCustPOS.Execute(SQLCheckARCustPOS)
	End If

		
	set rsCheckARCustPOS = nothing
	cnnCheckARCustPOS.close
	set cnnCheckARCustPOS = nothing
				
%>