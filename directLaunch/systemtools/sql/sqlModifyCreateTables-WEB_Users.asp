<%	

	Set cnnWEB_Users = Server.CreateObject("ADODB.Connection")
	cnnWEB_Users.open (Session("ClientCnnString"))
	Set rsWEB_Users = Server.CreateObject("ADODB.Recordset")
	rsWEB_Users.CursorLocation = 3 


	Err.Clear
	on error resume next
	Set rsWEB_Users = cnnWEB_Users.Execute("SELECT TOP 1 * FROM WEB_Users")
	If Err.Description <> "" Then
		If InStr(Ucase(Err.Description),"INVALID OBJECT NAME") <> 0 Then
			On error goto 0		
			'The table is not there, we need to create it	
			
			
			SQLWEB_Users = "CREATE TABLE [WEB_Users]( "
            SQLWEB_Users = SQLWEB_Users & "[userNo] [int] IDENTITY(1610,1) NOT NULL, "
            SQLWEB_Users = SQLWEB_Users & "[CustID] [varchar](15) NOT NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userName] [varchar](75) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userPhone] [varchar](25) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userPhoneExt] [varchar](10) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userFax] [varchar](25) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userEmail] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userPassword] [varchar](25) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userAddress1] [varchar](50) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userAddress2] [varchar](50) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCity] [varchar](50) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userState] [varchar](25) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userZip] [varchar](10) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCompany] [varchar](50) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userPO] [varchar](50) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userDepartment] [varchar](50) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCostCenter] [varchar](50) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCDescription] [varchar](50) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userDevDate] [varchar](50) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userReceipt] [bit] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userPOP] [int] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCemail] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCemail2] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCemail3] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCemail4] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCemail5] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCemail6] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCemail7] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCemail8] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCemail9] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userCCemail10] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userInternalCC] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userRestricted] [bit] NULL CONSTRAINT [DF_WEB_Users_userRestricted]  DEFAULT ((0)), "
            SQLWEB_Users = SQLWEB_Users & "[userRestricted_Level2] [bit] NULL CONSTRAINT [DF_WEB_Users_userRestricted_Level2]  DEFAULT ((0)), "
            SQLWEB_Users = SQLWEB_Users & "[userGetsSpecial] [bit] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userGetsFreight] [bit] NULL CONSTRAINT [DF_WEB_Users_userGetsFreight]  DEFAULT ((1)), "
            SQLWEB_Users = SQLWEB_Users & "[userClearanceEmail] [bit] NOT NULL CONSTRAINT [DF_WEB_Users_userClearanceEmail]  DEFAULT ((0)), "
            SQLWEB_Users = SQLWEB_Users & "[userReminderEmail] [bit] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userLastEmail] [datetime] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userTaxCode1] [varchar](2) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userTaxCode2] [varchar](2) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userTaxCode3] [varchar](2) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userEditShipTo] [bit] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[FTM_Enabled] [bit] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[FTM_Limit] [money] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[FTM_Approver] [int] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[FTM_Master] [bit] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userLast2ndEmail] [datetime] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userGetsNewsLetter] [bit] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userLastNewsFile] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userLastNewsDate] [datetime] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[customURL] [varchar](100) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[userAdded] [datetime] NULL, "
            SQLWEB_Users = SQLWEB_Users & "[billCardOnFile] [varchar](50) NULL CONSTRAINT [DF_WEB_Users_billCardOnFile]  DEFAULT ('INVOICE'), "
            SQLWEB_Users = SQLWEB_Users & "[userType] [varchar](50) NULL CONSTRAINT [DF_WEB_Users_userType]  DEFAULT ('B'), "
            SQLWEB_Users = SQLWEB_Users & "[userReminderStatus] [varchar](255) NULL, "
            SQLWEB_Users = SQLWEB_Users & "[showSuggestedItems] [bit] NULL CONSTRAINT [DF_WEB_Users_showSuggestedItems]  DEFAULT ((1)), "
            SQLWEB_Users = SQLWEB_Users & "[userPaymentType] [varchar](255) NULL CONSTRAINT [DF_WEB_Users_userPaymentType]  DEFAULT ('INVOICE'), "
            SQLWEB_Users = SQLWEB_Users & "[userDisplayCCInfo] [bit] NULL CONSTRAINT [DF_WEB_Users_userDisplayCCInfo]  DEFAULT ((1)), "
            SQLWEB_Users = SQLWEB_Users & "[user_AR_ShowUnappliedCredits] [bit] NULL CONSTRAINT [DF_WEB_Users_user_AR_ShowUnappliedCredits_1]  DEFAULT ((0)), "
            SQLWEB_Users = SQLWEB_Users & "[user_AR_ShowARInvoices] [bit] NULL CONSTRAINT [DF_WEB_Users_user_AR_ShowARInvoices_1]  DEFAULT ((1)), "
            SQLWEB_Users = SQLWEB_Users & "[user_AR_SubMasterBoth] [varchar](50) NULL CONSTRAINT [DF_WEB_Users_user_AR_SubMasterBoth_1]  DEFAULT ('B'), "
            SQLWEB_Users = SQLWEB_Users & "[user_AR_ShowZeroDollarItems] [bit] NULL CONSTRAINT [DF_WEB_Users_user_AR_ShowZeroDollarItems_1]  DEFAULT ((0)), "
			SQLWEB_Users = SQLWEB_Users & "CONSTRAINT [PK_WEB_Users] PRIMARY KEY CLUSTERED  "
			SQLWEB_Users = SQLWEB_Users & "( "
			SQLWEB_Users = SQLWEB_Users & " [userNo] ASC "
			SQLWEB_Users = SQLWEB_Users & ")WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY] "
			SQLWEB_Users = SQLWEB_Users & ") ON [PRIMARY] "

			
			Set rsWEB_Users = cnnWEB_Users.Execute(SQLWEB_Users)
		End If
	End If
	
	
	set rsWEB_Users = nothing
	cnnWEB_Users.close
	set cnnWEB_Users = nothing


%>