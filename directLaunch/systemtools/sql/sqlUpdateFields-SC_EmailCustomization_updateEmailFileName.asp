<%	

    Set cnnUpdateFileName = Server.CreateObject("ADODB.Connection")
    Set rsUpdateFileName = Server.CreateObject("ADODB.Recordset")
    cnnUpdateFileName.Open(Session("ClientCnnString"))

    'add emailFileName field if it is not there
	SQL_CheckEmailCustomization = "SELECT COL_LENGTH('SC_EmailCustomization', 'emailFileName') AS IsItThere"
	Set rsCheckEmailCustomization = cnnUpdateFileName.Execute(SQL_CheckEmailCustomization)
	If IsNull(rsCheckEmailCustomization("IsItThere")) Then
		SQL_CheckEmailCustomization = "ALTER TABLE SC_EmailCustomization ADD emailFileName varchar(200) NULL"
        Response.Write(SQL_CheckEmailCustomization & "<br/>")
		cnnUpdateFileName.Execute(SQL_CheckEmailCustomization)
	End If


    'update emailFileName
    SQL_emailFileName = "SELECT * FROM SC_EmailCustomization WHERE emailFileName IS NULL OR emailFileName=''"
    rsUpdateFileName.Open SQL_emailFileName,cnnUpdateFileName,3,3

    If Not rsUpdateFileName.Eof Then

	    Do While Not rsUpdateFileName.EOF
            emailFileName = ""
            emailName = rsUpdateFileName("emailName")
            emailType = rsUpdateFileName("emailType")
            If InStr(emailName,"Open")>0 Then
                If InStr(emailType, "Internal")>0 Then
                    emailFileName = "openServiceTicketInternal.txt"
                End If
                If InStr(emailType, "External")>0 Then
                    emailFileName = "openServiceTicketExternal.txt"
                End If
            End If

            If InStr(emailName,"Close")>0 Then
                If InStr(emailType, "Internal")>0 Then
                    emailFileName = "closeServiceTicketInternal.txt"
                End If
                If InStr(emailType, "External")>0 Then
                    emailFileName = "closeServiceTicketExternal.txt"
                End If
            End If

            If InStr(emailName,"Cancel")>0 Then
                If InStr(emailType, "Internal")>0 Then
                    emailFileName = "cancelServiceTicketInternal.txt"
                End If
                If InStr(emailType, "External")>0 Then
                    emailFileName = "cancelServiceTicketExternal.txt"
                End If
            End If

            If InStr(LCase(emailName),"swap")>0 Then
                emailFileName = "swap.txt"
            End if
            If emailFileName<>"" Then
                SQL_updateEmailFileName = "UPDATE SC_EmailCustomization SET emailFileName = '" & emailFileName & "' WHERE emailName='" & emailName & "' AND emailType='" & emailType & "'"
		        cnnUpdateFileName.Execute(SQL_updateEmailFileName)
                Response.Write(SQL_updateEmailFileName & "<br/>")
            End If
            
    	    rsUpdateFileName.movenext	
	    Loop

	    rsUpdateFileName.Close
	    Set rsUpdateFileName = Nothing
    End If
    cnnUpdateFileName.Close
	Set cnnUpdateFileName = Nothing

				
%>