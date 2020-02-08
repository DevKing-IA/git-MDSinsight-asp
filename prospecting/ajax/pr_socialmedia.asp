<!--#include file="../../inc/InSightFuncs.asp"-->
<!--#include file="../../inc/InSightFuncs_Users.asp"-->
<!--#include file="../../inc/InSightFuncs_Prospecting.asp"-->
<% If Session("Userno") = "" Then Response.End() %>
<%
Response.ContentType = "application/json"
ProspectIntRecID = Request.QueryString("i") 
If ProspectIntRecID = "" Then Response.End()

Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open (Session("ClientCnnString"))

If Request.Form("updateAction")="save" Then

	'***************************************************************************************
	'Lookup the record as it exists now so we can fillin the social media
	'***************************************************************************************
	SQL = "SELECT * FROM PR_ProspectSocialMedia WHERE ProspectIntRecID="&ProspectIntRecID &" AND InternalRecordIdentifier="&Request.Form("updateActionId")		

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn.Execute(SQL)
		
	If not rs.EOF Then
		Orig_SocialMediaLink = rs("SocialMediaLink")
		Orig_SocialMediaPlatform = rs("SocialMediaPlatform")
	End If


	
	set rs = Nothing

	

	'***************************************************************************************
	'Perform update on record in SQL
	'***************************************************************************************

	Query = "UPDATE PR_ProspectSocialMedia SET "		
	Query = Query & " SocialMediaPlatform='"&EscapeSingleQuotes(Request.Form("SocialMediaPlatform"))&"', "
	Query = Query & " SocialMediaLink='"&EscapeSingleQuotes(Request.Form("SocialMediaLink"))&"' "		
	Query = Query & "WHERE ProspectIntRecID="&ProspectIntRecID &" AND InternalRecordIdentifier="&Request.Form("updateActionId")
	
	cnn.Execute(Query)
	
	

	'***************************************************************************************
	'After SQL update, record entries in social media
	'***************************************************************************************
	
	SocialMediaLink		= Request.Form("SocialMediaLink")
	SocialMediaPlatform	= Request.Form("SocialMediaPlatform")

		
	
	'***********************************************************************
	'End Lookup the record as it exists now so we can fillin the social  media
	'***********************************************************************

	Description = ""
	

	
	
	If SocialMediaPlatform  <> Orig_SocialMediaPlatform Then
	
		Description =  "Social media platform has changed from " & Orig_SocialMediaPlatform  & " to " & SocialMediaPlatform & "  for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Social Media ") & " platform change ",GetTerm("Social Media"),"Minor",0,Description
		
		Description = "Social media platform changed from: <em><strong> " & Orig_SocialMediaPlatform   & "</em></strong> to: <em><strong>" & SocialMediaPlatform  & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If
	
	If SocialMediaLink  <> Orig_SocialMediaLink Then
	
		Description =  "Social media link has changed from " & Orig_SocialMediaLink & " to " & SocialMediaLink & "  for the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Social Media ") & " link change ",GetTerm("Social Media"),"Minor",0,Description
		
		Description = "Social media link changed from: <em><strong> " & Orig_SocialMediaLink   & "</em></strong> to: <em><strong>" & SocialMediaLink & "</em></strong>"
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")

	End If	
		

	
	

		
End If







If Request.Form("updateAction")="insert" Then


	SocialMediaPlatform= EscapeSingleQuotes(Request.Form("SocialMediaPlatform"))
	SocialMediaLink= EscapeSingleQuotes(Request.Form("SocialMediaLink"))

	Query = "INSERT INTO PR_ProspectSocialMedia (ProspectIntRecID, SocialMediaPlatform, SocialMediaLink) "
	Query = Query & " VALUES "
	Query = Query & "(" & ProspectIntRecID & ",'" & SocialMediaPlatform & "','" & SocialMediaLink & "') "	
	cnn.Execute(Query)
	


	
		Description = SocialMediaLink & " was added to the prospect " & GetProspectNameByNumber(ProspectIntRecID) 
		CreateAuditLogEntry GetTerm("Social Media") & " link added ",GetTerm("Social Media"),"Minor",0,Description
		
		Description = SocialMediaLink & " social media link was added to this prospect."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")
	
	
	
End If




If Request.Form("updateAction")="delete" Then

	SQL = "SELECT * FROM PR_ProspectSocialMedia WHERE ProspectIntRecID="&ProspectIntRecID &" AND InternalRecordIdentifier="&Request.Form("updateActionId")	

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn.Execute(SQL)
		
	If not rs.EOF Then
		SocialMediaLink = rs("SocialMediaLink")
	End If
	
	set rs = Nothing



	
		Description = "The social media link " & SocialMediaLink & "  was removed from the prospect " & GetProspectNameByNumber(ProspectIntRecID)
		CreateAuditLogEntry GetTerm("Social Media") & " " & SocialMediaLink& " link removed from prospect ",GetTerm("Social Media"),"Minor",0,Description
		
		Description = "The social media link  " & SocialMediaLink & " was removed from this prospect."
		Record_PR_Activity ProspectIntRecID, Description, Session("UserNo")


	Query = "DELETE FROM PR_ProspectSocialMedia WHERE ProspectIntRecID= " & ProspectIntRecID & " AND InternalRecordIdentifier = " & Request.Form("updateActionId") 
	cnn.Execute(Query)
	
End If





Query = "SELECT * FROM PR_ProspectSocialMedia WHERE ProspectIntRecID = " & ProspectIntRecID & " ORDER BY SocialMediaPlatform DESC, SocialMediaLink"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3 
Set rs = cnn.Execute(Query)

Response.Write("[")

If not rs.EOF Then
	sep = ""
	Do While Not rs.EOF
	
			SocialMediaID 		= rs("InternalRecordIdentifier")
			SocialMediaPlatform = rs("SocialMediaPlatform")
			SocialMediaLink 	= rs("SocialMediaLink")

			
				Response.Write(sep)
					sep = ","
					Response.Write("{")
					Response.Write("""SocialMediaID"":""" & EscapeQuotes(SocialMediaID) & """")
					Response.Write(",""SocialMediaPlatform"":""" & EscapeQuotes(SocialMediaPlatform) & """")
					Response.Write(",""SocialMediaLink"":""" & EscapeQuotes(SocialMediaLink) & """")

					
					
					Response.Write("}")
								
			
					
		rs.MoveNext						
	Loop
End If
					
Response.Write("]")
Set rs = Nothing
cnn.Close
Set cnn = Nothing

Function EscapeQuotes(val)
	If val <> "" Then
		EscapeQuotes = Replace(val, """", "\""")
	End If
End Function
Function EscapeSingleQuotes(val)
	If val <> "" Then
		EscapeSingleQuotes = Replace(val, "'", "''")
	End If
End Function

%> 
