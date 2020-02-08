<!--#include file="../../../inc/header.asp"-->

<%

	'***********************************************************
	'Get Values Of All Form Fields Posted
	'***********************************************************

	Dim CatGroup_CategoryArray(22)
	Dim CatGroup_SortOrderArray(22)
	Dim CatGroup_GroupNameArray(22)
	Dim CatGroup_ShowOnGArray(22)
	Dim CatGroup_InternalID(22)
	
			
	CatGroup_CategoryArray(0) = Request.Form("txtCatID0")
	CatGroup_SortOrderArray(0) = Request.Form("txtSortOrder0")
	CatGroup_GroupNameArray(0) = Request.Form("txtGroupName0")
	CatGroup_ShowOnGArray(0) = Request.Form("chkGScreen0")
	CatGroup_CategoryArray(1) = Request.Form("txtCatID1")
	CatGroup_SortOrderArray(1) = Request.Form("txtSortOrder1")
	CatGroup_GroupNameArray(1) = Request.Form("txtGroupName1")
	CatGroup_ShowOnGArray(1) = Request.Form("chkGScreen1")
	CatGroup_CategoryArray(2) = Request.Form("txtCatID2")
	CatGroup_SortOrderArray(2) = Request.Form("txtSortOrder2")
	CatGroup_GroupNameArray(2) = Request.Form("txtGroupName2")
	CatGroup_ShowOnGArray(2) = Request.Form("chkGScreen2")
	CatGroup_CategoryArray(3) = Request.Form("txtCatID3")
	CatGroup_SortOrderArray(3) = Request.Form("txtSortOrder3")
	CatGroup_GroupNameArray(3) = Request.Form("txtGroupName3")
	CatGroup_ShowOnGArray(3) = Request.Form("chkGScreen3")
	CatGroup_CategoryArray(4) = Request.Form("txtCatID4")
	CatGroup_SortOrderArray(4) = Request.Form("txtSortOrder4")
	CatGroup_GroupNameArray(4) = Request.Form("txtGroupName4")
	CatGroup_ShowOnGArray(4) = Request.Form("chkGScreen4")
	CatGroup_CategoryArray(5) = Request.Form("txtCatID5")
	CatGroup_SortOrderArray(5) = Request.Form("txtSortOrder5")
	CatGroup_GroupNameArray(5) = Request.Form("txtGroupName5")
	CatGroup_ShowOnGArray(5) = Request.Form("chkGScreen5")
	CatGroup_CategoryArray(6) = Request.Form("txtCatID6")
	CatGroup_SortOrderArray(6) = Request.Form("txtSortOrder6")
	CatGroup_GroupNameArray(6) = Request.Form("txtGroupName6")
	CatGroup_ShowOnGArray(6) = Request.Form("chkGScreen6")
	CatGroup_CategoryArray(7) = Request.Form("txtCatID7")
	CatGroup_SortOrderArray(7) = Request.Form("txtSortOrder7")
	CatGroup_GroupNameArray(7) = Request.Form("txtGroupName7")
	CatGroup_ShowOnGArray(7) = Request.Form("chkGScreen7")
	CatGroup_CategoryArray(8) = Request.Form("txtCatID8")
	CatGroup_SortOrderArray(8) = Request.Form("txtSortOrder8")
	CatGroup_GroupNameArray(8) = Request.Form("txtGroupName8")
	CatGroup_ShowOnGArray(8) = Request.Form("chkGScreen8")
	CatGroup_CategoryArray(9) = Request.Form("txtCatID9")
	CatGroup_SortOrderArray(9) = Request.Form("txtSortOrder9")
	CatGroup_GroupNameArray(9) = Request.Form("txtGroupName9")
	CatGroup_ShowOnGArray(9) = Request.Form("chkGScreen9")
	CatGroup_CategoryArray(10) = Request.Form("txtCatID10")
	CatGroup_SortOrderArray(10) = Request.Form("txtSortOrder10")
	CatGroup_GroupNameArray(10) = Request.Form("txtGroupName10")
	CatGroup_ShowOnGArray(10) = Request.Form("chkGScreen10")
	CatGroup_CategoryArray(11) = Request.Form("txtCatID11")
	CatGroup_SortOrderArray(11) = Request.Form("txtSortOrder11")
	CatGroup_GroupNameArray(11) = Request.Form("txtGroupName11")
	CatGroup_ShowOnGArray(11) = Request.Form("chkGScreen11")
	CatGroup_CategoryArray(12) = Request.Form("txtCatID12")
	CatGroup_SortOrderArray(12) = Request.Form("txtSortOrder12")
	CatGroup_GroupNameArray(12) = Request.Form("txtGroupName12")
	CatGroup_ShowOnGArray(12) = Request.Form("chkGScreen12")
	CatGroup_CategoryArray(13) = Request.Form("txtCatID13")
	CatGroup_SortOrderArray(13) = Request.Form("txtSortOrder13")
	CatGroup_GroupNameArray(13) = Request.Form("txtGroupName13")
	CatGroup_ShowOnGArray(13) = Request.Form("chkGScreen13")
	CatGroup_CategoryArray(14) = Request.Form("txtCatID14")
	CatGroup_SortOrderArray(14) = Request.Form("txtSortOrder14")
	CatGroup_GroupNameArray(14) = Request.Form("txtGroupName14")
	CatGroup_ShowOnGArray(14) = Request.Form("chkGScreen14")
	CatGroup_CategoryArray(15) = Request.Form("txtCatID15")
	CatGroup_SortOrderArray(15) = Request.Form("txtSortOrder15")
	CatGroup_GroupNameArray(15) = Request.Form("txtGroupName15")
	CatGroup_ShowOnGArray(15) = Request.Form("chkGScreen15")
	CatGroup_CategoryArray(16) = Request.Form("txtCatID16")
	CatGroup_SortOrderArray(16) = Request.Form("txtSortOrder16")
	CatGroup_GroupNameArray(16) = Request.Form("txtGroupName16")
	CatGroup_ShowOnGArray(16) = Request.Form("chkGScreen16")
	CatGroup_CategoryArray(17) = Request.Form("txtCatID17")
	CatGroup_SortOrderArray(17) = Request.Form("txtSortOrder17")
	CatGroup_GroupNameArray(17) = Request.Form("txtGroupName17")
	CatGroup_ShowOnGArray(17) = Request.Form("chkGScreen17")
	CatGroup_CategoryArray(18) = Request.Form("txtCatID18")
	CatGroup_SortOrderArray(18) = Request.Form("txtSortOrder18")
	CatGroup_GroupNameArray(18) = Request.Form("txtGroupName18")
	CatGroup_ShowOnGArray(18) = Request.Form("chkGScreen18")
	CatGroup_CategoryArray(19) = Request.Form("txtCatID19")
	CatGroup_SortOrderArray(19) = Request.Form("txtSortOrder19")
	CatGroup_GroupNameArray(19) = Request.Form("txtGroupName19")
	CatGroup_ShowOnGArray(19) = Request.Form("chkGScreen19")
	CatGroup_CategoryArray(20) = Request.Form("txtCatID20")
	CatGroup_SortOrderArray(20) = Request.Form("txtSortOrder20")
	CatGroup_GroupNameArray(20) = Request.Form("txtGroupName20")
	CatGroup_ShowOnGArray(20) = Request.Form("20")
	CatGroup_CategoryArray(21) = Request.Form("txtCatID21")
	CatGroup_SortOrderArray(21) = Request.Form("txtSortOrder21")
	CatGroup_GroupNameArray(21) = Request.Form("txtGroupName21")
	CatGroup_ShowOnGArray(21) = Request.Form("chkGScreen21")


		
	'**********************************************************
	' Now do the audit trail entries for the grouped categories
	'**********************************************************
	
	'Now re-select in case we just did the first time insert
	
	SQL = "SELECT * FROM Settings_CatGroups order by Category"
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 
	Set rs = cnn8.Execute(SQL)


	x = 0

	If not rs.EOF Then
		Do
			If CatGroup_GroupNameArray(x) <> rs("GroupName") Then
				CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Category " & CatGroup_CategoryArray(x) & " - " &  GetCategoryByID(x) & " changed to a new group. Changed from  " & rs("GroupName") & " to " & CatGroup_GroupNameArray(x)
			End If
			If Cint(CatGroup_SortOrderArray(x)) <> rs("SortOrder") Then
				CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Minor", 1, "Category " & CatGroup_CategoryArray(x) & " - " &  GetCategoryByID(x) & " Sort order changed from  " & rs("SortOrder") & " to " & CatGroup_SortOrderArray(x)
			End If
			If CatGroup_ShowOnGArray(x) = "on" Then EvalForGScreen = vbTrue Else EvalForGScreen = vbFalse
			If EvalForGScreen  <> rs("ShowOnGScreen") Then
				'CHange it again just to get the wording right
				If EvalForGScreen = vbTrue then EvalForGScreen = "True" Else EvalForGScreen = "False"
				CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Category " & CatGroup_CategoryArray(x) & " - " &  GetCategoryByID(x) & " Show On Period Sales Screen changed from  " & rs("ShowOnGScreen") & " to " & EvalForGScreen 
			End If
			x=x+1
			rs.movenext
		Loop until rs.eof
	End If
		
		
	'***********************************************************
	'Update SQL with Request Form Field Data
	'***********************************************************
	'Update Category Group Table as well
	Set cnn8 = Server.CreateObject("ADODB.Connection")
	cnn8.open (Session("ClientCnnString"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3 

	For x = 0 to 21
		'This line has to stay here
		If Trim(CatGroup_GroupNameArray(x)) = "" Then 	
			CreateAuditLogEntry "Global Settings Change", "Global Settings Change", "Major", 1, "Category " & CatGroup_CategoryArray(x) & " - " &  GetCategoryByID(x) & " The group name was left blank. Insight automatically reset it to " & GetCategoryByID(x)
		End If				
		SQL = "UPDATE  " & MUV_Read("SQL_Owner") & ".Settings_CatGroups SET "
		If Trim(CatGroup_GroupNameArray(x)) <> "" Then 
			SQL = SQL & " GroupName = '" & CatGroup_GroupNameArray(x) & "',"
		Else
			SQL = SQL & " GroupName = '" & GetCategoryByID(x) & "'," ' To ensure no blank group names
		End If
		SQL = SQL & " SortOrder = '" & CatGroup_SortOrderArray(x) & "',"
		SQL = SQL & " ShowOnGScreen = " 
		If CatGroup_ShowOnGArray(x)="on" then SQL = SQL & vbTrue Else SQL = SQL & vbFalse
		SQL = SQL & " WHERE Category = " & x
		'response.write(SQL)
		Set rs = cnn8.Execute(SQL)
	Next 	
	
	
	
	Response.Redirect("category-groupings.asp")
%>
<!--#include file="../../../inc/footer-main.asp"-->