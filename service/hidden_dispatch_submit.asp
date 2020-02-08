<%

MemoToToggle = Request.querystring("txtMemoToToggle")

Set cnnTemp = Server.CreateObject("ADODB.Connection")
cnnTemp.open (Session("ClientCnnString"))
Set rsTemp = Server.CreateObject("ADODB.Recordset")
rsTemp.CursorLocation = 3 

SQLtemp = "Update FS_ServiceMemos Set Dispatched = CASE WHEN Dispatched = 0 THEN -1 ELSE 0 END Where MemoNumber = '"  & MemoToToggle & "'"

'Response.Write(SQLtemp)

Set rsTemp = cnnTemp.Execute(SQLtemp)

set rsTemp = Nothing
cnnTemp.close
set cnnTemp = nothing

Response.redirect ("main.asp")
%>















