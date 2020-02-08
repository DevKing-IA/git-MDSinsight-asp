<!--#include file="../inc/subsandfuncs.asp"-->
<!--#include file="../inc/InSightFuncs.asp"-->

<% 
maxsuggestion = 10

wordcount = 0
Dim Words()

txtkeyword = Request("query")
If txtkeyword<>"" Then
	txtkeyword = Replace(txtkeyword,"'","")
End If




SQL8 = "SELECT top "&maxsuggestion&" Company,FirstName,LastName,City,State, ProspectIntRecID AS Expr1  FROM PR_ProspectContactSearch  "
		SQL8 = SQL8 & " WHERE Company LIKE '%"&txtkeyword&"%'"
		SQL8 = SQL8 & " OR FirstName LIKE '%"&txtkeyword&"%'"
		SQL8 = SQL8 & " OR LastName LIKE '%"&txtkeyword&"%'"
		SQL8 = SQL8 & " ORDER BY Company ASC"

		'Response.write(SQL8)
		
		Set cnn8 = Server.CreateObject("ADODB.Connection")
		cnn8.CursorLocation = adUseClient ' Wierd but must do this
		cnn8.open (Session("ClientCnnString"))
		Set rs8 = Server.CreateObject("ADODB.Recordset")
		rs8.CursorLocation = adUseClient
		
		Set rs8 = cnn8.Execute(SQL8)

		
DO While Not rs8.EOF
	txt = rs8("Company") & " ("&rs8("FirstName") & " " &rs8("LastName") & ", " & rs8("City") & "," & rs8("State") &")"
	Call AddWord(txt,rs8("Expr1"))
    

rs8.MoveNext
Loop
rs8.Close
Set rs8 = Nothing



Response.Write("{"&vbcrlf)
Response.Write("""query"": ""Unit"","&vbcrlf)
Response.Write("""suggestions"": [")

If wordcount>0 Then
For t=0 To UBound(Words,2)
	If t = 0 Then
		'Response.Write("""" & Words(t) & """")
		Response.Write("{ ""value"":"""&Words(0,t)& """, ""data"": """&Words(1,t)&"""}")		 
	Else
		'Response.Write(",""" & Words(t) & """")
		Response.Write(",{ ""value"":"""&Words(0,t)& """, ""data"": """&Words(1,t)&"""}")
	End If	
	If t>maxsuggestion Then
		Exit For
	End If
Next
End If

Response.Write("]}"&vbcrlf)


cnn8.Close
Set cnn8 = Nothing

Sub AddWord(wword,id)
	f = false
	If wordcount>0 Then
		For k=0 To UBound(Words,2)
			If LCase(Words(0,k)) = LCase(wword) Then
				f = true
				Exit For
			End If
		Next
	End If
	
	If f = false Then
		Redim Preserve Words(2,wordcount)
		Words(0,wordcount) = wword
		Words(1,wordcount) = id
		wordcount = wordcount + 1
	End If
End Sub




%>

