<%
Dim objStream
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Type = 1
objStream.Open
pathToFile = server.mappath(".") & "\captcha\" & right(left(session("captcha"),request("captchaID")),1) & ".gif"
objStream.LoadFromFile (pathToFile)
Response.ContentType = "image/gif"
Response.BinaryWrite objStream.Read
objStream.Close
Set objStream = Nothing
%>