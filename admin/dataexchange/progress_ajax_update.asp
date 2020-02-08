<%@EnableSessionState=False%>
<%
	response.expires = -1
	response.contenttype = "text/xml"

	PID = Request.QueryString("pid")

	Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
	Response.Write UploadProgress.XmlProgress(PID)
%>