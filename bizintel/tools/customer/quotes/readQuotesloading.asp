<!--#include file="../../../../inc/header.asp"-->

<%
If Request.QueryString("custID") = "" Then Response.Redirect ("reports.asp") Else custID = Request.QueryString("custID")
%>
<br><br>
<br><br>
<br><br>
<center>
<strong>Retrieving quotes from <%=GetTerm("Backend")%>, please wait...</strong>

<br><br>

<img src='../../../../img/loading.gif'/>

</center>

<script language="javascript">
     window.location = "readQuotes.asp?custID="+<%=custID%>;
</script>
