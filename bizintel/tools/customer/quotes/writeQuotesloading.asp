<!--#include file="../../../../inc/header.asp"-->
<%
'Doesn't need to be passed a customer number becuase there is only 1 cust in our SQL work table
'but we check it just so if this page is loaded by spiders, etc, it doesn't execute
If Request.QueryString("custID") = "" Then Response.Redirect ("reports.asp") Else custID = Request.QueryString("custID")
%>
<br><br>
<br><br>
<br><br>
<center>
<strong>Sending quotes to your <%=GetTerm("Backend")%>, please wait...</strong>

<br><br>

<img src='../../../../img/loading.gif'/>

</center>

<script language="javascript">
     window.location = "writeQuotes.asp?custID="+<%=custID%>;
</script>
