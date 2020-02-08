<!--#include file="adovbs.inc"-->
<!--#include file="settings.asp"-->
<!--#include file="InsightFuncs_Users.asp"-->
<%

Dim SQL, cnn, rs, SQL2, cnn2, rs2, ADOrst, rs3, rs4, SQL3, SQL4, cnn3, rs1, cnn1

Dim rsPricing, rsContract, rsContract1, rsProduct, SQL_Pricing, SQL_Contract, SQL_Contract1, SQL_Product, result, Group, Level

Dim rsCatXref, SQL_CatXref, prodSKU1, prodSKU2, RecToGet

Function stripHTML(strHTML)
'Strips the HTML tags from strHTML

  Dim objRegExp, strOutput
  Set objRegExp = New Regexp

  objRegExp.IgnoreCase = True
  objRegExp.Global = True
  objRegExp.Pattern = "<(.|\n)+?>"

  'Replace all HTML tag matches with the empty string
  strOutput = objRegExp.Replace(strHTML, "")
  
  'Replace all < and > with &lt; and &gt;
  strOutput = Replace(strOutput, "<", "&lt;")
  strOutput = Replace(strOutput, ">", "&gt;")
  
  stripHTML = strOutput    'Return the value of strOutput

  Set objRegExp = Nothing
End Function

Function Hacker_Filter1(strInput)

	dim i, badChars, newChars
	badChars = array("&","<",">","'","""","select","drop","!",":",";","--","insert","update","delete","xp_","@@","=","<script>","</script>","\", "/",")","(","[","]","|")
	newChars = strInput
	for i = 0 to uBound(badChars)
		newChars = replace(newChars, badChars(i), "")
	next
	Hacker_Filter1 = newChars
	
End Function


'*****************************************************************************************************************
'*****************************************************************************************************************





'**************************************************************************************
' HACKER SAFE FILTERING FUNCTIONS
'**************************************************************************************

Function Hacker_Filter1(strInput)

	dim i, badChars, newChars
	badChars = array("&","<",">","'","""","select","drop","!",":",";","--","insert","update","delete","xp_","@@","=","<script>","</script>","\", "/",")","(","[","]","|")
	newChars = strInput
	for i = 0 to uBound(badChars)
		newChars = replace(newChars, badChars(i), "")
	next
	Hacker_Filter1 = newChars
	
End Function

function Hacker_Filter2(cleanvar) 
 
       // Encode Percent 
       cleanvar = replace(cleanvar,"%", "&#37") 

       // Encode Ampersand 
       cleanvar = replace(cleanvar,"&", "&#38") 
 
       // Encode Single Quote 
       cleanvar = replace(cleanvar,"'", "&#39") 
 
       // Encode Double Quote 
       cleanvar = replace(cleanvar,"""", "&quot") 
 
       // Encode Less Than 
        cleanvar = replace(cleanvar,"<", "&lt") 
 
       // Encode Greater Than 
       cleanvar = replace(cleanvar,">", "&gt") 
 
       // Encode Close Bracket 
       cleanvar = replace(cleanvar,")", "&#41") 
 
       // Encode Open Bracket 
       cleanvar = replace(cleanvar,"(", "&#40") 
 
       // Encode Close Square Bracket 
       cleanvar = replace(cleanvar,"]", "&#93") 
 
       // Encode Open Square Bracket 
       cleanvar = replace(cleanvar,"[", "&#91") 
 
       // Encode Semicolon 
       cleanvar = replace(cleanvar,";", "&#59") 
 
       // Encode Colon 
       cleanvar = replace(cleanvar,":", "&#58") 
 
       // Encode Forward Slash 
       cleanvar = replace(cleanvar,"/", "&#47") 
 
       // Encode Left Brace 
       cleanvar = replace(cleanvar,"}", "&#125") 
 
       // Encode Right Brace 
       cleanvar = replace(cleanvar,"{", "&#123") 
 
       // Encode Exclamation 
       cleanvar = replace(cleanvar,"!", "&#33") 
 
       // Encode Double Dash 
       cleanvar = replace(cleanvar,"--", "&#45&#45") 
 
       // Encode Equal Sign 
       cleanvar = replace(cleanvar,"=", "&#61") 
 
       // Encode Underscore 
       'cleanvar = replace(cleanvar,"_", "&#95") 
 
       Hacker_Filter2 = cleanvar 
 
End Function
 
'**************************************************************************************
' END HACKER SAFE FILTERING FUNCTIONS
'**************************************************************************************
%>