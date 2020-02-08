<html>
<head>
<title>Metroplex</title>
<script type="text/javascript">
  function checkForm(form)
  {
    var elem = document.getElementById('status');
    elem.style.visibility='visible'
    return true;
  }
</script>
</head>
<body>
<form method="get" action="http://98.6.75.158:3291/mds/SelectTest4.do" onsubmit="return checkForm(this);">
<p>
<input type="submit" value="Download Products Spreadsheet"/>
</form>
<div id="status" style="color: red; visibility: hidden;"><br><br>Generating Excel Spreadsheet. Please wait...</div>
</body>
</html>