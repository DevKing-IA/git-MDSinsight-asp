<script type="text/javascript">
	function monChanged1() {
	
		var str = document.getElementById('txtmonday').value;
		var res = str.substring(0, 2);
		var result = parseInt(res , 10)
		
        if (result > 6) {
            document.getElementById('lblMon').innerHTML = 'Invoices will post using Monday\'s date';
        } else {
      	  document.getElementById('lblMon').innerHTML = 'Invoices will post using Sunday\'s date';
		}
		
		if (str==null || str.trim()=="") {
			document.getElementById('lblMon').innerHTML = 'n/a';
		}
		
		}
		
	$(function () {
		monChanged1();
	});
</script>

<script type="text/javascript">
	function tueChanged1() {
	
		var str = document.getElementById('txttuesday').value;
		var res = str.substring(0, 2);
		var result = parseInt(res , 10)
		
        if (result > 6) {
            document.getElementById('lblTues').innerHTML = 'Invoices will post using Tuesday\'s date';
        } else {
      	  document.getElementById('lblTues').innerHTML = 'Invoices will post using Monday\'s date';
		}
		
		if (str==null || str.trim()=="") {
			document.getElementById('lblTues').innerHTML = 'n/a';
		}
		
		}
		
	$(function () {
		tueChanged1();
	});
</script>

<script type="text/javascript">
	function wedChanged1() {
	
		var str = document.getElementById('txtwednesday').value;
		var res = str.substring(0, 2);
		var result = parseInt(res , 10)
		
        if (result > 6) {
            document.getElementById('lblWed').innerHTML = 'Invoices will post using Wednesday\'s date';
        } else {
      	  document.getElementById('lblWed').innerHTML = 'Invoices will post using Tuesday\'s date';
		}
		
		if (str==null || str.trim()=="") {
			document.getElementById('lblWed').innerHTML = 'n/a';
		}

		}
		
	$(function () {
		wedChanged1();
	});
</script>

<script type="text/javascript">
	function thuChanged1() {
	
		var str = document.getElementById('txtthursday').value;
		var res = str.substring(0, 2);
		var result = parseInt(res , 10)
		
        if (result > 6) {
            document.getElementById('lblThu').innerHTML = 'Invoices will post using Thursday\'s date';
        } else {
      	  document.getElementById('lblThu').innerHTML = 'Invoices will post using Wednesday\'s date';
		}
				
		if (str==null || str.trim()=="") {
			document.getElementById('lblThu').innerHTML = 'n/a';
		}

		}
		
	$(function () {
		thuChanged1();
	});
</script>

<script type="text/javascript">
	function friChanged1() {
	
		var str = document.getElementById('txtfriday').value;
		var res = str.substring(0, 2);
		var result = parseInt(res , 10)
		
        if (result > 6) {
            document.getElementById('lblFri').innerHTML = 'Invoices will post using Friday\'s date';
        } else {
      	  document.getElementById('lblFri').innerHTML = 'Invoices will post using Thursday\'s date';
		}
						
		if (str==null || str.trim()=="") {
			document.getElementById('lblFri').innerHTML = 'n/a';
		}

		}
		
	$(function () {
		friChanged1();
	});
</script>

<script type="text/javascript">
	function satChanged1() {
	
		var str = document.getElementById('txtsaturday').value;
		var res = str.substring(0, 2);
		var result = parseInt(res , 10)
		
        if (result > 6) {
            document.getElementById('lblSat').innerHTML = 'Invoices will post using Saturday\'s date';
        } else {
      	  document.getElementById('lblSat').innerHTML = 'Invoices will post using Friday\'s date';
		}
								
		if (str==null || str.trim()=="") {
			document.getElementById('lblSat').innerHTML = 'n/a';
		}

		}
		
	$(function () {
		satChanged1();
	});
</script>

<script type="text/javascript">
	function sunChanged1() {
	
		var str = document.getElementById('txtsunday').value;
		var res = str.substring(0, 2);
		var result = parseInt(res , 10)
		
        if (result > 6) {
            document.getElementById('lblSun').innerHTML = 'Invoices will post using Sunday\'s date';
        } else {
      	  document.getElementById('lblSun').innerHTML = 'Invoices will post using Saturday\'s date';
		}
										
		if (str==null || str.trim()=="") {
			document.getElementById('lblSun').innerHTML = 'n/a';
		}

		}
		
	$(function () {
		sunChanged1();
	});
</script>

<script type="text/javascript">
	function monChanged() {
		$("#pnlMonday").hide();
		if($("#chkmonday").is(':checked'))
			$("#pnlMonday").show();
	}
	$(function () {
		monChanged();
	});
</script>	

<script type="text/javascript">
	function tueChanged() {
		$("#pnlTuesday").hide();
		if($("#chktuesday").is(':checked'))
			$("#pnlTuesday").show();
	}
	$(function () {
		tueChanged();
	});
</script>

<script type="text/javascript">
	function wedChanged() {
		$("#pnlWednesday").hide();
		if($("#chkwednesday").is(':checked'))
			$("#pnlWednesday").show();
	}
	$(function () {
		wedChanged();
	});
</script>

<script type="text/javascript">
	function thuChanged() {
		$("#pnlThursday").hide();
		if($("#chkthursday").is(':checked'))
			$("#pnlThursday").show();
	}
	$(function () {
		thuChanged();
	});
</script>

<script type="text/javascript">
	function friChanged() {
		$("#pnlFriday").hide();
		if($("#chkfriday").is(':checked'))
			$("#pnlFriday").show();
	}
	$(function () {
		friChanged();
	});
</script>

<script type="text/javascript">
	function satChanged() {
		$("#pnlSaturday").hide();
		if($("#chkSaturday").is(':checked'))
			$("#pnlSaturday").show();
	}
	$(function () {
		satChanged();
	});
</script>

<script type="text/javascript">
	function sunChanged() {
		$("#pnlSunday").hide();
		if($("#chkSunday").is(':checked'))
			$("#pnlSunday").show();
	}
	$(function () {
		sunChanged();
	});
</script>
