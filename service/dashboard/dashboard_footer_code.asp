
<!-- Resources -->
<script src="https://www.amcharts.com/lib/3/amcharts.js"></script>
<script src="https://www.amcharts.com/lib/3/pie.js"></script>
<script src="https://www.amcharts.com/lib/3/serial.js"></script>
<script src="https://www.amcharts.com/lib/3/plugins/export/export.min.js"></script>

<head>
<link rel="stylesheet" href="https://www.amcharts.com/lib/3/plugins/export/export.css" type="text/css" media="all" />
<script src="https://www.amcharts.com/lib/3/themes/light.js"></script>

<link rel="stylesheet" href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" />
<script type="text/javascript" src="https://cdn.datatables.net/1.10.16/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" src="https://cdn.datatables.net/plug-ins/1.10.18/sorting/currency.js"></script>

<script type="text/javascript">

$(document).ready(function() {

    $("#PleaseWaitPanel").hide();

});
</script>
<!-- Chart code -->
<script>

	var chart = AmCharts.makeChart("ServiceCallActivityDiv", {
		"type": "serial",
	    "theme": "light",
		"categoryField": "dayofweek",
		"rotate": false,
		"startDuration": 1,
		"categoryAxis": {
			"gridPosition": "start",
			"position": "left"
		},
		"trendLines": [],
		"graphs": [
			{
				"balloonText": "Open:[[value]]",
				"fillAlphas": 0.8,
				"id": "AmGraph-1",
				"lineAlpha": 0.2,
				"title": "Open",
				"type": "column",
				"valueField": "open"
			},
			{
				"balloonText": "Close:[[value]]",
				"fillAlphas": 0.8,
				"id": "AmGraph-2",
				"lineAlpha": 0.2,
				"title": "Closed",
				"type": "column",
				"valueField": "close"
			},
			{
				"balloonText": "Cancel:[[value]]",
				"fillAlphas": 0.8,
				"id": "AmGraph-3",
				"lineAlpha": 0.2,
				"title": "Cancelled",
				"type": "column",
				"valueField": "cancel"
			},
			{
				"balloonText": "Awaiting Dispatch:[[value]]",
				"fillAlphas": 0.8,
				"id": "AmGraph-4",
				"lineAlpha": 0.2,
				"title": "Awaiting Dispatch",
				"type": "column",
				"valueField": "awaitdispatch"
			}
			
			
		],
		"guides": [],
		"valueAxes": [
			{
				"id": "ValueAxis-1",
				"position": "bottom",
				"axisAlpha": 0,
				"title" : "Number of Tickets"
			}
		],
		"allLabels": [],
		"balloon": {},
		"titles": [],
	    "legend": {
	    	"useGraphSettings": true
	  	},		
		"dataProvider": [
			{
				"dayofweek": 'Sunday',
				"open": '<%= SundayOpen %>',
				"close": '<%= SundayClose %>',
				"cancel": '<%= SundayCancel %>',
				"awaitdispatch": '<%= SundayDisp %>'
			},
			{
				"dayofweek": 'Monday',
				"open": '<%= MondayOpen %>',
				"close": '<%= MondayClose %>',
				"cancel": '<%= MondayCancel %>',
				"awaitdispatch": '<%= MondayDisp %>'
			},
			{
				"dayofweek": 'Tuesday',
				"open": '<%= TuesdayOpen %>',
				"close": '<%= TuesdayClose %>',
				"cancel": '<%= TuesdayCancel %>',
				"awaitdispatch": '<%= TuesdayDisp %>'
			},
			{
				"dayofweek": 'Wednesday',
				"open": '<%= WednesdayOpen %>',
				"close": '<%= WednesdayClose %>',
				"cancel": '<%= WednesdayCancel %>',
				"awaitdispatch": '<%= WednesdayDisp %>'
			},
			{
				"dayofweek": 'Thursday',
				"open": '<%= ThursdayOpen %>',
				"close": '<%= ThursdayClose %>',
				"cancel": '<%= ThursdayCancel %>',
				"awaitdispatch": '<%= ThursdayDisp %>'
			},
			{
				"dayofweek": 'Friday',
				"open": '<%= FridayOpen %>',
				"close": '<%= FridayClose %>',
				"cancel": '<%= FridayCancel %>',
				"awaitdispatch": '<%= FridayDisp %>'
			},			
			{
				"dayofweek": 'Saturday',
				"open": '<%= SaturdayOpen %>',
				"close": '<%= SaturdayClose %>',
				"cancel": '<%= SaturdayCancel %>',
				"awaitdispatch": '<%= SaturdayDisp %>'
			}
		],
	    "export": {
	    	"enabled": true
	     }
	
	});
</script>