<!-- Resources -->
<script src="https://www.amcharts.com/lib/3/amcharts.js"></script>
<script src="https://www.amcharts.com/lib/3/serial.js"></script>
<script src="https://www.amcharts.com/lib/3/plugins/export/export.min.js"></script>
<link rel="stylesheet" href="https://www.amcharts.com/lib/3/plugins/export/export.css" type="text/css" media="all" />
<script src="https://www.amcharts.com/lib/3/themes/light.js"></script>

<!-- Chart code -->
<script>

	var chart = AmCharts.makeChart("OrdersAPIActivityDiv", {
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
				"balloonText": "Orders:[[value]]",
				"fillAlphas": 0.8,
				"id": "AmGraph-1",
				"lineAlpha": 0.2,
				"title": "Orders",
				"type": "column",
				"valueField": "orders"
			},
			{
				"balloonText": "Invoices:[[value]]",
				"fillAlphas": 0.8,
				"id": "AmGraph-2",
				"lineAlpha": 0.2,
				"title": "Invoices",
				"type": "column",
				"valueField": "invoices"
			},
			{
				"balloonText": "RAs:[[value]]",
				"fillAlphas": 0.8,
				"id": "AmGraph-3",
				"lineAlpha": 0.2,
				"title": "RAs",
				"type": "column",
				"valueField": "ras"
			},
			{
				"balloonText": "Credit Memos:[[value]]",
				"fillAlphas": 0.8,
				"id": "AmGraph-4",
				"lineAlpha": 0.2,
				"title": "Credit Memos",
				"type": "column",
				"valueField": "creditmemos"
			},
			{
				"balloonText": "Summary Invoices:[[value]]",
				"fillAlphas": 0.8,
				"id": "AmGraph-5",
				"lineAlpha": 0.2,
				"title": "Summary Inv",
				"type": "column",
				"valueField": "summaryinvoices"
			}
			
			
		],
		"guides": [],
		"valueAxes": [
			{
				"id": "ValueAxis-1",
				"position": "bottom",
				"axisAlpha": 0,
				"title" : "Number of Records"
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
				"orders": '<%= SundayNumOrders %>',
				"invoices": '<%= SundayNumInvoices %>',
				"ras": '<%= SundayNumRAs %>',
				"creditmemos": '<%= SundayNumCMs %>',
				"summaryinvoices": '<%= SundayNumSummaryInvoices %>'
			},
			{
				"dayofweek": 'Monday',
				"orders": '<%= MondayNumOrders %>',
				"invoices": '<%= MondayNumInvoices %>',
				"ras": '<%= MondayNumRAs %>',
				"creditmemos": '<%= MondayNumCMs %>',
				"summaryinvoices": '<%= MondayNumSummaryInvoices %>'
			},
			{
				"dayofweek": 'Tuesday',
				"orders": '<%= TuesdayNumOrders %>',
				"invoices": '<%= TuesdayNumInvoices %>',
				"ras": '<%= TuesdayNumRAs %>',
				"creditmemos": '<%= TuesdayNumCMs %>',
				"summaryinvoices": '<%= TuesdayNumSummaryInvoices %>'
			},
			{
				"dayofweek": 'Wednesday',
				"orders": '<%= WednesdayNumOrders %>',
				"invoices": '<%=WednesdayNumInvoices %>',
				"ras": '<%= WednesdayNumRAs %>',
				"creditmemos": '<%= WednesdayNumCMs %>',
				"summaryinvoices": '<%= WednesdayNumSummaryInvoices %>'
			},
			{
				"dayofweek": 'Thursday',
				"orders": '<%= ThursdayNumOrders %>',
				"invoices": '<%= ThursdayNumInvoices %>',
				"ras": '<%= ThursdayNumRAs %>',
				"creditmemos": '<%= ThursdayNumCMs %>',
				"summaryinvoices": '<%= ThursdayNumSummaryInvoices %>'
			},
			{
				"dayofweek": 'Friday',
				"orders": '<%= FridayNumOrders %>',
				"invoices": '<%= FridayNumInvoices %>',
				"ras": '<%= FridayNumRAs %>',
				"creditmemos": '<%= FridayNumCMs %>',
				"summaryinvoices": '<%= FridayNumSummaryInvoices %>'
			},			
			{
				"dayofweek": 'Saturday',
				"orders": '<%= SaturdayNumOrders %>',
				"invoices": '<%= SaturdayNumInvoices %>',
				"ras": '<%= SaturdayNumRAs %>',
				"creditmemos": '<%= SaturdayNumCMs %>',
				"summaryinvoices": '<%= SaturdayNumSummaryInvoices %>'
			}
		],
	    "export": {
	    	"enabled": false
	     }
	
	});
</script>