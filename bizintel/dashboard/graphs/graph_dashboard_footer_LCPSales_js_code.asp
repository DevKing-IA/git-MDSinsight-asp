
<!-- Resources -->


<!-- Chart code -->
<script>
	
	var chart = AmCharts.makeChart( "chartdivLCPSalesRef", {
	
		"titles": [
				{
					"text": "Sales By Referral for <%=Session("TimePeriod")%>",
					"size": 18,
					"bold": "false"
				}				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[LCPreferral]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataLCPReferral%> ],
		"valueField": "contribPercent",
		"titleField":  "LCPreferral",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",
		"pullOutRadius": 0,
		"balloonText": "[[LCPreferral]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	
	
	var chart = AmCharts.makeChart( "chartdivLCPSalesCustType", {
	
		"titles": [
				{
					"text": "Sales By Cust Type for <%=Session("TimePeriod")%>",
					"size": 18,
					"bold": "false"
				}				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[LCPcusttype]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataLCPCustType%> ],
		"valueField": "contribPercent",
		"titleField": "LCPcusttype",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",		
		"pullOutRadius": 0,
		"balloonText": "[[LCPcusttype]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	
	var chart = AmCharts.makeChart( "chartdivLCPSalesSls1", {
	
		"titles": [
				{
					"text": "Sales By <%= GetTerm("Primary Salesman") %> for <%=Session("TimePeriod")%>",
					"size": 18,
					"bold": "false"
				}				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[LCPprimary]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataLCPSls1%> ],
		"valueField": "contribPercent",
		"titleField": "LCPprimary",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",		
		"pullOutRadius": 0,
		"balloonText": "[[LCPprimary]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	
	
	var chart = AmCharts.makeChart( "chartdivLCPSalesSls2", {
	
		"titles": [
				{
					"text": "Sales By <%= GetTerm("Secondary Salesman") %> for <%=Session("TimePeriod")%>",
					"size": 18,
					"bold": "false"
				}				
			],
		"creditsPosition":"bottom-left",
		"labelRadius": -50,
		"labelText": "[[LCPsecondary]]",
		"type": "pie",
		"theme": "light",
		"dataProvider": [ <%=amChartDataLCPSls2%> ],
		"valueField": "contribPercent",
		"titleField": "LCPsecondary",
		"maxLabelWidth" : "100",
		"autoMargins": false,
		"marginTop": 0,
		"marginBottom": 0,
		"marginLeft": 0,
		"marginRight": 0,
		"startEffect": "easeInSine",		
		"pullOutRadius": 0,
		"balloonText": "[[LCPsecondary]]" + ' ' + "[[contribPercent]]" + '% ($' + "[[contribDollars]]" + ')',
		"balloon":{
			"fixedPosition":true
		},
		"export": {
		"enabled": false
		}
	  
	});
	
var chart = AmCharts.makeChart("chartdivThisPerProj", {
	"titles": [
			{
				"text": "This Period Projected Sales",
				"size": 18,
				"bold": "false"
			}				
		],
    "theme": "light",
    "type": "serial",
	"startDuration": 1,
    "dataProvider": [ <%=amChartDataThisPerProj%>],
    "valueAxes": [{
        "position": "left",
        "title": "Sales Dollars"
    }],
    "graphs": [{
        "balloonText": "[[category]]: <b>[[value]]</b>",
        "fillColorsField": "color",
        "fillAlphas": 1,
        "lineAlpha": 0.1,
        "type": "column",
        "valueField": "dollars"
    }],
    "depth3D": 20,
	"angle": 30,
    "chartCursor": {
        "categoryBalloonEnabled": false,
        "cursorAlpha": 0,
        "zoomable": false
    },
    "categoryField": "period",
    "categoryAxis": {
        "gridPosition": "start",
        "labelRotation": 90
    },
    "export": {
    	"enabled": false
     }

});</script>
<!-- am chart js !-->  
