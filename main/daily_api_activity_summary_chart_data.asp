<!--#include file="../inc/InsightFuncs_API.asp"-->

<%

SundayNumOrders = GetNumberOfOrdersByDate(DateAdd("d",(WeekDay(Date())-1)*-1,Date()))
SundayNumInvoices = GetNumberOfInvoicesByDate(DateAdd("d",(WeekDay(Date())-1)*-1,Date()))
SundayNumRAs = GetNumberOfRAsByDate(DateAdd("d",(WeekDay(Date())-1)*-1,Date()))
SundayNumCMs = GetNumberOfCMsByDate(DateAdd("d",(WeekDay(Date())-1)*-1,Date()))
SundayNumSummaryInvoices = GetNumberOfSummaryInvoicesByDate(DateAdd("d",(WeekDay(Date())-1)*-1,Date()))


MondayNumOrders = GetNumberOfOrdersByDate(DateAdd("d",(WeekDay(Date())-2)*-1,Date()))
MondayNumInvoices = GetNumberOfInvoicesByDate(DateAdd("d",(WeekDay(Date())-2)*-1,Date()))
MondayNumRAs = GetNumberOfRAsByDate(DateAdd("d",(WeekDay(Date())-2)*-1,Date()))
MondayNumCMs = GetNumberOfCMsByDate(DateAdd("d",(WeekDay(Date())-2)*-1,Date()))
MondayNumSummaryInvoices = GetNumberOfSummaryInvoicesByDate(DateAdd("d",(WeekDay(Date())-2)*-1,Date()))


TuesdayNumOrders = GetNumberOfOrdersByDate(DateAdd("d",(WeekDay(Date())-3)*-1,Date()))
TuesdayNumInvoices = GetNumberOfInvoicesByDate(DateAdd("d",(WeekDay(Date())-3)*-1,Date()))
TuesdayNumRAs = GetNumberOfRAsByDate(DateAdd("d",(WeekDay(Date())-3)*-1,Date()))
TuesdayNumCMs = GetNumberOfCMsByDate(DateAdd("d",(WeekDay(Date())-3)*-1,Date()))
TuesdayNumSummaryInvoices = GetNumberOfSummaryInvoicesByDate(DateAdd("d",(WeekDay(Date())-3)*-1,Date()))


WednesdayNumOrders = GetNumberOfOrdersByDate(DateAdd("d",(WeekDay(Date())-4)*-1,Date()))
WednesdayNumInvoices = GetNumberOfInvoicesByDate(DateAdd("d",(WeekDay(Date())-4)*-1,Date()))
WednesdayNumRAs = GetNumberOfRAsByDate(DateAdd("d",(WeekDay(Date())-4)*-1,Date()))
WednesdayNumCMs = GetNumberOfCMsByDate(DateAdd("d",(WeekDay(Date())-4)*-1,Date()))
WednesdayNumSummaryInvoices = GetNumberOfSummaryInvoicesByDate(DateAdd("d",(WeekDay(Date())-4)*-1,Date()))


ThursdayNumOrders = GetNumberOfOrdersByDate(DateAdd("d",(WeekDay(Date())-5)*-1,Date()))
ThursdayNumInvoices = GetNumberOfInvoicesByDate(DateAdd("d",(WeekDay(Date())-5)*-1,Date()))
ThursdayNumRAs = GetNumberOfRAsByDate(DateAdd("d",(WeekDay(Date())-5)*-1,Date()))
ThursdayNumCMs = GetNumberOfCMsByDate(DateAdd("d",(WeekDay(Date())-5)*-1,Date()))
ThursdayNumSummaryInvoices = GetNumberOfSummaryInvoicesByDate(DateAdd("d",(WeekDay(Date())-5)*-1,Date()))


FridayNumOrders = GetNumberOfOrdersByDate(DateAdd("d",(WeekDay(Date())-6)*-1,Date()))
FridayNumInvoices = GetNumberOfInvoicesByDate(DateAdd("d",(WeekDay(Date())-6)*-1,Date()))
FridayNumRAs = GetNumberOfRAsByDate(DateAdd("d",(WeekDay(Date())-6)*-1,Date()))
FridayNumCMs = GetNumberOfCMsByDate(DateAdd("d",(WeekDay(Date())-6)*-1,Date()))
FridayNumSummaryInvoices = GetNumberOfSummaryInvoicesByDate(DateAdd("d",(WeekDay(Date())-6)*-1,Date()))


SaturdayNumOrders = GetNumberOfOrdersByDate(DateAdd("d",(WeekDay(Date())-7)*-1,Date()))
SaturdayNumInvoices = GetNumberOfInvoicesByDate(DateAdd("d",(WeekDay(Date())-7)*-1,Date()))
SaturdayNumRAs = GetNumberOfRAsByDate(DateAdd("d",(WeekDay(Date())-7)*-1,Date()))
SaturdayNumCMs = GetNumberOfCMsByDate(DateAdd("d",(WeekDay(Date())-7)*-1,Date()))
SaturdayNumSummaryInvoices = GetNumberOfSummaryInvoicesByDate(DateAdd("d",(WeekDay(Date())-7)*-1,Date()))


DateRangeTitleForAPIGraph = DateAdd("d",(WeekDay(Date())-1)*-1,Date()) & " - " & DateAdd("d",(WeekDay(Date())-7)*-1,Date())


%>