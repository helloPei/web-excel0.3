<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ taglib uri="http://java.sun.com/jsp/jstl/core" prefix="c"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
<title>Excel</title>
</head>
<body>
<h3>${excel.excelName}</h3>
<h4>${excel.excelDate}-${excel.week}</h4>
<div class="table-responsive" style="margin-top: 20px">
    <table class="table" border="1" cellpadding="0" cellspacing="0">
        <thead style="background-color: #96b97d; border-color: #96b97d">
            <th style="width: 100px">Time</th>
            <th>Incoming Call Answer</th>
            <th>Incoming Average Holding Time Per Call (sec)</th>
            <th>Incoming Total Seconds In The Hour (sec)</th>
            <th>Outgoing Call Answer</th>
            <th>Outgoing Average Holding Time Per Call (sec)</th>
            <th>Outgoing Total Seconds In The Hour (sec)</th>
            <th>Service Capacity</th>
            <th>Capacity Needed</th>
            <th>Occupancy Hour(hour)</th>
            <th>Occupancy Rate(%)</th>
        </thead>
        <tbody id="tbodyId" style="border-color: #ccc">
            <c:forEach var="excel" items="${excelAll}">
                <tr>
                    <td>${excel.time}</td>
                    <td>${excel.inCallAnswer}</td>
                    <td>${excel.inAverageHoldingTPC}</td>
                    <td>${excel.inTotalHour}</td>
                    <td>${excel.outCallAnswer}</td>
                    <td>${excel.outAverageHoldingTPC}</td>
                    <td>${excel.outTotalHour}</td>
                    <td>${excel.serviceCapacity}</td>
                    <td>${excel.capacityNeeded}</td>
                    <td>${excel.occupancyHour}</td>
                    <td>${excel.occupancyRate}%</td>
                </tr>
            </c:forEach>
        </tbody>
    </table>
    <a href="javascript:doExport()"><button type="button" class="btn btn-primary">导出</button></a>
    <a href="doIndexUI.do"><button type="button" class="btn btn-primary">返回</button></a>
</div>
</body>
<script type="text/javascript" src="http://code.jquery.com/jquery-2.1.1.min.js"></script>
<script type="text/javascript">
$(function(){
	var occupancyRate = ${excel.occupancyRate};
	var trList = $("#tbodyId").children("tr")
	  for (var i = 0; i < trList.length; i++) {
	    var tdArr = trList.eq(i).find("td");
	    var tdVal = tdArr.eq(10).text();
		if(occupancyRate+"%" == tdVal || occupancyRate+"0%" == tdVal 
				|| occupancyRate+".00%" == tdVal) {
			tdArr.css('background-color','#FFBB66');
		}
	  }
});
function doExport(){
	var excelId = ${excel.excelId};
    window.location.href="doExport.do?excelId="+excelId;
}    
</script>
</html>