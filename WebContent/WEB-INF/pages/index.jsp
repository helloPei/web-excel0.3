<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
    <title>Excel</title>
</head>
<body>
<div style="margin-top: 20px">
	Excel Name : 
	<input type="text" id="excelName" name="excelName" placeholder="请输入Excel名称"/>
	Excel Day : 
	<input type="text" id="excelDate" name="excelDate" placeholder="查询日期格式：yyyy-mm-dd" disabled="disabled"/>
	<button onclick="doSearch()" type="button">搜索</button>
	<button onclick="doDelete()" type="button">删除</button>
	<br/>
	<input type="checkbox" name="isSearchMax" onclick="doState()">Occupancy Rate (max)
	<br/>
	<br/>
	<form id="form_table" method="post" enctype="multipart/form-data">
        <button onclick="doSubmit()" type="submit">导入</button>
        <input id="file_excel" type="file" name="filename" multiple="multiple" accept=".xls,.xlsx,.csv"></input>
    </form>
</div>
<div style="margin-top: 20px">
    <table border="1" cellpadding="0" cellspacing="0">
        <thead style="background-color: #96b97d; border-color: #96b97d">
            <th><input type="checkbox" id="checkAllId">全选</th>
            <th align="center" style="width:400px">Excel Name</th>
            <th align="center" style="width:130px">Excel Type</th>
            <th align="center" style="width:100px">Excel Date</th>
            <th align="center" style="width:100px">Week</th>
            <th align="center" style="width:200px">Import Date</th>
            <th align="center">Occupancy Rate(%)</th>
        </thead>
        <tbody id="tbodyId" style="border-color: #ccc"></tbody>
    </table>
</div>
<div id="pageId" style="margin-top: 10px">
    <table border="1" cellpadding="0" cellspacing="0">
        <tr>
            <td><button type='button' class="first">首页</button></td>
            <td><button type='button' class="pre">上一页</button></td>
            <td><button type='button' class="next">下一页</button></td>
            <td><button type='button' class="rowCount">总记录数(0)</button></td>
            <td><button type='button' class="pageCount">总页数(0)</button></td>
            <td><button type='button' class="pageCurrent">当前页(1)</button></td>
        </tr>
    </table>
</div>
</body>
<script type="text/javascript" src="http://code.jquery.com/jquery-2.1.1.min.js"></script>
<script type="text/javascript" src="common/page.js"></script>
<script>
$(function(){
    doSearch();//初始页面加载
    doState();//初始查询框使用状态
	$("#checkAllId").change(doChangeTBodyCheckBoxState);
	$("#tbodyId").on("change", ".excelId", doChangeTHeadCheckBoxState);
    //事件注册(表示在#pageId对应容器对象的相关子元素中注册click事件)
    $("#pageId").on("click", ".first,.pre,.next,.last", doJumpToPage);
});
//判断查询框使用状态操作
function doState(){
	if($('input[name="isSearchMax"]:checked').prop("checked")){
		$("#excelDate").attr("disabled","disabled");
	}else{
		$("#excelDate").removeAttr("disabled");
	}
}
//点击Excel名称跳转Excel内容操作
function getExcelAll(excelId){
	window.location.href="doExcelAllUI.do?excelId=" + excelId;
}
//上传文件和上传前检验操作--------------------------------------------------------------------
//上传文件提交
function doSubmit(){
	if(checkout()){
		var formData = new FormData($("#form_table")[0]);
		$.ajax({
	        url : "doImport.do",//这里写你的url
	        type : "post",//post请求
	        dataType : "json",//这里是返回类型，一般是json,text等
	        cache : false,//上传文件不需要缓存
	        data : formData,
	        //告诉jQuery不要去设置Content-Type请求头
	        contentType : false,//当有文件要上传时,此项是必须的,否则后台无法识别文件流的起始位置
	        //告诉jQuery不要去处理发送的数据
	        processData : false,//是否序列化data属性,默认true(注意：false时type必须是post)               
	        clearForm : true,//提交后是否清空表单数据
	        async : false,//用于控制同步和异步，默认是true
	        success : function(result){
	        	if(result.state == 1){
	        		window.location.href = "doIndexUI.do";//重新页面加载
				}else{
					alert(result.message);
				}
	        }
	    });
	}
}
//上传文件上传前校验
function checkout(){
	var file = $("#file_excel").val();
	if(file == ""){
		alert("请选择文件！！！");
		return false;
	}else{
		return true;
	}
}
//单选和复选框操作--------------------------------------------------------------
//单选操作
function doChangeTHeadCheckBoxState(){
	var flag = true;
	$("#tbodyId input[type='checkbox']").each(function() {
		flag = flag && $(this).prop("checked");
	});
	$("#checkAllId").prop("checked", flag);
}
//全选操作
function doChangeTBodyCheckBoxState(){
	var state = $(this).prop("checked");
	$("#tbodyId input[type='checkbox']").prop("checked", state);
}
//获取去选项删除操作------------------------------------------------------------
//删除操作
function doDelete() {
	var ids = doGetCheckedIds();
	if(ids.length == 0) {
		alert("请先选择");
		return;
	}
	if(window.confirm("您确定要删除吗？")){
		var params = {"excelIds" : ids.toString()}
		var url = "doDeleteExcel.do";
		$.post(url, params, function(result){
			if(result.state == 1){
				window.location.href="doIndexUI.do";//重新页面加载
			}else{
				alert(result.message);
			}
		});
	}
}
//获取用户选中的记录id
function doGetCheckedIds(){
	var array = [];
	$("#tbodyId input[type='checkbox").each(function(){
		if($(this).prop("checked")){
			array.push($(this).val());
		}
	});
	return array;
}
//查询后返回数据页面显示操作----------------------------------------------------------------
//起始查询和查询框操作
function doSearch(){
    var pageCurrent = $("#pageId").data("pageCurrent");
    if(!pageCurrent)pageCurrent = 1;
	var params = {//传输的参数集
	    "pageCurrent" : pageCurrent,
		"excelDate" : $("#excelDate").val(), //获取input的id值，excelDate
		"excelName" : $("#excelName").val(), //获取input的id值，excelName
		//获取选项状态,true返回1、false返回0
		"isSearchMax" : $('input[name="isSearchMax"]:checked').prop("checked") ? 1:0
	}
	var url = "doSearch.do"; //接口
	$.post(url, params, function(result){//post请求
		doHandleResponseResult(result);//处理响应结果
        doSetPagination(result.data);
		autoAddNumber();//自增序列号
        doChangeTHeadCheckBoxState();//重置复选框
	});
}
function doHandleResponseResult(result){
	if(result.state == 1){
		doSetTableBobyRows(result.data.records);
	}else{
		alert(result.message);
	}
}
function doSetTableBobyRows(data){
	var tbodyId = $("#tbodyId");
	tbodyId.empty();
	for(var i in data){
		var tr = $("<tr></tr>");
		tr.data("rowData", data[i]);
		var tds = doCreateIds(data[i]);
		tr.append(tds);
		tbodyId.append(tr);
	}
}
function doCreateIds(row){
	var tds = 
		"<td><input type='checkbox' name='excelId' class='excelId' value='"+row.excelId+"'><span class='numberClass'></span></td>"+
        "<td align='center'><input type='submit' value='"+row.excelName+"' onclick='getExcelAll("+row.excelId+")'></td>"+
        "<td align='center'>"+row.type+"</td>"+
        "<td align='center'>"+row.excelDate+"</td>"+
        "<td align='center'>"+row.week+"</td>"+
        "<td align='center'>"+row.createDate+"</td>"+
        "<td align='center'>"+row.occupancyRate+"%</td>";
	return tds;
}
//页面数据自增序列号
function autoAddNumber(){
    var pageCurrent = $("#pageId").data("pageCurrent");
    var pageNo = (pageCurrent-1) * 20;
	function number(){
	    var index = 0;
    	for(var i = 0+pageNo;i < $(".numberClass").length+pageNo; i++){
			$(".numberClass").get(index).innerHTML = i+1;
			index++;
    	}
	}
	number();
}
</script>
</html>