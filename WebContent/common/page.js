//初始化分页信息
function doSetPagination(pageObject) {
    //初始化页面分页数据
    $(".rowCount").html("总记录数(" + pageObject.rowCount + ")");
    $(".pageCount").html("总页数(" + pageObject.pageCount + ")");
    $(".pageCurrent").html("当前页(" + pageObject.pageCurrent + ")");
    //绑定数据(将相关数据保存到某个对象上)
    //data(key,[value]) 用于在某个对象上绑定或获取值
    $("#pageId").data("pageCurrent", pageObject.pageCurrent);
    $("#pageId").data("pageCount", pageObject.pageCount);
}

//基于pageCurrent进行下一步的查询
function doJumpToPage() {
    console.log("doJumpToPage()");
    //1.获取点击对象class属性,基于属性值判定对象
    //基于prop方法获取class属性值
    var cls = $(this).prop("class");
    //2.基于Class属性值的不同修改PageCurrent的值
    //2.1获取pageCurrent的值
    var pageCurrent = $("#pageId").data("pageCurrent");
    //2.2获取pageCount的值
    var pageCount = $("#pageId").data("pageCount");
    //2.3修改pageCurrent的值
    //debugger
    if (cls == "first") {
        pageCurrent = 1;
    } else if (cls == "next" && pageCurrent < pageCount) {
        pageCurrent++;
    } else if (cls == "pre" && pageCurrent > 1) {
        pageCurrent--;
    } else if (cls == "last") {
        pageCurrent = pageCount;
    }
    //3.基于PageCurrent新的值进行当前页查询
    //3.1 重新绑定pageCurrent值
    $("#pageId").data("pageCurrent", pageCurrent);
    //3.2基于PageCurrent执行分页查询
    doSearch();
}