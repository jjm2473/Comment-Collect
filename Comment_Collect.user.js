// ==UserScript==
// @name         Comment Collect
// @namespace    https://openuserjs.org/users/jjm2473/
// @version      0.7
// @encoding     utf-8
// @description  no public usage
// @author       jjm2473
// @match        http://admin2.mdl.com/admin/page/content/comment
// @grant        GM_xmlhttpRequest
// @run-at       document-end
// ==/UserScript==
(
function(){ 
    var div=document.getElementsByClassName("form-inline")[0];
    if(div === null)return;
    
    var tidv=div.children[2].children[0];
    
    div.appendChild(document.createElement("br"));
    var datetimet=new Date();
    
    var offset=datetimet.getTimezoneOffset()*60*1000;
    
    var dateev=parseInt(datetimet.getTime()/1000)*1000-offset;//+timezone
    var datesv=parseInt(dateev/(24*60*60*1000))*(24*60*60*1000);
    
    var dates=document.createElement("input");
    dates.type="datetime-local";
    dates.valueAsNumber=datesv;
    
    var datee=document.createElement("input");
    datee.type="datetime-local";
    datee.valueAsNumber=dateev;
    
    div.appendChild(dates);
    div.appendChild(datee);
    
    var aa=document.createElement("button");
    aa.classList.add("btn");
    aa.classList.add("btn-primary");
    aa.innerText="导出";
    div.appendChild(aa);

    var script=document.createElement("script");
    script.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.core.min.js";
    div.appendChild(script);

    /** begin xlsx code */
    var datenum =function (v, date1904) {
        if(date1904) v+=1462;
        var epoch = Date.parse(v);
        return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    };

    var sheet_from_array_of_arrays=function (data, opts) {
        var ws = {};
        var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
        for(var R = 0; R != data.length; ++R) {
            for(var C = 0; C != data[R].length; ++C) {
                if(range.s.r > R) range.s.r = R;
                if(range.s.c > C) range.s.c = C;
                if(range.e.r < R) range.e.r = R;
                if(range.e.c < C) range.e.c = C;
                var cell = {v: data[R][C] };
                if(cell.v === null) continue;
                var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

                if(typeof cell.v === 'number') cell.t = 'n';
                else if(typeof cell.v === 'boolean') cell.t = 'b';
                else if(cell.v instanceof Date) {
                    cell.t = 'n'; cell.z = XLSX.SSF._table[22];
                    cell.v = datenum(cell.v);
                }
                else cell.t = 's';

                ws[cell_ref] = cell;
            }
        }
        if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
        return ws;
    };

    var Workbook=function () {
        if(!(this instanceof Workbook)) return new Workbook();
        this.SheetNames = [];
        this.Sheets = {};
    };
    
    var s2ab=function(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    };
    /** end xlsx code */
    
    var getComment=function(tid,cb){
        var url = "/api/admin/comment/list?gid="+tid+"&from="+parseInt((dates.valueAsNumber+offset)/1000)+"&to="+parseInt((datee.valueAsNumber+offset)/1000)+"&type=1&draw=1&start=0&size=10000&_="+(new Date().getTime());
        $.getJSON(url, function(data){
            var Sheet1=[["楼层","用户名","用户ID","引用贴标题","引用贴ID","内容","时间"]];
            var ary=data.obj;
            for(var i=0;i<ary.length;++i){
                var cont=ary[i].content;
                var uid=ary[i].commentUser.userid;
                var uname=ary[i].commentUser.nickname;
                var fl=ary[i].floor;
                var time=new Date(ary[i].commentTime*1000);
                var qtid=ary[i].quote? ary[i].quote.gid : null;
                var qtitle=ary[i].quote? ary[i].quote.title : "";
                
                var row=[fl,uname,uid,qtitle,qtid,cont,time];
                Sheet1.push(row);
            }
            cb(Sheet1);
        });
    };

    var toDownload = function ( output, filename) {
		var blob = new Blob([output],
            {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
		var objectURL = URL.createObjectURL( blob );
        var a=document.createElement('a');
        a.href=objectURL;
        a.download=filename||'Untitled';
        a.click();
	};
    aa.addEventListener("click",function(){
        var tid=tidv.value;
        if(!tid){
            if(!confirm("!!No gid, are you sure?"))return;
        }
        getComment(tid,function(Sheet1){
            var wb = new Workbook(), ws = sheet_from_array_of_arrays(Sheet1);
            /* add worksheet to workbook */
            wb.SheetNames.push("Sheet1");
            wb.Sheets["Sheet1"] = ws;
            var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:true, type: 'binary'});
            toDownload(s2ab(wbout),"Comments_"+tid+".xlsx")
        });
    });
    console.error("jjm2473,I am here!");//just for debug
})();
