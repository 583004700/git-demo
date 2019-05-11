var d = new Date();
var name = "朱未斌";
var date = d.getFullYear()+"-"+(parseInt(d.getMonth())+1)+"-"+d.getDate();
var nowYear = d.getFullYear();
/**
	邮件操作相关
*/
var mailObj = {
	getMails:function(t){
		var lis = $("[toggle-list="+t+"]").find("li");
		var mailJsons = [];
		for(var i=0;i<lis.length;i++){
			var jtitle = $(lis[i]).find(".subject").html().replace(/\s/g,"");
			var jfrom = $(lis[i]).find(".from").html().replace(/\s/g,"");
			var monthDay = $(lis[i]).find(".desc-time").html().replace(/\s/g,"");
			var obj = [nowYear+"-"+monthDay,name,lis.length-i,jfrom,jtitle,"无需处理",""];
			mailJsons.unshift(obj);
		}
		return mailJsons;
	}
}

/**
	excel操作相关
*/
var ExcelObj = {
	exportFile:function(t){
		var header = ["日期","姓名","序号","发起人","邮件主题","处理状态","备注"];
		var data = mailObj.getMails();
		var FileName = "邮件数据采集_"+name;
        this.JSONToExcelConvertor(header, data,FileName);
	},
	JSONToExcelConvertor:function(ShowLabel,JSONData, FileName){
		var header = [];
		var data = [];
		for(var i=0;i<ShowLabel.length;i++){
			var o = {"value":ShowLabel[i], "type":"ROW_HEADER_HEADER", "datatype":"string"};
			header.push(o);
		}

		for(var j = 0;j<JSONData.length;j++){
			var datan = [];
			for(var k = 0;k<JSONData[j].length;k++){
				var od = {"value":JSONData[j][k], "type":"ROW_HEADER"};
				datan.push(od);
			}
			data.push(datan);
		}

        //先转化json
        var arrData = typeof data != 'object' ? JSON.parse(data) : data;

        var excel = '<table>';

        //设置表头
        var row = "<tr>";
        for (var i = 0, l = header.length; i < l; i++) {
            row += "<td>" + header[i].value + '</td>';
        }


        //换行
        excel += row + "</tr>";

        //设置数据
        for (var i = 0; i < arrData.length; i++) {
            var row = "<tr>";

            for (var index in arrData[i]) {
                var value = arrData[i][index].value === "." ? "" : arrData[i][index].value;
                row += '<td>' + value + '</td>';
            }

            excel += row + "</tr>";
        }

        excel += "</table>";

        var excelFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>";
        excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8">';
        excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-excel';
        excelFile += '; charset=UTF-8">';
        excelFile += "<head>";
        excelFile += "<!--[if gte mso 9]>";
        excelFile += "<xml>";
        excelFile += "<x:ExcelWorkbook>";
        excelFile += "<x:ExcelWorksheets>";
        excelFile += "<x:ExcelWorksheet>";
        excelFile += "<x:Name>";
        excelFile += "{worksheet}";
        excelFile += "</x:Name>";
        excelFile += "<x:WorksheetOptions>";
        excelFile += "<x:DisplayGridlines/>";
        excelFile += "</x:WorksheetOptions>";
        excelFile += "</x:ExcelWorksheet>";
        excelFile += "</x:ExcelWorksheets>";
        excelFile += "</x:ExcelWorkbook>";
        excelFile += "</xml>";
        excelFile += "<![endif]-->";
        excelFile += "</head>";
        excelFile += "<body>";
        excelFile += excel;
        excelFile += "</body>";
        excelFile += "</html>";


        var uri = 'data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(excelFile);

        var link = document.createElement("a");
        link.href = uri;

        link.style = "visibility:hidden";
        link.download = FileName + ".xls";

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
}
//ExcelObj.exportFile('earlier');
//ExcelObj.exportFile('today');
