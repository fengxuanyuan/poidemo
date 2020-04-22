<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%
    String path = request.getContextPath();
    String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
    <title>POI测试</title>
</head>

<body>
<a href="poitest.jspx?_m=poi_down">下载</a><br>
<form action="poitest.jspx?_m=poi_upload" method="post" enctype="multipart/form-data">
    <input type="file" name="file">
    <input type="submit" value="submit">
</form>
</body>

<script>
    function saveUpload() {
        alert("hja");
        var upExcel = $("#upExcel").val();//这个是个很奇怪的值，但是可以获取想要上传的文件结尾
        if(upExcel == ''){
            alert("请选择excel,再上传");
        }else if(upExcel.lastIndexOf(".xls")<0){//可判断以.xls和.xlsx结尾的excel
            alert("只能上传Excel文件");
        }else {
            var formData = new FormData($("#excel")[0]);//表单id
            var xhLxComb= $("#xhLxComb").combobox('getValue');
            formData.append("xhLxComb",xhLxComb);//参数
            //formData.append("file",document.getElementById("excel"));

            $.ajax({
                url: basePath + "/dataCollectServlet.do?action=uploadExcelSO2",
                type: 'POST',
                data: formData,
                async: false,
                cache: false,
                contentType: false,
                processData: false,
                success: function (data) {
                    if(data == "1"){
                        $('#uploadDlg').dialog('close');//关闭补采窗口
                        $.messager.show({  //这里其实就在在屏幕的右下角显示一个提示框
                            title: '提示',
                            msg:  '补采成功'
                        })
                    }else if(data == "0"){
                        $('#uploadDlg').dialog('close');//关闭补采窗口
                        $.messager.show({  //这里其实就在在屏幕的右下角显示一个提示框
                            title: '提示',
                            msg:  '补采失败,数据可能存在问题'
                        })
                    }
                    $("#dg").datagrid("reload"); //重新加载数据，即:刷新页面
                }
            });
        }
    }
</script>
</html>
