<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Redirect "login.asp"
end if
%>
<html>
<head>
<title>SCAN ASIA綴怢奪燴炵苀</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #FF0000}
.style2 {color: #CCCCCC}
-->
</style>
</head>
<body>
<table width="91%" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr> 
    <td height="20"> 
      <div align="center"><font color="#000000">SCAN ASIA後臺管理</font></div>
    </td>
  </tr>
  <tr> 
    <td height="100" valign="top"><br>
      <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td>
            <div align="center"><font color="#FF0000">歡迎管理員對本後台系統進行操作！！<br>
            </font></div>
          </td>
        </tr>
      </table>
      <table width="90%"  border="0" cellspacing="0" cellpadding="0"  align="center">
        <tr>
          <td><p>系統操作指南：<br>
              <span class="style2">本系統爲SCAN ASIA v1.0，主要由圖?展示系統，信息發布系統以及管理員系統組成， 詳況如下：</span><br>
            <span class="go-wenbenkuang"><strong>+ SCAN ASIA's產品管理系統</strong></span><br>
此欄目爲SCAN ASIA的産品展示系統，主要有以下三個大欄目：<br>
<span class="go-wenbenkuang">- 類別管理/添加</span><br>
<span class="style1">此欄目爲産品的類别添加，首先添加完大類别，之後在對準需要的大類後面添加相應的小類别<br>
</span><span class="go-wenbenkuang">- 產品添加</span><br>
<span class="style1">此欄目産品圖片上傳，上傳時要對應圖片大小和注意選擇好大小類别</span><br>
              <span class="go-wenbenkuang">- 產品管理</span><br>
  此欄目爲對已經上傳的産品圖片進行管理，主要有修改、删除等功能<br>
  </p>
            <p><span class="go-wenbenkuang"><strong>+ 新聞公告管理</strong></span><br>
  此欄目爲新聞公告的發布、修改和删除<br>
  <span class="go-wenbenkuang">- 新聞管理</span><br>
  此欄目爲對已經發布的新聞進行修改、删除等功能管理<br>
  <span class="go-wenbenkuang">- 新聞發布</span><br>
  此欄目爲新聞的發布</p>
            <p><span class="go-wenbenkuang"><strong>+ 管理員設置</strong></span><br>
  此欄目就是對後台管理員帳号密碼的修改</p>
<p><span class="go-wenbenkuang"><strong>+ 系統管理</strong></span><br>
  注銷此次登陸</p></td>
        </tr>
      </table>      <br>
      <table border="0" align="center" cellpadding="0" cellspacing="0" width="95%">
        <tr> 
          <td>            <div align="center"><br>
              </div>
          </td>
        </tr>
        <%								
function liuyan()
dim tmprs
last_logintime
tmprs=conn.execute("Select Count(guest_id) from guest Where datediff(n,now,guest_time) > 0 and datediff(n,guest_time,last_logintime) > 0")
liuyan=tmprs(0)
set tmprs=nothing
if isnull(liuyan) then liuyan=0
end function
%>
        <tr> 
          <td> 
            <div align="center">  
              <p>上次登陸時間：<%=session("last_logintime")%></p>
              <p>&nbsp;</p>
              <p><br>
              </p>
            </div>
          </td>
        </tr>
    </table>    </td>
  </tr><A NAME="StranLink"></A>
</table><SCRIPT LANGUAGE=Javascript SRC="../images/big5style.js"></SCRIPT>
<br>
</body>
</html>
