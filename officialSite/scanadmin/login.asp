<%'禁止网页缓存
Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache"
'验证码生成
dim yz
randomize timer
yz=Int((8999)*Rnd +1009)
session("ok")=yz
%>
<html>
<head>
<title>管理員登陸</title>
<script language="JavaScript">
<!--
if (self != top) top.location.href = window.location.href
//-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script LANGUAGE="javascript">
<!--
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function check()
{
  if(checkspace(document.admininfo.admin.value)) {
	document.admininfo.admin.focus();
    alert("管理员不能为空！");
	return false;
  }
  if(checkspace(document.admininfo.password.value)) {
	document.admininfo.password.focus();
    alert("密码不能为空！");
	return false;
  }
    if(checkspace(document.admininfo.yzm.value)) {
	document.admininfo.yzm.focus();
    alert("验证码不能为空！");
	return false;
  }
    if(document.admininfo.yzm.value != document.admininfo.yzm1.value) {
	document.admininfo.yzm.focus();
	document.admininfo.yzm.value = '';
	document.admininfo.yzm1.value = '';
    alert("验证码不同，请重新输入！");
	return false;
  }
	document.admininfo.submit();
  }
//-->
</script> 
<link href="css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #666666}
.style2 {color: #FF0000}
-->
</style>
</head>
<body  bgcolor="EFEFEF" bottommargin="0" topmargin="0" rightmargin="0"  leftmargin="0">
<form name="admininfo" method="post" action="chkadmin.asp" >
  <div align="center"> 
    <div align="center"> 
	<table border="0" align="center" cellpadding="0" cellspacing="0" height="100%" bgcolor="#444444" width="100%" valign="top"><tr><td valign="top">
      <table border="0" align="center" cellpadding="0" cellspacing="0" height="580" width="100%" >
        <tr> 
          <td align="center"  bgcolor="ECECEC" height="100%"><p>&nbsp;</p>
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="600">&nbsp;</td>
    <td valign="bottom">
	<table width="333" border="0" cellpadding="5" cellspacing="0" >
	<tr align="center"><td height="65" colspan="2"><img src="images/lb.png"></td>
	</tr><tr><td colspan="2" height="1"><img src="../images/hx.png"></td></tr>
              <tr> 
                <td width="30%" height="22" align="right" ><span class="style1">用戶:</span></td>
                <td width="40%" height="22"  >
<input name="admin" class="wenbenkuang" type="text" id="admin" size="16"></td>
              
              </tr>
              <tr> 
                <td height="22" align="right"  ><span class="style1">密碼:</span></td>
                <td height="22"  >
<input name="password" class="wenbenkuang" type="password" id="password" size="16"></td>
               
              </tr>
              <tr> 
                <td height="22" align="right"  ><span class="style2">驗證:</span></td>
                <td height="22" >
<input name="yzm" type="password" class="wenbenkuang" id="yz3" size="8" maxlength="5">
                  <span style="background-color: #FFFFFF"><font color="#ff3333">
                  <%=yz%><input name="yzm1" type="hidden" class="wenbenkuang" id="yzm1" value="<%=yz%>" size="8" maxlength="5">
                  </font></span></td>
              </tr>
              <tr  bgcolor="#e5e5e5"> 
               <td  colspan="2"><input border=0 onClick="return check();" name=Submit32 src="Images/D_Tu_DL.gif" type=image> 
                  <input type="hidden" name="action" value="login"></td>
              </tr>
          </table></td>
  </tr>
</table>

          </td>
        </tr>
        <tr> 
          <td height="124"  bgcolor="#333333" align="center">&nbsp; </td>
        </tr>
        <tr> 
          <td bgcolor="#444444" height="30">&nbsp;</td>
        </tr>
      </table></td></tr><A NAME="StranLink"></A>
   </TABLE><!--SCRIPT LANGUAGE=Javascript SRC="../images/big5style.js"></SCRIPT-->
    </div>
  </div>
</form>
</body>
</html>