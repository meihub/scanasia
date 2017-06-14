<!--#include file="check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>修改密碼</title>
<script language=javascript>
function CheckForm()
{
   
   if (document.form1.pwd1.value=="")
  {
    alert(" 密码不能为空！");
	document.form1.pwd1.focus();
	return false;
  }
  
    if (document.form1.pwd2.value=="")
  {
    alert(" 确认密码不能为空！");
	document.form1.pwd2.focus();
	return false;
  }
 
    if (document.form1.pwd1.value!=document.form1.pwd2.value)
  {
    alert("两次密码不一样！");
	document.form1.pwd1.focus();
	return false;
  }
  return true;  
}
</script>
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {color: #FF0000}
-->
</style>
<link href="images/style.css" rel="stylesheet" type="text/css">
</head>

<body><center>
<form name="form1" method="post" action="saveadmin.asp" onSubmit="return CheckForm();">
  <table width="35%"  border="0" cellpadding="3" cellspacing="1" bordercolor="#FFFFFF" bgcolor="#CCCCCC">
    <tr bgcolor="#CCCCCC">
      <td width="45%" height="25" bgcolor="#FFFFFF"><div align="center"><%=session("admin")%></div></td>
      <td width="55%" bgcolor="#FFFFFF"><div align="center"><span class="style2">只修改用戶名的，密碼請放空</span></div></td>
    </tr>
	  <tr>
     <td width="45%" height="25" align="center" bgcolor="#FFFFFF"><span class="style1">用戶名：</span></td>
      <td bgcolor="#FFFFFF"><div align="center">
        <input name="username" type="username" class="input" style="width:180; height:23" >
      </div></td>
    </tr>
    <tr>
     <td width="45%" height="25" align="center" bgcolor="#FFFFFF"><span class="style1">新密碼：</span></td>
      <td bgcolor="#FFFFFF"><div align="center">
        <input name="pwd1" type="password" class="input" style="width:180; height:23" size="14"  >
      </div></td>
    </tr>
    <tr>
      <td height="25" align="center" bgcolor="#FFFFFF"><span class="style1">確認新密碼：</span></td>
      <td bgcolor="#FFFFFF"><div align="center">
        <input name="pwd2" type="password" class="input" style="width:180; height:23 " size="14">
      </div></td>
    </tr>
    <tr>
      <td width="45%" height="25" align="center" bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF"><input type="submit" value="確定">&nbsp;<input type="reset" value="取消"></td>
    </tr>
<A NAME="StranLink"></A>
   </TABLE><SCRIPT LANGUAGE=Javascript SRC="../images/big5style.js"></SCRIPT>
</form></center>
</body>
</html>
