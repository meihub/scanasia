<%
userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
userip2 = Request.ServerVariables("REMOTE_ADDR")
if userip = ""  then
session("ip")=userip2
else
session("ip")=userip
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>befeeling</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<link href="css.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="23"><%=session("admin")%>��� ����ӭ��¼�����˴ε�¼IPΪ:<%=session("ip")%>;ʱ��Ϊ:<%=now()%>;<a href="logout.asp" target="_top">ע����½</a></td>
  </tr>
<A NAME="StranLink"></A>
   </TABLE><SCRIPT LANGUAGE=Javascript SRC="../images/big5style.js"></SCRIPT>
</body>
</html>
