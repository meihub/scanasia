<!--#include file="check.asp"-->
<HTML>
<HEAD>
<TITLE>SCAN ASIA后台管理系统</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<STYLE>.navPoint {
	COLOR:#ffffff; CURSOR: hand; FONT-FAMILY: Webdings; FONT-SIZE: 9pt
}
</STYLE>
<SCRIPT>
function switchSysBar(){
	if (switchPoint.innerText==7){
		switchPoint.innerText=8
		document.all("frmTitle").style.display="none"
	}
	else{
		switchPoint.innerText=7
		document.all("frmTitle").style.display=""
	}
}
</SCRIPT>
</HEAD>
<BODY scroll=no style="MARGIN: 0px">
<!--#include file="top2.asp"-->
<TABLE border=0 cellPadding=0 cellSpacing=0 height="100%" width="100%">
  <TBODY>
  <TR>
    <TD align=middle id=frmTitle noWrap vAlign=center name="frmTitle"><IFRAME 
      frameBorder=0 id=BoardTitle name=BoardMenu scrolling=no 
      src="menu.asp" 
      style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 180px; Z-INDEX: 2"></IFRAME>
    <TD bgColor=#999999 onclick=switchSysBar() style="WIDTH: 10pt"><SPAN class=navPoint id=switchPoint title=关闭/打开菜单>7</SPAN>    </TD>
    <TD style="WIDTH: 100%"><IFRAME  scrolling="NO" frameBorder=0 border="0" framespacing="0"  id=top name=top 
      src="top.asp" 
      style="HEIGHT: 29; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"></IFRAME><IFRAME frameBorder=0 id=frmright name=forum 
      src="admin.asp" 
      style="HEIGHT: 100%; VISIBILITY: inherit; WIDTH: 100%; Z-INDEX: 1"></IFRAME>
      </TD>
    </TR>
    </TBODY><A NAME="StranLink"></A>
   </TABLE><SCRIPT LANGUAGE=Javascript SRC="../images/big5style.js"></SCRIPT>
  </BODY>
</HTML>