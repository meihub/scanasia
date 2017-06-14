<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Connections/chao.asp"-->
<%if session("admin")="" then
response.Redirect "../login.asp"
end if
%>
<%
dim ID
ID=request("ClassID")
set rs2=server.CreateObject("adodb.recordset")
sql="select * from cc where ClassID="&ID
rs2.open sql,MM_chao_STRING,1,1                
%>
<%
dim ClassID
ClassID=request("ClassID")
    sql="select * from c where ClassID="& ClassID &" and New="& 1 &" order by id DESC"
              
%>



<script language="javascript">
<!-- 
//图片按比例显示程序
var flag=false; 
function DrawImage(ImgD,Width,Height){ 
var image=new Image(); 
image.src=ImgD.src; 
if(image.width>0 && image.height>0){ 
flag=true; 
if(image.width/image.height>= Width/Height){ 
if(image.width>Width){ 
ImgD.width=Width; 
ImgD.height=(image.height*Width)/image.width; 
}else{ 
ImgD.width=image.width; 
ImgD.height=image.height; 
} 
} 
else{ 
if(image.height>Height){ 
ImgD.height=Height; 
ImgD.width=(image.width*Height)/image.height; 
}else{ 
ImgD.width=image.width; 
ImgD.height=image.height; 
} 
} 
} 
} 
//--> 
</script> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>预览：<%=rs2("ClassName")%></title>
<link href="../css.css" rel="stylesheet" type="text/css">
</head>

<body   bgcolor="f2f2f2">
<table width="1000"  border="0" cellspacing="0" cellpadding="0" height="100%">
  <tr>
    <td width="130">&nbsp;</td>
    <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0" >
      <tr>
        <td height="55">&nbsp;</td>
      </tr>
      <tr>
        <td width="207" align="right"><img src="images/l_1.jpg" width="207" height="530"></td>
      </tr>
    </table></td>
    <td width="434" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="23">&nbsp;</td>
      </tr>
      <tr>
        <td><table width="434"  border="0" cellspacing="0" cellpadding="0" background="images/bb.jpg">
          <tr>
            <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="75">&nbsp;</td>
              </tr>
              <tr>
                <td width="12"><img src="images/l_2.jpg" width="12" height="383"></td>
              </tr>
            </table></td>
            <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="243" height="32" background="images/b_1.jpg">&nbsp;&nbsp;<font color="#FFFFFF" size="+1"><b><%=rs2("ClassName")%></b></font></td>
              </tr>
              <tr>
                <td height="43">&nbsp;</td>
              </tr>
              <tr>
                <td height="3"><img src="images/zt.jpg" width="243" height="2"></td>
              </tr>
              <tr>
                <td height="380" bgcolor="#000000">
				
				<DIV align=left valign=top>
                  <TABLE id=table1 style="BORDER-COLLAPSE: collapse" 
                  cellPadding=0   border=0>
                    <TBODY>
                    <TR>
                      
                      <SCRIPT language=JavaScript>
function scroll(n)
{temp=n;
News.scrollTop=News.scrollTop+temp;
if (temp==0) return;
setTimeout("scroll(temp)",20);
}
</SCRIPT>

                      <TD vAlign=top width="98%">
                        <DIV class=wenbenquyu id=News 
                        style="BORDER-RIGHT: 0px dashed; BORDER-TOP: 0px dashed; OVERFLOW: hidden; BORDER-LEFT: 0px dashed; WIDTH: 220px; BORDER-BOTTOM: 0px dashed; HEIGHT: 347px">
                        <DIV align=left>
                        <TABLE id=table31 style="BORDER-COLLAPSE: collapse" 
                        cellPadding=0 width="100%" border=0 align="center">
                          <TBODY>
                          <TR>
                            <TD class=wenbenquyu>
							<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="100%" align="center" valign="top"><br>
		<%
set rs= Server.CreateObject("adodb.recordset")

rs.PageSize = 100 '这里设定每页显示的记录数
rs.CursorLocation= 3
rs.Open sql,MM_chao_STRING,0,2,1 '这里执行你查询SQL并获得结果记录集
pre = true
last = true
page = trim(Request.QueryString("page"))

if len(page) = 0 then
intpage = 1
pre = false
else
if cint(page) =< 1 then
intpage = 1
pre = false
else
if cint(page) >= rs.PageCount then
intpage = rs.PageCount
last = false
else
intpage = cint(page)
end if
end if
end if
if rs.recordcount=0 then
%>
            <%'''筛选产品并分页%>
            <br />
            <div align="center">暂无产品   ！</div>
            <%
else
rs.AbsolutePage = intpage
for i=1 to rs.pagesize
if  rs.eof or  rs.bof then exit for
%>
            <table width="100" border="0" align="center" cellpadding="0" cellspacing="0" style="float:left;margin:0px 2px 2px 2px">
              <tr>
                <td width="100" height="76" align="center" valign="middle" ><table width="100" border="0" align="center" cellpadding="0" cellspacing="0"  >
                    <tr>
                      <td  width="100" height="76" align="center"><table width="100" border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr><td align="center" class="thumbnail" ><A  href="../../Showproduct2.asp?id=<%=rs("scc")%>" ><img src="../../ProImages/<%=rs("ProImg")%>" onload='DrawImage(this,93,70);' border="0"  style="FILTER: alpha(opacity=100)" ></a></td>
                      </tr></table></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td height="20" align="center" valign="middle" background="images/hb.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td align="center" bgcolor="666666"><font color="#ffffff">
					<% 
			if len(rs("ProName"))>5 then
			response.Write(left(rs("ProName"),5)&"…")
			else
			response.Write(rs("ProName"))
			end if
			%>
			</font></td>
                   
                  </tr>
                </table></td>
              </tr>
            </table>
            <%
rs.movenext
next
%></td>
      </tr>
     
	  
      <%
end if
rs.close
set rs=nothing
%>
    </table>
							
							
							</TD>
                          </TR>
                          </TBODY></TABLE></DIV></DIV></TD>
                      <TD vAlign=top width="2%">
                        <TABLE id=table11 style="BORDER-COLLAPSE: collapse" 
                        cellPadding=0 width="80%" border=0 bgcolor="#333333">
                          <TBODY>
                          <TR>
                            <TD><IMG onmousedown=scroll(-3) 
                              onmouseover=scroll(-3) title=按下鼠标速度会更快 
                              style="CURSOR: hand" onmouseout=scroll(0) 
                              src="../../images/up.gif" 
                            border=0></TD></TR>
                          <TR>
                            <TD height=345>　</TD></TR>
                          <TR>
                            <TD><IMG onmousedown=scroll(3) 
                              onmouseover=scroll(3) title=按下鼠标速度会更快 
                              style="CURSOR: hand" onmouseout=scroll(0) 
                              src="../../images/down.gif" 
                              border=0></TD></TR></TBODY></TABLE></TD>
                      </TR>
                    <TR>
                     
                    </TR></TBODY></TABLE></DIV>
				
				</td>
              </tr>
            </table></td>
            <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td height="75">&nbsp;</td>
              </tr>
              <tr>
                <td width="12"><img src="images/r_2.jpg" width="12" height="383"></td>
              </tr>
            </table></td>
            <td valign="top" width="168"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/pz.jpg" width="168" height="75"></td>
              </tr>
              <tr>
                <td><img src="images/logo.jpg" width="168" height="128"></td>
              </tr>
              <tr>
                <td align="center"><img src="images/menu.jpg" width="121" height="235" border="0" usemap="#Map"></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
      <tr>
        <td width="434" height="47"><img src="images/f.jpg" width="434" height="47"></td>
      </tr>
    </table></td>
    <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="55">&nbsp;</td>
      </tr>
      <tr>
        <td width="221"><img src="images/r_1.jpg" width="221" height="530"></td>
      </tr>
    </table></td>
    <td >&nbsp;</td>
  </tr>
</table>
<map name="Map">
  <area shape="rect" coords="7,43,121,70" href="yl.asp?ClassID=1">
  <area shape="rect" coords="7,70,121,96" href="yl2.asp?ClassID=2">
  <area shape="rect" coords="6,97,121,120" href="yl2.asp?ClassID=3">
  <area shape="rect" coords="6,121,119,147" href="yl3.asp?ClassID=33">
  <area shape="rect" coords="5,148,118,175" href="yl2.asp?ClassID=5">
  <area shape="rect" coords="5,176,116,204" href="yl3.asp?ClassID=31">
  <area shape="rect" coords="6,204,117,230" href="yl3.asp?ClassID=28">
</map>
</body>
</html>
