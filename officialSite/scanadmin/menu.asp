<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link href="css.css" rel="stylesheet" type="text/css">
<SCRIPT lanuage="Javascript">
var tt='start';
var ii='start';
function turnit(ss,bb) {

  if (ss.style.display=="none") {
    if(tt!='start') tt.style.display="none";
    if(ii!='start') ii.src="";
    ss.style.display="";
    tt=ss;
    ii=bb;
    bb.src="";
  }
  else {
    ss.style.display="none"; 
    bb.src="";
  }
}

function openWindow(url) {
  popupWin = window.open(url, 'new_page', 'width=400,height=410,scrollbars')
}
</SCRIPT>
</head>

<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#999999">
<tr>
  <td valign=top> 
    <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#999999">
      <tr>
        <td> 
          <div align="center">
            <table cellpadding=0 cellspacing=0 width=180 bgcolor="#000000">
              <tr> 
                <td height=25 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title';  > 
                  <div align="center"><font color="#FFFFFF"><strong>控制面版/<span><a href="index.asp" target=_top><b><span><font color="#FFFFFF">操作指南</font></span></b></a></span></strong></font></div></td>
              </tr>
            </table>
            <br>
            <table  border="1" cellspacing="0" cellpadding="0"  width='180' align="center" bgcolor="#000000" bordercolor="999999">
              <tbody>
			  <tr> 
                  <td height="22" align="center" valign=middle id=tag1     
    style='CURSOR: hand' onClick=turnit(Content1,tag1); language=JScript><font color="#FFFFFF">SCAN ASIA产品管理</font></td>
                </tr>
                <tr> 
                  <td align="right" valign=top id=Content1 style='DISPLAY: none'> 
                    <table width="100%">
                      <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=cb/fenlei1.asp target=forum>项目管理/添加</a></td>
                      </tr>
                      <!--tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=cb/TVPfen-2.asp target=forum>多图片项目添加</a></td>
                      </tr--><tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=cb/TVPphoto-2.asp target=forum>产品添加</a></td>
                      </tr>
					  <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=cb/TVPphoto-1.asp target=forum>产品管理</a></td>
                      </tr>
					    
                    </table></td>
                </tr>
			   <!--tr> 
                  <td height="22" align="center" valign=middle id=tag5     
    style='CURSOR: hand' onClick=turnit(Content5,tag5); language=JScript><font color="#FFFFFF">一图片产品</font></td>
                </tr>
                <tr> 
                  <td align="right" valign=top id=Content5 style='DISPLAY: none'> 
                    <table width="100%">
                      <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=photo/TVPfen-1.asp target=forum>分类管理</a></td>
                      </tr>
                      <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=photo/TVPfen-2.asp target=forum>分类添加</a></td>
                      </tr>
					  <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=photo/TVPphoto-1.asp target=forum>产品管理</a></td>
                      </tr>
					    <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=photo/TVPphoto-2.asp target=forum>产品添加</a></td>
                      </tr>
                    </table></td>
                </tr-->
                <tr> 
                  <td height="22" align="center" valign=middle id=tag2     
    style='CURSOR: hand' onClick=turnit(Content2,tag2); language=JScript><font color="#FFFFFF">新闻公告管理</font></td>
                </tr>
                <tr> 
                  <td align="right" valign=top id=Content2 style='DISPLAY: none'> 
                    <table width="100%">
					  <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=news/newsadmin.asp target=forum>新闻管理</a></td>
                      </tr>
					    <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=news/newstj.asp target=forum>新闻发布</a></td>
                      </tr>
                    </table></td>
                </tr>
				<!--tr> 
                  <td height="22" align="center" valign=middle  id=tag16    
    style='CURSOR: hand' onClick=turnit(Content16,tag16); language=JScript><font color="#FFFFFF">客户反馈管理</font></td>
                </tr>
                <tr> 
                  <td align="right" valign=top id=Content16 style='DISPLAY: none'><table width="100%">
                      <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=gbook1.asp?action=manage target=forum>客户反馈</a></td>
                      </tr>
                    </table></td>
                </tr-->
                <tr> 
                  <td height="22" align="center" valign=middle  id=tag13     
    style='CURSOR: hand' onClick=turnit(Content13,tag13); language=JScript><font color="#FFFFFF">管理员设置</font></td>
                </tr>
                <tr> 
                  <td align="right" valign=top id=Content13 style='DISPLAY: none'><table width="100%">
                      <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href=xgmima.asp target=forum>密码修改</a></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="22" align="center" valign=middle  id=tag14     
    style='CURSOR: hand' onClick=turnit(Content14,tag14); language=JScript><font color="#FFFFFF">系统管理</font></td>
                </tr>
                <tr> 
                  <td align="right" valign=top id=Content14 style='DISPLAY: none'><table width="100%">
                      <tr> 
                        <td height="22" align="center" valign="middle" bgcolor="#FFFFFF"><a href="logout.asp" target="_top">注销登陆</a></td>
                      </tr>
                    </table></td>
                </tr>
              </tbody>
            </table>
          </div>
          <br>
          <div align="center"><br>
          </div>
          
        </td>
      </tr>
<A NAME="StranLink"></A>
   </TABLE><SCRIPT LANGUAGE=Javascript SRC="../images/big5style.js"></SCRIPT>
  