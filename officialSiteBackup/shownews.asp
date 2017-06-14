<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="conn.asp" -->
<%
Dim Recordset3__MMColParam
Recordset3__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  Recordset3__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim Recordset3
Dim Recordset3_numRows

Set Recordset3 = Server.CreateObject("ADODB.Recordset")
Recordset3.ActiveConnection = conn
Recordset3.Source = "SELECT * FROM news WHERE ID = " + Replace(Recordset3__MMColParam, "'", "''") + ""
Recordset3.CursorType = 0
Recordset3.CursorLocation = 2
Recordset3.LockType = 1
Recordset3.Open()

Recordset3_numRows = 0
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
Recordset3_first = MM_offset + 1
Recordset3_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (Recordset3_first > MM_rsCount) Then
    Recordset3_first = MM_rsCount
  End If
  If (Recordset3_last > MM_rsCount) Then
    Recordset3_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!-- saved from url=(0027)http://www.gzlinyi.cn/scan/ -->
<HTML xmlns="http://www.w3.org/1999/xhtml"><HEAD><TITLE>SCAN ASIA</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<LINK href="images/css.css" type=text/css rel=stylesheet>
<LINK href="images/loremipsum2.css" type=text/css rel=stylesheet>

<SCRIPT src="images/ilank_flash.js" type=text/javascript></SCRIPT>
<SCRIPT src="images/ajax.js" type=text/javascript></SCRIPT>
<SCRIPT src="images/tooltip.js" type=text/javascript></SCRIPT>
<SCRIPT src="images/unitip.js" type=text/javascript></SCRIPT>
<script language=JavaScript>
////////////////////鼠标移上去图片淡入淡出的效果//////////////////////////////////////
nereidFadeObjects = new Object();
nereidFadeTimers = new Object();

function nereidFade(object, destOp, rate, delta){
if (!document.all)
return
    if (object != "[object]"){  
        setTimeout("nereidFade("+object+","+destOp+","+rate+","+delta+")",0);
        return;
    }
        
    clearTimeout(nereidFadeTimers[object.sourceIndex]);
    
    diff = destOp-object.filters.alpha.opacity;
    direction = 1;
    if (object.filters.alpha.opacity > destOp){
        direction = -1;
    }
    delta=Math.min(direction*diff,delta);
    object.filters.alpha.opacity+=direction*delta;

    if (object.filters.alpha.opacity != destOp){
        nereidFadeObjects[object.sourceIndex]=object;
        nereidFadeTimers[object.sourceIndex]=setTimeout("nereidFade(nereidFadeObjects["+object.sourceIndex+"],"+destOp+","+rate+","+delta+")",rate);
    }
}
/////////////////////////////////////////////////////////////////////////////////////

</script>

<META content="MSHTML 6.00.6000.21015" name=GENERATOR></HEAD>
<BODY>
<DIV id=main2>
<DIV id=menu align="center"></DIV>
<DIV  align="center">
<DIV align="center"	>
<table width="934"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/x_t1.png"></td>
  </tr>
  <tr>
    <td background="images/x_t2.png"  align="center"><table width="97%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
		<table width="100%"  border="0" cellspacing="0" cellpadding="0" height="500">
                       <tr>
                         <td valign="top">
						 
						 <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
      
       
        <tr> 
          <td height="282"></td>
          <td   valign="top"> <% If Not Recordset3.EOF Or Not Recordset1.BOF Then %>
            <table width="89%" height="375" border="0" align="center">
              <tr> 
                <td width="100%" height="22" align="center" class="vip2"><strong><font color="#555555" size="+2"><%=(Recordset3.Fields.Item("new_A1").Value)%></font></strong></td>
              
              </tr>
              <tr> 
                <td height="25" colspan="2" align="center" class="VIP3"><table><tr>
                  <td class="VIP3">Date
				:<%=(Recordset3.Fields.Item("new_A2").Value)%></td><td>&nbsp;</td>
				<td class="VIP3"><span class="VIP3">Author</span>:<%=(Recordset3.Fields.Item("new_A3").Value)%></td></tr></table></td>
              </tr>
			  <tr><td height="1" bgcolor="#999999"></td></tr>
              <tr> 
                <td height="220" colspan="2" valign="top"  align="left" style="padding-top:20px " ><%=(Recordset3.Fields.Item("new_A4").Value)%></td>
              </tr>
            </table>
            <% End If ' end Not Recordset3.EOF Or NOT Recordset3.BOF %> <% If Recordset3.EOF And Recordset3.BOF Then %>
            <table width="75%" border="0" align="center">
              <tr>
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td><div align="center"><font color="#FF0000"><B><FONT COLOR="#FF0000"><STRONG><FONT COLOR="#FF0000" FACE="Geneva, Arial, Helvetica, san-serif">Temporary have no the news!</FONT></STRONG></FONT></B></font></div></td>
              </tr>
            </table>
            <% End If ' end Recordset3.EOF And Recordset3.BOF %> 
<p class="fy"><b>&gt;&gt; </b><a href="news.asp"  >Back</a></p></td>
        </tr>
      </table>
						  </td>
                       </tr>
                     </table>
		</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><img src="images/x_t3.png"></td>
  </tr>
</table>
</DIV>
</DIV>
<DIV id=footer><SPAN class=f1>Copyright @Scan Asia Corporation. 2013. All Rights Reserved </SPAN> 
</DIV></DIV><SCRIPT type=text/javascript>
var objFlash = new ilankFlash("menu/event.swf","","960","157","9","","false","high");
objFlash.addParam("wmode","transparent");
objFlash.write("menu");
</SCRIPT></BODY></HTML>
