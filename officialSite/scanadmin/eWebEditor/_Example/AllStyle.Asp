<!--#include file = "../Asp/Config.Asp"-->
<%Dim l_Style, l_Skin, read
l_Skin = request.QueryString("Skin")
l_Style = Request.QueryString("Style")
Read = "1|7|15"
Dim i,nStyle,n,xStyle,showNum,ShowWidth,ShowHeigh
IF l_Style = "" Then
 l_Style = "standard650"
End IF
IF l_Skin = "" Then
 l_Skin = "office2003"
End IF
showNum = False
For i=1 to Ubound(aStyle)
nStyle = Split(aStyle(i),"|||")
  for n=0 to Ubound(nstyle)
    IF Lcase(nStyle(0)) = Lcase(l_Style) Then
	  ShowWidth = nStyle(4)
	  ShowHeigh = nStyle(5)
      showNum = True
	  Exit for
	End if
  next
IF showNum = True then
Exit for
End IF
next
%>
<HTML>
<HEAD>
<TITLE>eWebEditor �� ����ʵ��</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link rel='stylesheet' type='text/css' href='example.css'>
<SCRIPT language=javascript>
var s_Style = "<%=l_Style%>";
var s_Skin = "<%=l_Skin%>";

function doClickStyle(obj){
	s_Style = obj.value;
	doGoUrl();
}

function doClickSkin(obj){
	s_Skin = obj.value;
	doGoUrl();
}

function doGoUrl(){
	location.href = location.pathname + "?style=" + s_Style + "&skin=" + s_Skin;
}
</SCRIPT>
</HEAD>
<BODY>
<p><b>���� �� <a href="default.asp">ʾ����ҳ</a> &gt; ϵͳ�Դ�ȫ����ʽ��ʾ</b></p>
<p> ������ͨ����ͬ�������ʵ�ֲ�ͬ��Ч����</p>
<p> ע�⣺���µ���ʽ�Ǻͺ�̨ͬ���ģ������ں�����һ������ʽʱ������Ҳ����ʾ��</p>
<TABLE cellSpacing=2 cellPadding=0 WIDTH="700px" border=0 style=" border:1px solid #CCCCCC;">
  <tr>
    <td style="background:#EEEEEE; padding:2px 2px 2px 10px; font-weight:bold; height:30px; line-height:30px; " colspan="5">ϵͳȫ����ʽ������</td>
  </tr>
  <TR style="height:25px; line-height:25px;">
    <%Dim y,ShowListNum,arrStyle,arrRead,x
ShowListNum = 5
arrRead = Split(Read,"|")
For y=1 to Ubound(aStyle)
arrStyle = Split(aStyle(y),"|||")
Response.Write "<TD style='border-bottom:1px solid #EEEEEE; padding-left:10px;'>" & VBCrlf
Response.Write "<INPUT id=d_style_1 onclick=doClickStyle(this) type=radio value="&arrStyle(0)
IF Lcase(l_Style)=Lcase(arrStyle(0)) then 
response.Write " CHECKED"
End IF
response.Write "><LABEL for=d_style_1>"&arrStyle(0)
  For x=0 to ubound(arrRead)
    IF Lcase(Trim(y)) = Lcase(Trim(arrRead(x))) then
    response.Write "(�Ƽ�)"
    End if
   next
response.write "</LABEL></TD>"
if y mod showlistnum = 0 Then
Response.Write "</tr><tr>"
end if
next
%>
  </TR>
</TABLE>
<TABLE cellSpacing=2 cellPadding=0 WIDTH="700px;" border=0 style=" border:1px solid #CCCCCC; margin-top:10px;">
  <tr>
    <td style="background:#EEEEEE; padding:2px 2px 2px 10px; font-weight:bold; height:30px; line-height:30px; " colspan="5">ϵͳ�Դ�Ƥ��������</td>
  </tr>
  <TR>
    <TD WIDTH="20%"><INPUT id=d_skin_1 
                        onclick=doClickSkin(this) type=radio value=blue1 <%IF Lcase(l_Skin)=Lcase("blue1") then response.Write "CHECKED"%>>
      <LABEL 
                        for=d_skin_1> blue</LABEL></TD>
    <TD WIDTH="20%"><INPUT id=d_skin_2 
                        onclick=doClickSkin(this) type=radio value=blue2 <%IF Lcase(l_Skin)=Lcase("blue2") then response.Write "CHECKED"%>>
      <LABEL 
                        for=d_skin_2> blue1</LABEL></TD>
    <TD WIDTH="20%"><INPUT id=d_skin_3 
                        onclick=doClickSkin(this) type=radio value=green1 <%IF Lcase(l_Skin)=Lcase("green1") then response.Write "CHECKED"%>>
      <LABEL 
                        for=d_skin_3> green</LABEL></TD>
    <TD WIDTH="20%"><INPUT id=d_skin_4 
                        onclick=doClickSkin(this) type=radio value=light1 <%IF Lcase(l_Skin)=Lcase("light1") then response.Write "CHECKED"%>>
      <LABEL 
                        for=d_skin_4> light</LABEL></TD>
    <TD WIDTH="20%"><INPUT id=d_skin_5 
                        onclick=doClickSkin(this) type=radio 
                        value=office2000 <%IF Lcase(l_Skin)=Lcase("office2000") then response.Write "CHECKED"%>>
      <LABEL for=d_skin_5> office2000</LABEL></TD>
  </TR>
  <TR>
    <TD WIDTH="20%"><INPUT id=d_skin_6 
                        onclick=doClickSkin(this) type=radio
                        value=office2003 <%IF Lcase(l_Skin)=Lcase("office2003") then response.Write "CHECKED"%>>
      <LABEL for=d_skin_6> office2003(Ĭ��)</LABEL></TD>
    <TD WIDTH="20%"><INPUT id=d_skin_7 
                        onclick=doClickSkin(this) type=radio 
                        value=officexp <%IF Lcase(l_Skin)=Lcase("officexp") then response.Write "CHECKED"%>>
      <LABEL for=d_skin_7>officexp</LABEL></TD>
    <TD WIDTH="20%"><INPUT id=d_skin_8 
                        onclick=doClickSkin(this) type=radio value=red1 <%IF Lcase(l_Skin)=Lcase("red1") then response.Write "CHECKED"%>>
      <LABEL 
                        for=d_skin_8> red</LABEL></TD>
    <TD WIDTH="20%"><INPUT id=d_skin_9 
                        onclick=doClickSkin(this) type=radio value=vista1 <%IF Lcase(l_Skin)=Lcase("vista1") then response.Write "CHECKED"%>>
      <LABEL 
                        for=d_skin_9> vista</LABEL></TD>
    <TD WIDTH="20%"><INPUT id=d_skin_10 
                        onclick=doClickSkin(this) type=radio 
                        value=yellow1 <%IF Lcase(l_Skin)=Lcase("yellow1") then response.Write "CHECKED"%>>
      <LABEL for=d_skin_10> yellow</LABEL></TD>
  </TR>
</TABLE>
<TABLE cellSpacing=2 cellPadding=0 WIDTH="700px;" border=0 style=" border:1px solid #CCCCCC; margin-top:10px;">
  <tr>
    <td style="background:#EEEEEE; padding:2px 2px 2px 10px; font-weight:bold; height:30px; line-height:30px; " colspan="5">��ǰ��ʽ��ʾ��</td>
  </tr>
  <FORM name=myform action=retrieve.asp method=post>
  <%
  Dim IHtml
  iHtml = "<P>������ʹ�� <A href='http://www.ewebeditor.net/'>eWebEditor����HTML�༭�� V6.2</A>&nbsp;&nbsp;<A href='http://www.nmyxwl.cn'>��Ѷ����</A>������.</P><P><img src='/EwebEditor/_Example/images/logo.gif' border=0> &nbsp;- ����HTML�༭����HTML���߱༭�ð���<P>"
  iHtml = iHtml & "<strong>��������:</strong>��ϵͳ���հ�Ȩ��EwebEditor����,�����޸�ֻ����Ա�����ʹ�ù����� "
iHtml = iHtml & "<P>��һЩ����������޸ģ�����ֻ���޸Ĳ��ֽ���Ȩ����ʹ�ñ�ϵͳ�����µķ���ЧӦ<P>����û���κι�ϵ!��������;Υ����ewebeditor�ٷ��������������������<P>�����߱������ϸ�����EwebEditor�ٷ���Ȩ������ϣ�����֧�����棡лл��</P>" 
  %>
    <TR>
      <TD style='padding:10px; text-align:center'><INPUT type=hidden name=content1 value="<%=iHtml%>">
        <IFRAME id=eWebEditor1 src="../ewebeditor.htm?id=content1&style=<%=l_style%>&skin=<%=l_skin%>" frameBorder=0 width=<%=ShowWidth%> scrolling=no height=<%=ShowHeigh%>></IFRAME></TD>
    </TR>
    <TR>
      <TD style="padding:10px;text-align:center"><INPUT class=btn type=submit value=�ύ>
        <INPUT class=btn type=reset value=����>
      </TD>
    </TR>
  </FORM>
</TABLE>
<TABLE cellSpacing=0 cellPadding=0 WIDTH="700px" border=0 style=" border:1px dotted #CCCCCC; background:#F7F7F7; margin-top:10px;">
  <tr>
    <td style="background:#EEEEEE; padding:2px 2px 2px 10px; font-weight:bold; height:30px; line-height:30px; " colspan="5">��ʽ����������</td>
  </tr>
    <TR>
      <TD style="padding:10px;">
	  &lt;input type="hidden" name="content1" value=""&gt;<BR>
        &lt;iframe id="eWebEditor1" src="../ewebeditor.html?id=content1&amp;style=<SPAN class=red><%=l_style%></SPAN>&amp;skin=<SPAN class=red><%=l_skin%></SPAN>" frameborder="0" scrolling="no" width="<SPAN class=red><%=ShowWidth%></SPAN>" height="<SPAN class=red><%=ShowHeigh%></SPAN>"&gt;&lt;/iframe&gt; </TD>
    </TR>
</TABLE>