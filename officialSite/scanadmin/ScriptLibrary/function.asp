
<%
'------------------------------------------��վ��������--------------------------------------



'**************************************************
'������:deletefile
'����:path ------ �ļ�·��
'����:ͨ����ȡ�ļ�·��ɾ���ļ�
function deletefile(path)
set fso=server.CreateObject("Scripting.FilesystemObject")
if fso.FileExists(Server.MapPath(path)) then
fso.DeleteFile(Server.MapPath(path))
end if
set fso=nothing
end function
'**************************************************




'*************************************************************************************
'������:go\back
'����:rightmsg/error ------ �����Ի�����ȷ/������Ϣ
'     url ------ ת��ҳ��
'����:��ȷ/������Ϣ��ʾ��ҳ��ת��
function go(rightmsg,url)
response.write "<script>alert('"&rightmsg&"');window.location.href='"&url&"';</script>"
response.end()
end function
'---------------------------------
function back(errormsg)
response.write "<script>alert('"&errormsg&"');history.go(-1);</script>"
response.end()
end function
'************************************************************************************


'��ʱ�������ļ����ĺ���*****************
function makefilename(fname)
fname = fname
fname = replace(fname,"-","")
fname = replace(fname," ","") 
fname = replace(fname,":","")
fname = replace(fname,"PM","")
fname = replace(fname,"AM","")
fname = replace(fname,"����","")
fname = replace(fname,"����","")
makefilename=fname & ".html"
end function 
'***************************************


'======================��ʾ�ո��뻻��========================================
Function unHtml(content)
	ON ERROR RESUME NEXT
	unHtml=content
	IF content <> "" Then
		unHtml=Server.HTMLEncode(unHtml)
		unHtml=Replace(unHtml,vbcrlf,"<br>")
		unHtml=Replace(unHtml,chr(9),"&nbsp;&nbsp;&nbsp;&nbsp;")
		unHtml=Replace(unHtml," ","&nbsp;")
	End IF
End Function
'===========================================================================
%>
