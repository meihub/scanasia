
<%
'------------------------------------------网站公共函数--------------------------------------



'**************************************************
'函数名:deletefile
'参数:path ------ 文件路径
'作用:通过获取文件路径删除文件
function deletefile(path)
set fso=server.CreateObject("Scripting.FilesystemObject")
if fso.FileExists(Server.MapPath(path)) then
fso.DeleteFile(Server.MapPath(path))
end if
set fso=nothing
end function
'**************************************************




'*************************************************************************************
'函数名:go\back
'参数:rightmsg/error ------ 弹出对话框正确/错误信息
'     url ------ 转向页面
'作用:正确/错误信息提示及页面转向
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


'以时间生成文件名的函数*****************
function makefilename(fname)
fname = fname
fname = replace(fname,"-","")
fname = replace(fname," ","") 
fname = replace(fname,":","")
fname = replace(fname,"PM","")
fname = replace(fname,"AM","")
fname = replace(fname,"上午","")
fname = replace(fname,"下午","")
makefilename=fname & ".html"
end function 
'***************************************


'======================显示空格与换行========================================
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
