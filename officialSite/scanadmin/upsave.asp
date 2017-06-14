<!--#include FILE="upload_5xsoft.inc"-->
<!--#include FILE="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<div align=center><font size=80 color=red><b>您没有此项目管理权限！</b></font></div>"
response.End
end if
end if
%>
<%
filepath="../learn/website/"		'上传路径
set upload=new upload_5xSoft		'建立上传对象
'upload.NoAllowExt="asp;asa;cer;aspx;cs;vb;js;"		'设置上传类型的黑名单
'upload.GetData (3072000)	'取得上传数据,限制最大上传3M

if upload.form("act")="uploadfile" then
	for each formName in upload.File
		set file=upload.File(formName)
		'fileExt=lcase(file.FileExt)		'得到的文件扩展名不含有.
		fileExt=lcase(right(file.filename,3))
			if file.filesize<10 then
				response.write "<span style=""font-family: 宋体; font-size: 9pt"">请先选择你要上传的文件！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
				response.end
			end if
			if fileext<>"gif" and fileext<>"jpg"  then
				response.write "<span style=""font-family: 宋体; font-size: 9pt"">只能上传jpg或gif格式的图片！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
				response.end
			end if
			if file.filesize>(300*1024) then
				response.write "<span style=""font-family: 宋体; font-size: 9pt"">最大只能上传 300K 的图片文件！　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
				response.end
			end if

		dtNow=Now()
		randomize
		'ranNum=int(90000*rnd)+10000
		filename1=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2) &"."&fileExt
		filename=filepath&filename1
		
		if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
			'upload.SaveToFile formName,Server.mappath(FileName)
			file.SaveAs Server.mappath(FileName)   ''保存文件
 ' file.SaveAs Server.mappath(FileName1)

			'这里可以存数据库


			response.write "<script>window.opener.document.form1.web_piclink.value='" & filename1 & "'</script>"
			%>
			<script language="javascript">
				window.alert("文件上传成功!请不要修改生成的链接地址！");
				window.close();
			</script>
			<%
		end if
		set file=nothing
	next
	set upload=nothing
end if
%>