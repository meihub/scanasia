<!--#include FILE="upload_5xsoft.inc"-->
<!--#include FILE="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")>1 then
response.Write "<div align=center><font size=80 color=red><b>��û�д���Ŀ����Ȩ�ޣ�</b></font></div>"
response.End
end if
end if
%>
<%
filepath="../learn/website/"		'�ϴ�·��
set upload=new upload_5xSoft		'�����ϴ�����
'upload.NoAllowExt="asp;asa;cer;aspx;cs;vb;js;"		'�����ϴ����͵ĺ�����
'upload.GetData (3072000)	'ȡ���ϴ�����,��������ϴ�3M

if upload.form("act")="uploadfile" then
	for each formName in upload.File
		set file=upload.File(formName)
		'fileExt=lcase(file.FileExt)		'�õ����ļ���չ��������.
		fileExt=lcase(right(file.filename,3))
			if file.filesize<10 then
				response.write "<span style=""font-family: ����; font-size: 9pt"">����ѡ����Ҫ�ϴ����ļ�����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
				response.end
			end if
			if fileext<>"gif" and fileext<>"jpg"  then
				response.write "<span style=""font-family: ����; font-size: 9pt"">ֻ���ϴ�jpg��gif��ʽ��ͼƬ����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
				response.end
			end if
			if file.filesize>(300*1024) then
				response.write "<span style=""font-family: ����; font-size: 9pt"">���ֻ���ϴ� 300K ��ͼƬ�ļ�����[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
				response.end
			end if

		dtNow=Now()
		randomize
		'ranNum=int(90000*rnd)+10000
		filename1=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2) &"."&fileExt
		filename=filepath&filename1
		
		if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
			'upload.SaveToFile formName,Server.mappath(FileName)
			file.SaveAs Server.mappath(FileName)   ''�����ļ�
 ' file.SaveAs Server.mappath(FileName1)

			'������Դ����ݿ�


			response.write "<script>window.opener.document.form1.web_piclink.value='" & filename1 & "'</script>"
			%>
			<script language="javascript">
				window.alert("�ļ��ϴ��ɹ�!�벻Ҫ�޸����ɵ����ӵ�ַ��");
				window.close();
			</script>
			<%
		end if
		set file=nothing
	next
	set upload=nothing
end if
%>