<!--#include file="../Connections/chao.asp"-->
<!--#include file="../ScriptLibrary/function.asp" -->

<%



ID=request("ID")''''�޸�ʱ��Ӧ��ID
works_img=request("works_image")
works_imgo=request("works_imageo")

works_img2=request("works_image2")
works_imgo2=request("works_imageo2")


set rs=server.CreateObject("adodb.recordset")
rs.open "select * from c where ID="&ID,MM_chao_STRING,1,1



if works_img<>works_imgo then''''���и���ͼƬ��ɾ��ԭ���ľ�ͼƬ
call deletefile(works_imgo)
end if
rs.close
set rs=nothing




'=�޸����ݿ�works==============================================================
set rs=server.createobject("adodb.recordset")
rs.open "select * from c where ID="&ID,MM_chao_STRING,1,3


rs("ProName")=request("name")



rs("ClassID")=request("kind_id0")
rs("scc")=request("kind_id1")
rs("ProImg")=request("works_image")
rs("BigImg")=request("works_image2")


rs("ProInfo")=request("works_info")

rs.update
rs.close
set rs=nothing
'================================================================================





response.write "<script language=javascript>alert('��Ʒ�޸ĳɹ���');window.location.href='works_id_manage.asp?ID="&ID&"';</script>"
response.End


%>


