<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
end if
%>
<!--#include file="md5.asp"-->
<%
username=trim(request.Form("username"))
pwd=trim(request.Form("pwd1"))
pwd1=tirm(request.Form("pwd2"))
if pwd=pwd1 then response.Write("<script>alert('�������벻һ����');window.location='xgmima.asp';</script>"):response.End
sql="select * from admin where admin='" & session("admin") & "'"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
if rs.recordcount=0 then response.Write("<script>alert('�޸��������');window.location='xgmima.asp';</script>"):response.End
if pwd<>"" then rs("password")=md5(pwd)
if username<>"" then rs("admin")=username
rs.update
rs.close
set rs=nothing
response.Write("<script>alert('�����޸ĳɹ���');window.location='admin.asp';</script>"):response.End
%>