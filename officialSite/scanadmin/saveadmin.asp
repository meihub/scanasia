<!--#include file="conn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
end if
%>
<!--#include file="md5.asp"-->
<%
username=trim(request.Form("username"))
pwd=trim(request.Form("pwd1"))
pwd1=tirm(request.Form("pwd2"))
if pwd=pwd1 then response.Write("<script>alert('两次密码不一样！');window.location='xgmima.asp';</script>"):response.End
sql="select * from admin where admin='" & session("admin") & "'"
set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,3
if rs.recordcount=0 then response.Write("<script>alert('修改密码错误！');window.location='xgmima.asp';</script>"):response.End
if pwd<>"" then rs("password")=md5(pwd)
if username<>"" then rs("admin")=username
rs.update
rs.close
set rs=nothing
response.Write("<script>alert('密码修改成功！');window.location='admin.asp';</script>"):response.End
%>