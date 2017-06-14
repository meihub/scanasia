<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%dim admin,password
admin=replace(trim(request("admin")),"'","")
password=md5(replace(trim(request("password")),"'",""))
if admin="" or password="" then
response.Write "<center><a href=login.asp><font color=red size=2>对不起，登陆失败，请检查您的登陆名和密码</font></a></center>"
response.end
end if
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from admin where admin='" & admin & "' and password='" & password & "'" ,conn,1,3
if not(rs.bof and rs.eof) then
if password=rs("password") then
session("admin")=trim(rs("admin"))
session("last_logintime")=rs("last_logintime")
session.Timeout=30
rs("last_logintime")=now()
rs.update
rs.Close
set rs=nothing
response.Redirect "index.asp"
else
response.write "<script LANGUAGE='javascript'>alert('对不起，登陆失败！');history.go(-1);</script>"

end if
else
response.write "<script LANGUAGE='javascript'>alert('对不起，登陆失败！');history.go(-1);</script>"

end if

%>