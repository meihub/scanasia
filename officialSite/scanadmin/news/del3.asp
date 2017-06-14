<% @LANGUAGE="VBSCRIPT" %>
<!--#include file="../../conn2.asp" -->
<%if session("admin")="" then
response.Redirect "../login.asp"
end if
%>
<%
dim strID
strID=Request("checkbox")
if strID <>"" then
dim sql
sql="DELETE FROM news WHERE id IN "   
sql=sql &"(" & strID & ")"
set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = conn
Command1.CommandText = sql
Command1.Execute()
Response.Redirect("newsadmin.asp")
 else Response.redirect("newsadmin.asp")
 end if %>
