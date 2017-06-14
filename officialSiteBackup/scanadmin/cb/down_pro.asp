<!--#include file="../../conn2.asp"-->
<% 
Tid=Int(request( "tid ")) 
Pid=Tid+1 
set   rs=server.CreateObject( "adodb.recordset")
set   rso=server.CreateObject( "adodb.recordset") 
rs.open "select * from c where Classpx= "&Tid&"",conn,1,3 
rso.open "select * from c where Classpx= "&Pid&"",conn,1,3 
rs("Classpx")=rs("Classpx")+1 
rso("Classpx")=rso("Classpx")-1 
rs.Update 
rso.Update 
rs.Close 
rso.Close 
set   rs=Nothing 
set   rso=Nothing 
response.write"<script>window.location.href='TVPphoto-1.asp'</script>"
response.end
%> 
