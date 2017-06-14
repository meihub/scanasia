<!--#include file="../../conn2.asp"-->
<% 
Tid=Int(request( "tid ")) 
Pid=Tid-1 
set   rsw=Server.CreateObject( "adodb.recordset ") 
  sql="select   *   from   c   where   Classpx= "&Tid&" "
 rsw.open sql,conn,1,3 
  
set   rso=Server.CreateObject( "adodb.recordset ") 
sql="select   *   from  c   where   Classpx= "&Pid&"  "
 rs0.open sql,conn,1,3 

rsw("Classpx")=rsw("Classpx")-1 
rso("Classpx")=rso("Classpx")+1
rsw.Update 
rso.Update 
rsw.Close 
rso.Close 
set   rsw=Nothing 
set   rso=Nothing 
response.write"<script>window.location.href='TVPphoto-1.asp'</script>"
response.end
%>
