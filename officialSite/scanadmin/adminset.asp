
          <%
		  set rsqx=server.CreateObject("adodb.recordset")
            rsqx.Open "select * from shop_admin where admin='"&session("admin")&"'",conn,1,1
			if not(rsqx.bof and rsqx.eof) then
			qx=rsqx("qx")
			qx=split(qx,",")
qx1=qx(0)
qx2=qx(1)
qx3=qx(2)
qx4=qx(3)
qx5=qx(4)
qx6=qx(5)
qx7=qx(6)
qx8=qx(7)
qx9=qx(8)
qx10=qx(9)
qx11=qx(10)
qx12=qx(11)
qx13=qx(12)
qx14=qx(13)
qx15=qx(14)
qx16=qx(15)
qx17=qx(16)
qx18=qx(17)
qx19=qx(18)
qx20=qx(19)
else
response.write "<script LANGUAGE='javascript'>alert('对不起！操作过期！');window.location.href='login.asp';</script>"
end if
		rsqx.close
		set rsqx=nothing
			%>
         