<%
name=Request.Form("name")
mail=Request.Form("mail")
tel=Request.Form("tel")
sp=Request.Form("sp")
add=Request.Form("add")



if name="" then
	Response.write "<script language='javascript'>alert('���~�I  �п�J�z���m�W�I');history.back();</script>"	
	Response.end
end if
if mail="" then
	Response.write "<script language='javascript'>alert('���~�I  �п�J�z��email�I');history.back();</script>"	
	Response.end
end if
if sp="" then
	Response.write "<script language='javascript'>alert('���~�I  �п�J�z����ʹq�ܡI');history.back();</script>"	
	Response.end
end if





	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("login.mdb")
	set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM login"
	rs.LockType = 3
	rs.Open sql, conn, 1
	rs.addnew
	
	rs("name")=name
	rs("mail")=mail
	rs("tel")=tel
	rs("sp")=sp
	rs("add")=add
	
	rs.update
	rs.close
	conn.close




	Response.write "<script language='javascript'>alert('���\�I  �z����Ƥw�n�J���\�I');window.close();</script>"	
%>