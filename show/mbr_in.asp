<%


PS=Request.Form("PS")
ID=Request.Form("ID")
MAIL=Request.Form("MAIL")
strDAY=Request.Form("DAY")
TEL=Request.Form("TEL")
SP=Request.Form("SP")
NAME=Request.Form("NAME")
YN=Request.Form("YN")

if PS="" or Len(PS)>11 then
	Response.write "<script language='javascript'>alert('���~�I  �K�X��J���~�I');history.back();</script>"	
	Response.end
end if
if ID="" or Len(ID)>11 then
	Response.write "<script language='javascript'>alert('���~�I  �b����J���~�I');history.back();</script>"	
	Response.end
end if
if MAIL="" then
	Response.write "<script language='javascript'>alert('���~�I  E-Mail��J���~�I');history.back();</script>"	
	Response.end
end if
if NAME="" then
	Response.write "<script language='javascript'>alert('���~�I  �m�W��J���~�I');history.back();</script>"	
	Response.end
end if
if TEL="" or Len(TEL)<>10 or IsNumeric(TEL)<>True then
	Response.write "<script language='javascript'>alert('���~�I  �q�ܿ�J���~�I');history.back();</script>"	
	Response.end
end if


	set conns = Server.CreateObject("ADODB.Connection")
	conns.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("mbr.mdb")
	set rss = Server.CreateObject("ADODB.Recordset")
	sqls = "SELECT * FROM mbr where ID='"&Request.Form("ID")&"'"
	rss.LockType = 3
	rss.Open sqls, conns, 3
	
	
	if rss.eof then
	
	rss.addnew
	
	rss("PS")=PS
	rss("ID")=ID
	rss("MAIL")=MAIL
	rss("DAY")=strDAY
	rss("TEL")=TEL
	rss("SP")=SP
	rss("NAME")=NAME
	rss("YN")=YN

	rss.update
	rss.close
	conns.close
	
	Response.write "<script language='javascript'>alert('�z����Ƥw�����n��,�z�i�H�}�l�K�T���F! ');location.href= ('shox_paste.asp');</script>"	
	
	
	else
	
	rss.close
	conns.close
	
	Response.write "<script language='javascript'>alert('�z���ΦW�w�s�b,�Ч���A��! ');history.back();</script>"	
	
	
	
	end if
	
	
	
	
	
	
	

	
	
%>