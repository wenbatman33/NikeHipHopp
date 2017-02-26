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
	Response.write "<script language='javascript'>alert('錯誤！  密碼輸入錯誤！');history.back();</script>"	
	Response.end
end if
if ID="" or Len(ID)>11 then
	Response.write "<script language='javascript'>alert('錯誤！  帳號輸入錯誤！');history.back();</script>"	
	Response.end
end if
if MAIL="" then
	Response.write "<script language='javascript'>alert('錯誤！  E-Mail輸入錯誤！');history.back();</script>"	
	Response.end
end if
if NAME="" then
	Response.write "<script language='javascript'>alert('錯誤！  姓名輸入錯誤！');history.back();</script>"	
	Response.end
end if
if TEL="" or Len(TEL)<>10 or IsNumeric(TEL)<>True then
	Response.write "<script language='javascript'>alert('錯誤！  電話輸入錯誤！');history.back();</script>"	
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
	
	Response.write "<script language='javascript'>alert('您的資料已完成登錄,您可以開始貼訊息了! ');location.href= ('shox_paste.asp');</script>"	
	
	
	else
	
	rss.close
	conns.close
	
	Response.write "<script language='javascript'>alert('您的匿名已存在,請更改後再試! ');history.back();</script>"	
	
	
	
	end if
	
	
	
	
	
	
	

	
	
%>