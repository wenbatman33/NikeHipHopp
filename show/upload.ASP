<%@ENABLESESSIONSTATE=FALSE%>
<%
set Upload = Server.CreateObject("Dundas.Upload")
set NextFile = Upload.GetNextFile

	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("mbr.mdb")
	set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM mbr where PS='"&Upload.Form("PS")&"' and ID='"&Upload.Form("ID")&"'"
	rs.LockType = 3
	rs.Open sql, conn, 3
	if rs.eof then
	rs.close
	conn.close
	Response.write "<script language='javascript'>alert('錯誤！  您的帳號密碼不對,請您先去登錄資料！');location.href= ('shox_login_1.asp');</script>"	
	Response.end
	else
	rs.close
	conn.close
	end if

if Upload.Form("man")="" or len(Upload.Form("man"))>7 then
	Response.write "<script language='javascript'>alert('錯誤！  貼圖者輸入錯誤！');history.back();</script>"	
	Response.end
end if
if Upload.Form("sbj")="" then
	Response.write "<script language='javascript'>alert('錯誤！  主題輸入錯誤！');history.back();</script>"	
	Response.end
end if
'if Upload.Form("file1") = "" then
'	Response.write "<script language='javascript'>alert('錯誤！  您沒有選擇圖檔！');history.back();</script>"	
'	Response.end
'end if






imgname = mid(NextFile.Filename,instr(1,NextFile.Filename,"}_") + 2 ,len(NextFile.Filename) - instr(1,NextFile.Filename,"}_"))
imgnameSub = Right(imgname,4)
if imgnameSub = ".jpg" or imgnameSub = ".gif" then
else
	Response.write "<script language='javascript'>alert('錯誤！  請選擇適當的圖檔！');history.back();</script>"	
	Response.end
end if









	set conns = Server.CreateObject("ADODB.Connection")
	conns.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("mbr.mdb")
	set rss = Server.CreateObject("ADODB.Recordset")
	sqls = "SELECT * FROM body"
	rss.LockType = 2
	rss.Open sqls, conns, 3
	rss.MoveLast
	imgNo=(rss("ptr")+0)+1
	
	rss.addnew
	
	rss("man")=Upload.Form("man")
	rss("sbj")=Upload.Form("sbj")
	rss("body")=Upload.Form("body")
	rss("dateDay")=Upload.Form("dateDay")
	rss("dot")="0"
	rss("NowD")=date & " " &Hour(now)&":"&Minute(Now)
	rss("ptr")=imgNo
	rss("ID")=Upload.Form("ID")
	rss.update
	rss.close
	conns.close



Fullname=imgNo & imgnameSub


NextFile.Filename = Fullname
NextFile.Save server.MapPath("/hip_hoop_2002/show/ptr/")

	Response.write "<script language='javascript'>alert('您的訊息已貼上! ');location.href= ('index.asp');</script>"	
	
		
%>