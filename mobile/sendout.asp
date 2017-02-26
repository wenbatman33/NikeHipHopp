<%
	set conns = Server.CreateObject("ADODB.Connection")
	conns.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("mobile.mdb")
	set rss = Server.CreateObject("ADODB.Recordset")
	sqls = "SELECT * FROM only where dd='"&date&"'"
	rss.LockType = 1
	rss.Open sqls, conns, 1
	pp=(rss("pp")+0)
	rss.close
	conns.close
	if pp<=0 then
	Response.write "<script language='javascript'>alert('非常抱歉！  本功能手機圖案下載每天只提供300個名額,以用完！');window.close();</script>"	
	Response.end
	end if


body=Request.Form("body")
opr=Request.Form("opr")
iconno=Request.Form("iconno")
PhoneNo=Request.Form("PhoneNo")

if body="" then
	Response.write "<script language='javascript'>alert('錯誤！  請輸入您要傳送的訊息！');history.back();</script>"	
	Response.end
end if
if len(body)>70 then
	Response.write "<script language='javascript'>alert('錯誤！  您輸入的訊息不能超過70個字！');history.back();</script>"	
	Response.end
end if
if mid(Request.Form("PhoneNo"),1,2)<>"09" or len(Request.Form("PhoneNo"))<>10 or IsNumeric(Request.Form("PhoneNo"))<>True then
	Response.write "<script language='javascript'>alert('錯誤！  您輸入的手機號碼不正確！');history.back();</script>"	
	Response.end
end if


Sopr="NoCode"
if mid(PhoneNo,1,4)="0935" or mid(PhoneNo,1,4)="0939" or mid(PhoneNo,1,4)="0920" or mid(PhoneNo,1,4)="0922" or mid(PhoneNo,1,4)="0958" or mid(PhoneNo,1,4)="0952" or mid(PhoneNo,1,4)="0918" or mid(PhoneNo,1,4)="0953" then
Sopr=",64,f6,79,0a"
end if
if mid(PhoneNo,1,4)="0910" or mid(PhoneNo,1,4)="0911" or mid(PhoneNo,1,4)="0912" or mid(PhoneNo,1,4)="0919" or mid(PhoneNo,1,4)="0921" or mid(PhoneNo,1,4)="0928" or mid(PhoneNo,1,4)="0932" or mid(PhoneNo,1,4)="0933" or mid(PhoneNo,1,4)="0937" then
Sopr=",64,f6,29,0a"
end if
if mid(PhoneNo,1,4)="0917" or mid(PhoneNo,1,4)="0916" or mid(PhoneNo,1,4)="0930" or mid(PhoneNo,1,4)="0936" or mid(PhoneNo,1,4)="0955" or mid(PhoneNo,1,4)="0926" then
Sopr=",64,f6,10,0a"
end if
if mid(PhoneNo,1,4)="0938" or mid(PhoneNo,1,4)="0925" or mid(PhoneNo,1,4)="0915" or mid(PhoneNo,1,4)="0927" or mid(PhoneNo,1,4)="0913" then
Sopr=",64,f6,88,0a"
end if
if mid(PhoneNo,1,4)="0929" or mid(PhoneNo,1,4)="0956" then
Sopr=",64,f6,99,0a"
end if

if mid(PhoneNo,1,5)="09310" or mid(PhoneNo,1,4)="09311" or mid(PhoneNo,1,4)="09312" or mid(PhoneNo,1,4)="09313" then
Sopr=",64,f6,10,0a"
end if
if mid(PhoneNo,1,5)="09317" or mid(PhoneNo,1,4)="09318" or mid(PhoneNo,1,4)="09319" then
Sopr=",64,f6,99,0a"
end if
if mid(PhoneNo,1,5)="09314" or mid(PhoneNo,1,4)="09315" or mid(PhoneNo,1,4)="09316" then
Sopr=",64,f6,39,0a"
end if
if mid(PhoneNo,1,5)="09230" or mid(PhoneNo,1,4)="09231" or mid(PhoneNo,1,4)="09232" or mid(PhoneNo,1,5)="09233" or mid(PhoneNo,1,4)="09234" or mid(PhoneNo,1,4)="09235" or mid(PhoneNo,1,4)="09236" then
Sopr=",64,f6,39,0a"
end if

Shdr=",06,05,04,15,82,00,00,30"
if Sopr="NoCode" then
	Select Case opr
	
		case "1" 
			Sopr=",64,f6,29,0a"        '中華
		case "2"
			Sopr=",64,f6,79,0a"        '台灣
		case "3"
			Sopr=",64,f6,10,0a"        '遠傳
		case "4"
			Sopr=",64,f6,88,0a"        '和信
		case "5"
			Sopr=",64,f6,99,0a"        '泛亞
		case "6"
			Sopr=",64,f6,39,0a"        '東信
	End Select
end if

Select Case iconno

	case "1"        'icon1
		Sicon=",00,48,0E,01,00,00,00,00,00,00,00,00,00,7E,7E,7F,CF,DF,FC,00,00,00,41,43,41,32,66,64,00,00,00,4C,C9,41,32,66,64,00,00,06,4C,C9,4F,32,26,64,10,00,3C,4C,C9,4F,32,26,4C,30,03,E0,4C,C3,41,32,06,18,30,3F,80,4C,C9,41,32,46,48,79,FE,00,4C,C9,4F,32,46,48,7F,F8,00,4C,C9,49,32,66,64,7F,C0,00,4C,C9,C9,32,66,64,7F,00,00,41,49,48,86,66,64,3C,00,00,7E,7F,F8,7B,DF,FC,00,00,00,00,00,00,00,00,00,00,00,00"
	case "2"        'icon2
		Sicon=",00,48,0E,01,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,07,80,00,00,00,00,35,71,AE,EC,C0,00,00,00,00,35,49,AA,AC,C0,00,00,00,00,35,49,AA,AC,C0,00,80,00,07,3D,71,EA,AF,C0,03,00,00,0E,3D,61,AE,EF,88,0C,00,00,1E,35,41,AE,EF,18,78,00,00,1F,E5,40,00,0C,1F,E0,00,00,1F,C0,00,00,0C,1F,80,00,00,0F,80,00,00,0C,0E,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00,00"
	case "3"        'icon3
		Sicon=",00,48,0E,01,01,FC,5F,FF,00,00,00,00,00,01,FC,07,FF,00,00,00,00,00,00,F8,3F,FE,00,00,00,0C,00,00,FE,67,FE,02,00,00,70,00,00,7B,5B,FC,06,00,03,C0,00,00,3F,21,FC,06,00,1F,00,00,00,18,81,F8,0E,01,F8,00,00,00,0C,03,F0,0F,0F,E0,00,00,00,07,77,E0,1F,FF,80,00,00,00,06,1F,80,1F,FE,00,00,00,00,02,CF,80,0F,F8,00,00,00,00,03,0F,00,07,E0,00,00,00,00,01,9E,00,00,00,00,00,00,00,00,FC,00,00,00,00,00,00"
	case "4"        'icon4
		Sicon=",00,48,0E,01,12,00,07,40,20,FE,2F,FF,80,14,EE,E2,67,70,FE,03,FF,80,18,AA,A2,51,20,7C,1F,FF,00,18,EE,E2,57,20,7F,33,FF,00,14,88,82,55,20,3D,AD,FE,00,12,EE,82,57,30,1F,90,FE,00,00,00,80,00,00,0C,40,FC,00,01,C0,10,3A,80,06,01,F8,00,01,15,D2,2A,00,03,BB,F0,00,01,15,54,2A,AB,83,0F,C2,03,01,D5,58,2A,AA,81,67,C6,3E,01,15,54,3A,AB,81,87,87,F8,01,15,52,2A,AA,00,CF,07,E0,01,1D,51,2A,93,80,7E,03,80"
End Select



	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("mobile.mdb")
	set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM mobile"
	rs.LockType = 3
	rs.Open sql, conn, 1
	rs.addnew
	rs("PhoneNo")=PhoneNo
	rs("icon")=iconno
	rs.update
	rs.close
	conn.close
	
	set connp = Server.CreateObject("ADODB.Connection")
	connp.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("mobile.mdb")
	set rsp = Server.CreateObject("ADODB.Recordset")
	sqlp = "SELECT * FROM only where dd='"&date&"'"
	rsp.LockType = 2
	rsp.Open sqlp, connp, 1
	sppp=(rsp("pp")+0)-1
	rsp("pp")=sppp
	rsp.update
	rsp.close
	connp.close


set smstxt = Server.CreateObject("EMCTOOL.BTOSMS")
smstxt.SID="SM87001803"                '企業商訊帳號
smstxt.SPS="orange"                    '企業商訊密碼
smstxt.PhoneNo=PhoneNo                 '接收人 電話
smstxt.SMBody=body                    '接收內容
smstxt.STY="1"                         '發送類型
retcode=smstxt.retcode

if mid(retcode,1,2)="OK" then
	Sbody=Shdr & Sopr & Sicon
	set sms = Server.CreateObject("EMCTOOL.BTOSMS")
	sms.SID="SM87001803"                '企業商訊帳號
	sms.SPS="orange"                    '企業商訊密碼
	sms.PhoneNo=PhoneNo                 '接收人 電話
	sms.SMBody=Sbody                    '接收內容
	sms.STY="4"                         '發送類型
	retcode=sms.retcode
end if

if (mid(retcode,3,8)+0)<2500 then
			Dim myMailt
			Set myMailt = CreateObject("CDONTS.NewMail") 
			Bodyt=Bodyt & vbCrLf & "簡訊點數即將用完....."
			Bodyt=Bodyt & vbCrLf & "目前剩餘: " & (mid(retcode,3,8)+0) & " 點"
			myMailt.Body = Bodyt 
			myMailt.From = "matt@moulin-orange.com"
			myMailt.Bcc = "ekin@mobits.com"
			myMailt.Subject = "簡訊點數即將用完....." 
			myMailt.BodyFormat = 1 
			myMailt.MailFormat = 1 
			myMailt.Send 
end if

if retcode="ER07" then
	Response.write "<script language='javascript'>alert('成功！  您的訊息及圖案已發射成功！"& retcode &"');window.close();</script>"	
else
	Response.write "<script language='javascript'>alert('成功！  您的訊息及圖案已發射成功！');window.close();</script>"	
end if

%>