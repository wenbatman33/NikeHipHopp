<%

name=Request.Form("name")
mail=Request.Form("mail")
card=Request.Form("card")
if name="" then
	Response.write "<script language='javascript'>alert('錯誤！  請輸入您的大名！');history.back();</script>"	
	Response.end
end if
if mail="" then
	Response.write "<script language='javascript'>alert('錯誤！  請輸入您的e-mail！');history.back();</script>"	
	Response.end
end if




if card="1" then
Dim myMail 
Set myMail = CreateObject("CDONTS.NewMail") 
HTML = "<html>"
HTML = HTML & "<head>"
HTML = HTML & "<title>E-CARD</title>"
HTML = HTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=big5"">"
HTML = HTML & "</head>"
HTML = HTML & "<body bgcolor=""#000000"" text=""#000000"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
HTML = HTML & "<a href=""http://www.nike.com.tw/new/welcome/jason.html""><img src=""01_ecard.jpg"" width=""600"" height=""460"" border=""0""></a> "
HTML = HTML & "</body>"
HTML = HTML & "</html>"
myMail.AttachURL "D:\orange\152.104.125.96\hip_hoop_2002\e_card\images\e_card\01_ecard.jpg", "01_ecard.jpg" 
myMail.From = name & "<xx@xx.xx>"
myMail.To = mail 
myMail.Subject = "E-CARD" 
myMail.BodyFormat = 0 
myMail.MailFormat = 0 
myMail.Body = HTML 
myMail.Send 
end if



if card="2" then
Dim myMailb 
Set myMailb = CreateObject("CDONTS.NewMail") 
HTML = "<html>"
HTML = HTML & "<head>"
HTML = HTML & "<title>E-CARD</title>"
HTML = HTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=big5"">"
HTML = HTML & "</head>"
HTML = HTML & "<body bgcolor=""#000000"" text=""#000000"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
HTML = HTML & "<a href=""http://www.nike.com.tw/new/welcome/jason.html""><img src=""02_ecard.jpg"" width=""600"" height=""460"" border=""0""></a> "
HTML = HTML & "</body>"
HTML = HTML & "</html>"
myMailb.AttachURL "D:\orange\152.104.125.96\hip_hoop_2002\e_card\images\e_card\02_ecard.jpg", "02_ecard.jpg" 
myMailb.From = name & "<xx@xx.xx>"
myMailb.To = mail 
myMailb.Subject = "E-CARD" 
myMailb.BodyFormat = 0 
myMailb.MailFormat = 0 
myMailb.Body = HTML 
myMailb.Send 
end if



if card="3" then
Dim myMailc 
Set myMailc = CreateObject("CDONTS.NewMail") 
HTML = "<html>"
HTML = HTML & "<head>"
HTML = HTML & "<title>E-CARD</title>"
HTML = HTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=big5"">"
HTML = HTML & "</head>"
HTML = HTML & "<body bgcolor=""#000000"" text=""#000000"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
HTML = HTML & "<a href=""http://www.nike.com.tw/new/welcome/jason.html""><img src=""03_ecard.jpg"" width=""600"" height=""460"" border=""0""></a> "
HTML = HTML & "</body>"
HTML = HTML & "</html>"
myMailc.AttachURL "D:\orange\152.104.125.96\hip_hoop_2002\e_card\images\e_card\03_ecard.jpg", "03_ecard.jpg" 
myMailc.From = name & "<xx@xx.xx>"
myMailc.To = mail 
myMailc.Subject = "E-CARD" 
myMailc.BodyFormat = 0 
myMailc.MailFormat = 0 
myMailc.Body = HTML 
myMailc.Send 
end if





	Response.write "<script language='javascript'>alert('成功！  您的E-CARD已傳送成功！');window.close();</script>"	





%>