<%

name=Request.Form("name")
mail=Request.Form("mail")
sbj=Request.Form("sbj")
body=Request.Form("body")
if name="" then
	Response.write "<script language='javascript'>alert('錯誤！  請輸入您的大名！');history.back();</script>"	
	Response.end
end if
if mail="" then
	Response.write "<script language='javascript'>alert('錯誤！  請輸入e-mail！');history.back();</script>"	
	Response.end
end if
if sbj="" then
	Response.write "<script language='javascript'>alert('錯誤！  請輸入主旨！');history.back();</script>"	
	Response.end
end if



Dim myMail 
Set myMail = CreateObject("CDONTS.NewMail") 
HTML = "<html>"
HTML = HTML & "<head>"
HTML = HTML & "<title>Untitled Document</title>"
HTML = HTML & "<meta http-equiv=""Content-Type"" content=""text/html; charset=big5"">"
HTML = HTML & "</head>"
HTML = HTML & "<body bgcolor=""#000000"" text=""#000000"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"">"
HTML = HTML & "<div align=""center""><a href=""http://www.nike.com.tw/new/welcome/jason.html""><img src=""invitation_mail.jpg"" width=""650"" height=""340"" border=""0""></a> "
HTML = HTML & "</div>"
HTML = HTML & "</body>"
HTML = HTML & "</html>"




myMail.AttachURL "D:\orange\152.104.125.96\hip_hoop_2002\e_card\images\invitation_mail\invitation_mail.jpg", "invitation_mail.jpg" 
myMail.From = name & "<xx@xx.xx>"
myMail.To = mail 
myMail.Subject = sbj 
myMail.BodyFormat = 0 
myMail.MailFormat = 0 
myMail.Body = HTML 
myMail.Send 
	Response.write "<script language='javascript'>alert('成功！  您的邀請函已傳送成功！');window.close();</script>"	







%>