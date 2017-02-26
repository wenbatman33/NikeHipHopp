<%
	set conns = Server.CreateObject("ADODB.Connection")
	conns.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("mbr.mdb")
	set rss = Server.CreateObject("ADODB.Recordset")
	sqls = "SELECT * FROM mbr where ID='"&Request("IDno")&"'"
	rss.LockType = 3
	rss.Open sqls, conns, 3
	
	if rss.eof then
	Response.write "<script language='javascript'>window.close();</script>"	
	Response.end
	ens if
	NAME=rss("NAME")
	MAIL=rss("MAIL")
	strDAY=rss("DAY")
	TEL=rss("TEL")




	rss.close
	conns.close


%>
<html>
<head>
<title>show_profile.asp.jpg</title>
<meta http-equiv="Content-Type" content="text/html;">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Tue Jul 16 01:56:25 GMT+0800 (¢Dx¢D_?D¡PCRE?!) 2002-->
</head>
<body bgcolor="#ffffff" background="SHOW%20OFF%20YOUR%20STUFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellpadding="0" cellspacing="0" width="250">
  <!-- fwtable fwsrc="Untitled" fwbase="show_profile.asp.jpg" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
  <tr> 
    <td><img src="img/spacer.gif" width="59" height="1" border="0"></td>
    <td><img src="img/spacer.gif" width="141" height="1" border="0"></td>
    <td><img src="img/spacer.gif" width="7" height="1" border="0"></td>
    <td><img src="img/spacer.gif" width="43" height="1" border="0"></td>
    <td><img src="img/spacer.gif" width="1" height="1" border="0"></td>
  </tr>
  <tr> 
    <td rowspan="2"><img name="show_profileasp_r1_c1" src="img/show_profile.asp_r1_c1.jpg" width="59" height="45" border="0"></td>
    <td colspan="2"><img name="show_profileasp_r1_c2" src="img/show_profile.asp_r1_c2.jpg" width="148" height="7" border="0"></td>
    <td rowspan="2"><img name="show_profileasp_r1_c4" src="img/show_profile.asp_r1_c4.jpg" width="43" height="45" border="0"></td>
    <td><img src="img/spacer.gif" width="1" height="7" border="0"></td>
  </tr>
  <tr> 
    <td colspan="2" background="img/show_profile.asp_r2_c2.jpg"><font size="2">&nbsp;<%=NAME%></font></td>
    <td><img src="img/spacer.gif" width="1" height="38" border="0"></td>
  </tr>
  <tr> 
    <td rowspan="4"><img name="show_profileasp_r3_c1" src="img/show_profile.asp_r3_c1.jpg" width="59" height="105" border="0"></td>
    <td colspan="3" background="img/show_profile.asp_r3_c2.jpg"><font size="2"><%=MAIL%></font></td>
    <td><img src="img/spacer.gif" width="1" height="25" border="0"></td>
  </tr>
  <tr> 
    <td colspan="3" background="img/show_profile.asp_r4_c2.jpg"><font size="2"><%=strDAY%></font></td>
    <td><img src="img/spacer.gif" width="1" height="29" border="0"></td>
  </tr>
  <tr> 
    <td colspan="3" background="img/show_profile.asp_r5_c2.jpg"><font size="2"><%=TEL%></font></td>
    <td><img src="img/spacer.gif" width="1" height="33" border="0"></td>
  </tr>
  <tr> 
    <td><img name="show_profileasp_r6_c2" src="img/show_profile.asp_r6_c2.jpg" width="141" height="18" border="0"></td>
    <td colspan="2"><a href="JavaScript:window.close();"><img name="show_profileasp_r6_c3" src="img/show_profile.asp_r6_c3.jpg" width="50" height="18" border="0"></a></td>
    <td><img src="img/spacer.gif" width="1" height="18" border="0"></td>
  </tr>
</table>
</body>
</html>
