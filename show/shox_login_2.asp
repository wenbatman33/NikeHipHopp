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
	Response.write "<script language='javascript'>alert('¿ù»~¡I  ±K½X¿é¤J¿ù»~¡I');history.back();</script>"	
	Response.end
end if
if ID="" or Len(ID)>11 then
	Response.write "<script language='javascript'>alert('¿ù»~¡I  ±b¸¹¿é¤J¿ù»~¡I');history.back();</script>"	
	Response.end
end if
if MAIL="" then
	Response.write "<script language='javascript'>alert('¿ù»~¡I  E-Mail¿é¤J¿ù»~¡I');history.back();</script>"	
	Response.end
end if
if NAME="" then
	Response.write "<script language='javascript'>alert('¿ù»~¡I  ©m¦W¿é¤J¿ù»~¡I');history.back();</script>"	
	Response.end
end if
if TEL="" or Len(TEL)<>10 or IsNumeric(TEL)<>True then
	Response.write "<script language='javascript'>alert('¿ù»~¡I  ¹q¸Ü¿é¤J¿ù»~¡I');history.back();</script>"	
	Response.end
end if
%>
<html>
<head>
<title>SHOW OFF YOUR STUFF</title>
<meta http-equiv="Content-Type" content="text/html;">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Mon Jul 15 11:47:27 GMT+0800 (¢Dx¢D_?D¡PCRE?!) 2002-->
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellpadding="0" cellspacing="0" width="599">
      <form name="form1" method="post" action="mbr_in.asp">
  <tr>
   <td><img src="img/spacer.gif" width="126" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="201" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="158" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="44" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="29" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="18" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="23" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="1" border="0"></td>
  </tr>

  <tr>
   <td colspan="7"><img name="shox_login_2htmasp_r1_c1" src="img/shox_login_2.htm.asp_r1_c1.jpg" width="599" height="104" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="104" border="0"></td>
  </tr>
  <tr>
   <td colspan="7"><img name="shox_login_2htmasp_r2_c1" src="img/shox_login_2.htm.asp_r2_c1.jpg" width="599" height="21" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="21" border="0"></td>
  </tr>
  <tr>
   <td colspan="4"><img name="shox_login_2htmasp_r3_c1" src="img/shox_login_2.htm.asp_r3_c1.jpg" width="529" height="21" border="0"></td>
      <td colspan="2"><a href="javascript:history.back();"><img name="shox_login_2htmasp_r3_c5" src="img/shox_login_2.htm.asp_r3_c5.jpg" width="47" height="21" border="0"></a></td>
   <td><img name="shox_login_2htmasp_r3_c7" src="img/shox_login_2.htm.asp_r3_c7.jpg" width="23" height="21" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="21" border="0"></td>
  </tr>
  <tr>
   <td rowspan="8"><img name="shox_login_2htmasp_r4_c1" src="img/shox_login_2.htm.asp_r4_c1.jpg" width="126" height="272" border="0"></td>
   <td rowspan="8"><img name="shox_login_2htmasp_r4_c2" src="img/shox_login_2.htm.asp_r4_c2.jpg" width="201" height="272" border="0"></td>
    <td colspan="3" background="img/shox_login_2.htm.asp_r4_c3.jpg">&nbsp;<%=Request.Form("PS")%></td>
   <td rowspan="7" colspan="2"><img name="shox_login_2htmasp_r4_c6" src="img/shox_login_2.htm.asp_r4_c6.jpg" width="41" height="243" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="33" border="0"></td>
  </tr>
  <tr>
    <td colspan="3" background="img/shox_login_2.htm.asp_r5_c3.jpg">&nbsp;<%=Request.Form("ID")%></td>
   <td><img src="img/spacer.gif" width="1" height="33" border="0"></td>
  </tr>
  <tr>
    <td colspan="3" background="img/shox_login_2.htm.asp_r6_c3.jpg">&nbsp;<%=Request.Form("MAIL")%></td>
   <td><img src="img/spacer.gif" width="1" height="38" border="0"></td>
  </tr>
  <tr>
    <td colspan="3" background="img/shox_login_2.htm.asp_r7_c3.jpg">&nbsp;<%=Request.Form("DAY")%></td>
   <td><img src="img/spacer.gif" width="1" height="35" border="0"></td>
  </tr>
  <tr>
    <td colspan="3" background="img/shox_login_2.htm.asp_r8_c3.jpg">&nbsp;<%=Request.Form("TEL")%></td>
   <td><img src="img/spacer.gif" width="1" height="35" border="0"></td>
  </tr>
  <tr>
    <td colspan="3" background="img/shox_login_2.htm.asp_r9_c3.jpg">&nbsp;<%=Request.Form("SP")%></td>
   <td><img src="img/spacer.gif" width="1" height="34" border="0"></td>
  </tr>
  <tr>
    <td colspan="3" background="img/shox_login_2.htm.asp_r10_c3.jpg">&nbsp;<%=Request.Form("NAME")%>
        <input type="hidden" name="PS" value="<%=Request.Form("PS")%>">
        <input type="hidden" name="ID" value="<%=Request.Form("ID")%>">
        <input type="hidden" name="MAIL" value="<%=Request.Form("MAIL")%>">
        <input type="hidden" name="DAY" value="<%=Request.Form("DAY")%>">
        <input type="hidden" name="TEL" value="<%=Request.Form("TEL")%>">
        <input type="hidden" name="SP" value="<%=Request.Form("SP")%>">
        <input type="hidden" name="NAME" value="<%=Request.Form("NAME")%>">
        <input type="hidden" name="YN" value="<%=Request.Form("YN")%>">
      </td>
   <td><img src="img/spacer.gif" width="1" height="35" border="0"></td>
  </tr>
  <tr>
   <td><img name="shox_login_2htmasp_r11_c3" src="img/shox_login_2.htm.asp_r11_c3.jpg" width="158" height="29" border="0"></td>
      <td colspan="3"><a href="javascript:document.form1.submit();"><img name="shox_login_2htmasp_r11_c4" src="img/shox_login_2.htm.asp_r11_c4.jpg" width="91" height="29" border="0"></a></td>
   <td><img name="shox_login_2htmasp_r11_c7" src="img/shox_login_2.htm.asp_r11_c7.jpg" width="23" height="29" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="29" border="0"></td>
  </tr>
  <tr>
   <td colspan="7"><img name="shox_login_2htmasp_r12_c1" src="img/shox_login_2.htm.asp_r12_c1.jpg" width="599" height="172" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="172" border="0"></td>
  </tr>
      </form>
</table>
</body>
</html>
