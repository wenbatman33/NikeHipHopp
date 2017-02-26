<%
	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("mobile.mdb")
	set rs = Server.CreateObject("ADODB.Recordset")
	sql = "SELECT * FROM only where dd='"&date&"'"
	rs.LockType = 1
	rs.Open sql, conn, 1
	pp=(rs("pp")+0)
	rs.close
	conn.close
	if pp<=0 then
	Response.write "<script language='javascript'>alert('非常抱歉！  本功能手機圖案下載每天只提供300個名額,以用完！');window.close();</script>"	
	Response.end
	end if
%>
<html>
<head>
<title>mobile</title>
<meta http-equiv="Content-Type" content="text/html;">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Thu Jun 27 23:14:40 GMT+0800 (￥x￥_?D·CRE?!) 2002-->
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellpadding="0" cellspacing="0" width="780">
<!-- fwtable fwsrc="Untitled" fwbase="mobile.jpg" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
  <tr>
   <td><img src="images/spacer.gif" width="66" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="167" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="93" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="67" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="54" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="128" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="7" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="85" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="113" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="1" border="0"></td>
  </tr>

  <tr>
   <td colspan="4"><img name="mobile_r1_c1" src="images/mobile_r1_c1.jpg" width="393" height="96" border="0"></td>
   <td colspan="5"><img name="mobile_r1_c5" src="images/mobile_r1_c5.jpg" width="387" height="96" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="96" border="0"></td>
  </tr>
  <tr>
   <td colspan="2"><img name="mobile_r2_c1" src="images/mobile_r2_c1.jpg" width="233" height="36" border="0"></td>
    <td><img name="mobile_r2_c3" src="icon<%=Request.QueryString("iconno")%>.jpg" width="93" height="36" border="0"></td>
   <td colspan="6"><img name="mobile_r2_c4" src="images/mobile_r2_c4.jpg" width="454" height="36" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="36" border="0"></td>
  </tr>
      <form name="form1" method="post" action="sendout.asp">
  <tr>
   <td colspan="2"><img name="mobile_r3_c1" src="images/mobile_r3_c1.jpg" width="233" height="31" border="0"></td>
    <td colspan="4" background="images/mobile_r3_c3.jpg">
        <input type="text" name="body" size="40" maxlength="70">
    </td>
   <td colspan="3"><img name="mobile_r3_c7" src="images/mobile_r3_c7.jpg" width="205" height="31" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="31" border="0"></td>
  </tr>
  <tr>
   <td colspan="9"><img name="mobile_r4_c1" src="images/mobile_r4_c1.jpg" width="780" height="12" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="12" border="0"></td>
  </tr>

  <tr>
   <td colspan="2"><img name="mobile_r5_c1" src="images/mobile_r5_c1.jpg" width="233" height="34" border="0"></td>
    <td colspan="3" background="images/mobile_r5_c3.jpg">
        <select name="opr">
          <option value="1" selected>中華</option>
          <option value="2">台灣</option>
          <option value="3">遠傳</option>
          <option value="4">和信</option>
          <option value="5">泛亞</option>
          <option value="6">東信</option>
        </select>
        <input type="hidden" name="iconno" value="<%=Request.QueryString("iconno")%>">
      </td>
      <td colspan="4" background="images/mobile_r5_c6.jpg">
        <input type="text" name="PhoneNo" maxlength="10" size="10">
      </td>
   <td><img src="images/spacer.gif" width="1" height="34" border="0"></td>
  </tr>
      </form>
  <tr>
   <td rowspan="3" colspan="7"><img name="mobile_r6_c1" src="images/mobile_r6_c1.jpg" width="582" height="49" border="0"></td>
   <td><img name="mobile_r6_c8" src="images/mobile_r6_c8.jpg" width="85" height="15" border="0"></td>
   <td><img name="mobile_r6_c9" src="images/mobile_r6_c9.jpg" width="113" height="15" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="15" border="0"></td>
  </tr>
  <tr>
    <td><a href="javascript:document.form1.submit();"><img name="mobile_r7_c8" src="images/mobile_r7_c8.jpg" width="85" height="21" border="0"></a></td>
   <td rowspan="2"><img name="mobile_r7_c9" src="images/mobile_r7_c9.jpg" width="113" height="34" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="21" border="0"></td>
  </tr>
  <tr>
   <td><img name="mobile_r8_c8" src="images/mobile_r8_c8.jpg" width="85" height="13" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="13" border="0"></td>
  </tr>
  <tr>
   <td><img name="mobile_r9_c1" src="images/mobile_r9_c1.jpg" width="66" height="56" border="0"></td>
   <td colspan="8"><img name="mobile_r9_c2" src="images/mobile_r9_c2.jpg" width="714" height="56" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="56" border="0"></td>
  </tr>
</table>
</body>
</html>
