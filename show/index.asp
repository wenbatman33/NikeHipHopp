<%
	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("mbr.mdb")
	set rs = Server.CreateObject("ADODB.Recordset")

	sql = "SELECT * FROM body where man<>'a12345' and sbj<>'a12345' Order by NowD desc;"
	rs.LockType = 3
	rs.Open sql, conn, 1
	rs.PageSize=10
	PageCount=rs.PageCount
%>
<html>
<head>
<title>SHOW OFF YOUR STUFF</title>
<meta http-equiv="Content-Type" content="text/html;">
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellpadding="0" cellspacing="0" width="599">
<!-- fwtable fwsrc="Untitled" fwbase="show_index.htm.pp.jpg" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
  <tr>
   <td><img src="img/spacer.gif" width="24" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="114" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="319" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="36" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="19" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="17" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="31" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="22" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="17" height="1" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="1" border="0"></td>
  </tr>

  <tr>
   <td colspan="9"><img name="show_indexhtmpp_r1_c1" src="img/show_index.htm.pp_r1_c1.jpg" width="599" height="94" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="94" border="0"></td>
  </tr>
      <form name="form1" method="post" action="index.asp">
  <tr>
   <td colspan="3"><img name="show_indexhtmpp_r2_c1" src="img/show_index.htm.pp_r2_c1.jpg" width="457" height="26" border="0"></td>
    <td colspan="2" background="img/show_index.htm.pp_r2_c4.jpg" align="center" valign="middle"> 
        <input type="text" name="page" size="3" maxlength="5" value="<%=Request("page")%>">


    </td>
   <td><img name="show_indexhtmpp_r2_c6" src="img/show_index.htm.pp_r2_c6.jpg" width="17" height="26" border="0"></td>
      <td colspan="2"><a href="javascript:pageOut();"><img name="show_indexhtmpp_r2_c7" src="img/show_index.htm.pp_r2_c7.jpg" width="53" height="26" border="0"></a></td>
   <td><img name="show_indexhtmpp_r2_c9" src="img/show_index.htm.pp_r2_c9.jpg" width="17" height="26" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="26" border="0"></td>
  </tr>
      </form>
  <tr>
   <td colspan="5"><img name="show_indexhtmpp_r3_c1" src="img/show_index.htm.pp_r3_c1.jpg" width="512" height="25" border="0"></td>
    <td colspan="2" background="img/show_index.htm.pp_r3_c6.jpg" align="center" valign="bottom"><b><%=PageCount%></b></td>
   <td colspan="2"><img name="show_indexhtmpp_r3_c8" src="img/show_index.htm.pp_r3_c8.jpg" width="39" height="25" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="25" border="0"></td>
  </tr>
  <tr>
   <td><img name="show_indexhtmpp_r4_c1" src="img/show_index.htm.pp_r4_c1.jpg" width="24" height="29" border="0"></td>
    <td><a href="shox_login_1.asp"><img name="show_indexhtmpp_r4_c2" src="img/show_index.htm.pp_r4_c2.jpg" width="114" height="29" border="0"></a></td>
   <td colspan="2"><img name="show_indexhtmpp_r4_c3" src="img/show_index.htm.pp_r4_c3.jpg" width="355" height="29" border="0"></td>
    <td colspan="4"><a href="shox_paste.asp"><img name="show_indexhtmpp_r4_c5" src="img/show_index.htm.pp_r4_c5.jpg" width="89" height="29" border="0"></a></td>
   <td><img name="show_indexhtmpp_r4_c9" src="img/show_index.htm.pp_r4_c9.jpg" width="17" height="29" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="29" border="0"></td>
  </tr>
  <tr>
   <td colspan="9"><img name="show_indexhtmpp_r5_c1" src="img/show_index.htm.pp_r5_c1.jpg" width="599" height="22" border="0"></td>
   <td><img src="img/spacer.gif" width="1" height="22" border="0"></td>
  </tr>
  <tr>
    <td colspan="9" background="img/show_index.htm.pp_r6_c1.jpg" valign="top"> 






















<%






if Request("page")="" then
PageNum=1
else
PageNum=CInt(Request("page"))
end if

''''''''''''''''''''''''''''If rs.EOF then 
''''''''''''''''''''''''''''no=1        
''''''''''''''''''''''''''''else
rs.AbsolutePage=PageNum 
counts=0     
Do while Not counts=rs.PageSize
counts=counts+1
%> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="22">&nbsp;</td>
          <td width="563"> 
            <table border="0" cellpadding="0" cellspacing="0" width="560" align="center">
              <!-- fwtable fwsrc="Untitled" fwbase="show_index-1.jpg" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
              <tr> 
                <td><img src="img/spacer.gif" width="115" height="1" border="0"></td>
                <td><img src="img/spacer.gif" width="245" height="1" border="0"></td>
                <td><img src="img/spacer.gif" width="68" height="1" border="0"></td>
                <td><img src="img/spacer.gif" width="132" height="1" border="0"></td>
                <td><img src="img/spacer.gif" width="1" height="1" border="0"></td>
              </tr>
              <tr> 
                <td background="img/show_index-1_r1_c1.jpg" align="right" valign="bottom"><b><%=rs("man")%>&nbsp;&nbsp;</b></td>
                <td background="img/show_index-1_r1_c2.jpg" valign="bottom">&nbsp;&nbsp;&nbsp;<a href="shox_view.asp?nono=<%=rs("ptr")%>"><font color="#FF9900"><%=mid(rs("sbj"),1,20)%>..</font></a></td>
                <td background="img/show_index-1_r1_c3.jpg" valign="bottom">&nbsp;&nbsp;&nbsp;<font color="#FFFFFF"><%=rs("dot")%></font></td>
                <td background="img/show_index-1_r1_c4.jpg" valign="bottom">&nbsp;&nbsp;&nbsp;<font color="#FFFFFF" size=2><b><%=rs("NowD")%></b></font></td>
                <td><img src="img/spacer.gif" width="1" height="36" border="0"></td>
              </tr>
            </table>
          </td>
          <td width="14">&nbsp;</td>
        </tr>
      </table>


<%
rs.Movenext
if rs.EOF then exit Do
Loop

''''''''''''''''''''''''''''''''''''''''''''''end if

PageCount=rs.PageCount
recordCount=rs.recordCount
Pagesize=rs.Pagesize
rs.close
conn.close
%> 


























    </td>
   <td><img src="img/spacer.gif" width="1" height="367" border="0"></td>
  </tr>
  <tr>
    <td colspan="9"><img name="show_indexhtmpp_r7_c1" src="img/show_index.htm.pp_r7_c1.jpg" width="599" height="27" border="0" usemap="#show_indexhtmpp_r7_c1Map"></td>
   <td><img src="img/spacer.gif" width="1" height="27" border="0"></td>
  </tr>
</table>
<map name="show_indexhtmpp_r7_c1Map">
  <area shape="rect" coords="540,5,582,20" href="JavaScript:window.close();">
</map>
</body>
</html>
<script language=JavaScript>
<!--

//-------------------------
function pageOut() {  

	if ( document.form1.page.value == "" || document.form1.page.value == "0" ) {  
	alert("請輸入頁碼！");  
	document.form1.page.focus();  
	} 
	else if ( document.form1.page.value > <%=PageCount%> ) {  
	alert("您輸入的頁碼超出範圍！");  
	document.form1.page.focus(); 
	} 
	else 
	{document.form1.submit();}  



} 


//-->
</script>