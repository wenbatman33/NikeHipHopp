
<%

	set conns = Server.CreateObject("ADODB.Connection")
	conns.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("mbr.mdb")
	set rss = Server.CreateObject("ADODB.Recordset")
	sqls = "SELECT * FROM body where ptr='"&Request("nono")&"'"
	rss.LockType = 3
	rss.Open sqls, conns, 3
	
	
	man=rss("man")
	sbj=rss("sbj")
	body=rss("body")
	dateDay=rss("dateDay")

	rss("dot")=(rss("dot")+0)+1

	NowD=rss("NowD")
	ptr=rss("ptr")
	ID=rss("ID")
	
	rss.update
	rss.close
	conns.close





%>

<html>
<head>
<title>SHOW OFF YOUR STUFF</title>
<meta http-equiv="Content-Type" content="text/html;">
<link rel="stylesheet" href="css.css" type="text/css">
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellpadding="0" cellspacing="0" width="599">
  <tr>
    <td width="57"><img src="img/spacer.gif" width="57" height="1" border="0"></td>
    <td width="7"><img src="img/spacer.gif" width="7" height="1" border="0"></td>
    <td width="76"><img src="img/spacer.gif" width="76" height="1" border="0"></td>
    <td width="9"><img src="img/spacer.gif" width="9" height="1" border="0"></td>
    <td width="112"><img src="img/spacer.gif" width="112" height="1" border="0"></td>
    <td width="84"><img src="img/spacer.gif" width="84" height="1" border="0"></td>
    <td width="67"><img src="img/spacer.gif" width="67" height="1" border="0"></td>
    <td width="37"><img src="img/spacer.gif" width="37" height="1" border="0"></td>
    <td width="7"><img src="img/spacer.gif" width="7" height="1" border="0"></td>
    <td width="37"><img src="img/spacer.gif" width="37" height="1" border="0"></td>
    <td width="35"><img src="img/spacer.gif" width="35" height="1" border="0"></td>
    <td width="34"><img src="img/spacer.gif" width="34" height="1" border="0"></td>
    <td width="20"><img src="img/spacer.gif" width="20" height="1" border="0"></td>
    <td width="17"><img src="img/spacer.gif" width="17" height="1" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="1" border="0"></td>
  </tr>

  <tr>
   <td colspan="6"><img name="shox_viewhtmas_r1_c1" src="img/shox_view.htm.as_r1_c1.jpg" width="345" height="99" border="0"></td>
   <td colspan="8"><img name="shox_viewhtmas_r1_c7" src="img/shox_view.htm.as_r1_c7.jpg" width="254" height="99" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="99" border="0"></td>
  </tr>
  <tr>
   <td colspan="6"><img name="shox_viewhtmas_r2_c1" src="img/shox_view.htm.as_r2_c1.jpg" width="345" height="19" border="0"></td>
   <td rowspan="3" colspan="5"><img name="shox_viewhtmas_r2_c7" src="img/shox_view.htm.as_r2_c7.jpg" width="183" height="45" border="0"></td>
   <td rowspan="2" colspan="3"><img name="shox_viewhtmas_r2_c12" src="img/shox_view.htm.as_r2_c12.jpg" width="71" height="25" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="19" border="0"></td>
  </tr>
  <tr>
    <td rowspan="2" width="57"><img name="shox_viewhtmas_r3_c1" src="img/shox_view.htm.as_r3_c1.jpg" width="57" height="26" border="0"></td>
    <td rowspan="2" colspan="2"><a href="#" onClick="MM_openBrWindow('show_profile.asp?IDno=<%=ID%>','','width=250,height=150')"><img name="shox_viewhtmas_r3_c2" src="img/shox_view.htm.as_r3_c2.jpg" width="83" height="26" border="0"></a></td>
   <td rowspan="2" colspan="3"><img name="shox_viewhtmas_r3_c4" src="img/shox_view.htm.as_r3_c4.jpg" width="205" height="26" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="6" border="0"></td>
  </tr>
  <tr>
    <td colspan="2"><a href="index.asp"><img name="shox_viewhtmas_r4_c12" src="img/shox_view.htm.as_r4_c12.jpg" width="54" height="20" border="0"></a></td>
    <td rowspan="2" width="17"><img name="shox_viewhtmas_r4_c14" src="img/shox_view.htm.as_r4_c14.jpg" width="17" height="54" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="20" border="0"></td>
  </tr>
  <tr>
   <td colspan="4"><img name="shox_viewhtmas_r5_c1" src="img/shox_view.htm.as_r5_c1.jpg" width="149" height="34" border="0"></td>
    <td colspan="2" background="img/shox_view.htm.as_r5_c5.jpg"><b>&nbsp;<%=man%></b></td>
   <td colspan="4"><img name="shox_viewhtmas_r5_c7" src="img/shox_view.htm.as_r5_c7.jpg" width="148" height="34" border="0"></td>
    <td colspan="3"><a href="shox_paste.asp"><img name="shox_viewhtmas_r5_c11" src="img/shox_view.htm.as_r5_c11.jpg" width="89" height="34" border="0"></a></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="34" border="0"></td>
  </tr>
  <tr>
   <td colspan="4"><img name="shox_viewhtmas_r6_c1" src="img/shox_view.htm.as_r6_c1.jpg" width="149" height="35" border="0"></td>
    <td colspan="3" background="img/shox_view.htm.as_r6_c5.jpg"><font color="#FF9900"><%=sbj%></font></td>
   <td colspan="2"><img name="shox_viewhtmas_r6_c8" src="img/shox_view.htm.as_r6_c8.jpg" width="44" height="35" border="0"></td>
    <td colspan="4" background="img/shox_view.htm.as_r6_c10.jpg"><b><font size="2" color="#FFFFFF">&nbsp;&nbsp;<%=NowD%></font></b></td>
    <td width="17"><img name="shox_viewhtmas_r6_c14" src="img/shox_view.htm.as_r6_c14.jpg" width="17" height="35" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="35" border="0"></td>
  </tr>
  <tr>
   <td rowspan="5" colspan="2"><img name="shox_viewhtmas_r7_c1" src="img/shox_view.htm.as_r7_c1.jpg" width="64" height="285" border="0"></td>
   <td colspan="12"><img name="shox_viewhtmas_r7_c3" src="img/shox_view.htm.as_r7_c3.jpg" width="535" height="12" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="12" border="0"></td>
  </tr>
  <tr>
    <td width="197" rowspan="2" colspan="3" bgcolor="#CCCCCC" valign="top" background="img/shox_view.htm.as_r8_c3.jpg" height="247"> 
      <textarea name="textfield" class="box01" cols="22" rows="14" style="background-color: #CCCCCC; background: img/shox_view.htm.as_r8_c3.jpg;"><%=body%></textarea>
    </td>

<%
'這地方改為不管有無圖片  依順序抓圖如果fsPDps=1就是jpg  2是gif   3是無
FileNameJPG = Server.MapPath("ptr/" & ptr & ".jpg")
FileNameGIF = Server.MapPath("ptr/" & ptr & ".gif")
Set fsisps = CreateObject("Scripting.FileSystemObject")
IF fsisps.FileExists(FileNameJPG) Then
		fsPDps="ptr/" & ptr & ".jpg"
Else
		fsPDps="ptr/" & ptr & ".gif"
End IF
Set fsisps = nothing
%>
    <td colspan="7" height="203"><img name="shox_viewhtmas_r8_c6" src="<%=fsPDps%>" width="301" height="230" border="0"></td>
   <td rowspan="5" colspan="2"><img name="shox_viewhtmas_r8_c13" src="img/shox_view.htm.as_r8_c13.jpg" width="37" height="365" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="230" border="0"></td>
  </tr>
  <tr>
   <td rowspan="3" colspan="3"><img name="shox_viewhtmas_r9_c6" src="img/shox_view.htm.as_r9_c6.jpg" width="188" height="43" border="0"></td>
    <td rowspan="2" colspan="4" background="img/shox_view.htm.as_r9_c9.jpg"><font size="2"><%=dateDay%></font></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="17" border="0"></td>
  </tr>
  <tr>
   <td rowspan="2" colspan="3"><img name="shox_viewhtmas_r10_c3" src="img/shox_view.htm.as_r10_c3.jpg" width="197" height="26" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="15" border="0"></td>
  </tr>
  <tr>
   <td rowspan="2" colspan="4"><img name="shox_viewhtmas_r11_c9" src="img/shox_view.htm.as_r11_c9.jpg" width="113" height="103" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="11" border="0"></td>
  </tr>
  <tr>
   <td colspan="8"><img name="shox_viewhtmas_r12_c1" src="img/shox_view.htm.as_r12_c1.jpg" width="449" height="92" border="0"></td>
    <td width="1"><img src="img/spacer.gif" width="1" height="92" border="0"></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>