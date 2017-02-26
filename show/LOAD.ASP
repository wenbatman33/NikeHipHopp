<%@ENABLESESSIONSTATE=FALSE%>
<%
set Upload = Server.CreateObject("Dundas.Upload")
set NextFile = Upload.GetNextFile
''''imgname = mid(NextFile.Filename,instr(1,NextFile.Filename,"}_") + 2 ,len(NextFile.Filename) - instr(1,NextFile.Filename,"}_"))

imgname="123.gif"

NextFile.Filename = imgname
NextFile.Save server.MapPath("/hip_hoop_2002/show/")
%>
<%=imgname%>