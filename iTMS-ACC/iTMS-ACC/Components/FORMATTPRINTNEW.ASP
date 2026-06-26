<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
dim servername,filepath,exitpath,targetframe,serverport

'Response.Redirect("FormattPrint.asp?server=server&filepath=/Reports/"&Session("userid")&"-UcRegister.txt&exitpath=/asp/body_usermenu.asp&frame=_parent")

servername=Request.ServerVariables("SERVER_NAME")
serverport=Request.ServerVariables("SERVER_PORT")
filepath=Request.QueryString("filepath")
exitpath=Request.QueryString("exitpath")
targetframe=Request.QueryString("frame")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<META NAME="Generator" CONTENT="Microsoft FrontPage 4.0">
<TITLE>iTMS - Dot Matrix Printing</TITLE>
</HEAD>
<SCRIPT LANGUAGE=vbscript>
dim sFilePath

Function Init(sPara)
	sFilePath = sPara
end Function

Function window_onunload()
	set objhttp = CreateObject("MSXML2.XMLHTTP")

	objhttp.Open "GET","DeletePrintFile.asp?sPath="&sFilePath, false

	objhttp.send
end Function

</SCRIPT>
<BODY BGCOLOR="#CCCCCC" LINK="#0000FF" VLINK="#800080" TEXT="#000000" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 ONLOAD="Init('<%=filepath%>')">
    <div align="center">
      <center>
    <TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 NOF=LY>
        <TR VALIGN=TOP ALIGN=LEFT>
            <TD height="10"></TD>
        </TR>
        <TR VALIGN=TOP ALIGN=LEFT>
            <TD><P ALIGN=CENTER><B><FONT COLOR="#660000" SIZE="-1" FACE="Arial,Helvetica,Univers,Zurich BT">Document Printing</FONT></B></TD>
        </TR>
        <TR VALIGN=TOP ALIGN=LEFT>
            <TD HEIGHT=10></TD>
        </TR>
        <TR VALIGN=TOP ALIGN=LEFT>
            <TD HEIGHT=10>
              <p align="center">
<OBJECT id=WebRawPrint style="LEFT: 0px; TOP: 0px"  codebase="WebRawPrintingNew.CAB#version=1,0,0,12"
	classid="CLSID:DAA4EDBF-2C42-4C60-A867-970ADAB06E00" VIEWASTEXT width="237" height="67">
<PARAM NAME="_ExtentX" VALUE="6271">
	<PARAM NAME="_ExtentY" VALUE="1773">
	<PARAM NAME="BackColor" VALUE="14737632">
	<PARAM NAME="Enabled" VALUE="-1">
	<PARAM NAME="filename" VALUE="<%=filepath%>">
	<PARAM NAME="server" VALUE="<%=servername%>">
	<PARAM NAME="serverport" VALUE="<%=serverport%>">
	<PARAM NAME="exitpath" VALUE="<%=exitpath%>">
	<PARAM NAME="targetframe" VALUE="<%=targetframe%>">
	</OBJECT>
            </TD>
        </TR>

      </table>
      </center>
    </div>
</BODY>
</HTML>

