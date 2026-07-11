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
<SCRIPT>
var sFilePath = "";

function Init() {
	var fileField = document.getElementById("printFilePath");
	var frame = document.getElementById("printFrame");
	sFilePath = fileField ? fileField.value : "";
	if (frame && sFilePath) {
		frame.src = sFilePath;
	}
}

function cleanupPrintFile() {
	var url;
	if (!sFilePath) {
		return;
	}
	url = "DeletePrintFile.asp?sPath=" + encodeURIComponent(sFilePath);
	if (navigator.sendBeacon) {
		navigator.sendBeacon(url);
		return;
	}
	try {
		fetch(url, { method: "GET", keepalive: true, cache: "no-cache" });
	} catch (ignore) {}
}

function PrintDocument() {
	var frame = document.getElementById("printFrame");
	if (frame && frame.contentWindow) {
		frame.contentWindow.focus();
		frame.contentWindow.print();
	} else {
		window.print();
	}
}

window.addEventListener("beforeunload", cleanupPrintFile);
</SCRIPT>
<BODY BGCOLOR="#CCCCCC" LINK="#0000FF" VLINK="#800080" TEXT="#000000" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 ONLOAD="Init()">
	<input type="hidden" id="printFilePath" value="<%=Server.HTMLEncode(filepath)%>">
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
				<input type="button" value="Print" onclick="PrintDocument()">
				<input type="button" value="Open Document" onclick="window.open(sFilePath,'_blank')">
				<br><br>
				<iframe id="printFrame" title="Print document" style="width:760px;height:420px;background:#fff;border:1px solid #999;"></iframe>
            </TD>
        </TR>

      </table>
      </center>
    </div>
</BODY>
</HTML>

