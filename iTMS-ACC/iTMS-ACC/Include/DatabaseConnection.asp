<!-- #include File="session.asp"-->
<%
'Response.ExpiresAbsolute=#Now#
'Response.Expires=10
'Response.AddHeader "pragma","no-cache"
'Response.AddHeader "cache-control","private"
'Response.CacheControl = "no-cache"
%>
<%
	'on error resume next
	dim con,connFile,connString,connPath,connArray
	redim connArray(3)
	const connReading = 1
	set con = Server.CreateObject("ADODB.Connection")
	set connFile = Server.CreateObject("Scripting.FileSystemObject")
	connPath = Server.MapPath("/include/Settings.inf")
	set connFile = connFile.OpenTextFile(connPath,connReading,true)
	connFile.ReadLine()
	connString = connFile.ReadLine()
	connArray = split(connString,":")
	con.Open connArray(0),connArray(1),connArray(2)
	'if con.Errors.count <> 0 then
	'	Response.Redirect "../../welcome.html"
	'end if
%>

