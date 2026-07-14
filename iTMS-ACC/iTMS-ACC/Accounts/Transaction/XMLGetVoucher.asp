<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLGetVoucher.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	MANOHAR PRABHU.R
	'Created On					:	June 10, 2005
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:
	'Procedures/Functions Used	:
	'Internal Variables			:
	'Database					:
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>
<!--#include virtual="/include/Accpopulate.asp"-->
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
	Dim sTransNo,oDom,sRetVal
	
	Set oDom = Server.CreateObject("Microsoft.XMLDOM")
	sTransno = Request("TransNo")
	
	sRetVal = GetVouchXML(sTransno)
	oDOM.Load server.MapPath(sRetVal)
	
	'oDOM.load server.MapPath("../xmldata/Voucher/"&sTransNo&".xml")
	'oDOM.Save server.MapPath("../temp/transaction/Voucher AMD_CA_"&Session.SessionID&".xml")
	Response.ContentType="text/xml"
	Response.Write oDom.xml
		
%>