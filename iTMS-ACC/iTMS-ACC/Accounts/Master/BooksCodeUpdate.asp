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
	'Program Name				:	BooksCodeUpdate.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Senthil E
	'Created On					:	December 31, 2002
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<%
	dim dcrs,sSql,OutData,Root,newElem
	dim objRs,objRs1,sQuery,sorgID,iBookNo,iGlHead,iBookid
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs1 = Server.CreateObject("ADODB.Recordset")
	
	sorgID = Request("orgID")
	iBookNo= Request("BookNo")
	iBookid= Request("Bookid")
	iGlHead= Request("GlHead")
	
	sQuery="update Acc_R_ApplicableAccountHeads set BookAccountHead="&iGlHead&" where  OUDefinitionID='"&sorgID&"'"&_
			" and BookCode="&iBookid&" and BookNumber="&iBookNo
	
	con.Execute (sQuery)	
	
	sQuery=" select BookNumber,BookName,isnull(BookAccountHead,0) from Acc_R_ApplicableAccountHeads Where"&_
			" OUDefinitionID='"&sorgID&"' and BookCode="&iBookid&" and Useable = '0' "

	with objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	
	set objRs.ActiveConnection = nothing
	
	Set Root = OutData.createElement("Root")												
	OutData.appendChild Root

	if not objRs.EOF then

		do while not objRs.EOF
			if cint(objRs(2)) = 0  then
				Set newElem = OutData.createElement("Book")
				newElem.setAttribute "BookNo", trim(objRs(0))
				newElem.setAttribute "BookName",trim(objRs(1))
				newElem.setAttribute "GlCode", trim(objRs(2))
				newElem.setAttribute "GlName", ""
				newElem.setAttribute "GlShortName",""
				newElem.setAttribute "TransCount", "0"
				Root.appendChild newElem
				
			else
				sQuery="select AccountDescription,AccountHeadCode from Acc_M_GLAccountHead where AccountHead="&objRs(2)
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
				set objRs1.ActiveConnection = nothing
	
				Set newElem = OutData.createElement("Book")
				newElem.setAttribute "BookNo", trim(objRs(0))
				newElem.setAttribute "BookName",trim(objRs(1))
				newElem.setAttribute "GlCode", trim(objRs(2))
				newElem.setAttribute "GlName", trim(objRs1(0))
				newElem.setAttribute "GlShortName", trim(objRs1(1))
				newElem.setAttribute "TransCount", "0"
				Root.appendChild newElem
				
				objRs1.Close
			end if

		objRs.MoveNext
		loop

		Response.ContentType="text/xml"
		Response.Write OutData.xml
	end if
	objRs.Close
%>
