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
	'Program Name				:	OrgGLHeadXML.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Senthil E
	'Created On					:	January 21,2003
	'Modified On				: 
	'Tables Used				: 
	'Temporary Tables			: 
	'Temporary Files			: 
	'Input Parameter			:	None
	'							:
	'Connects To				:	ContraEntry.asp	
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
	dim OutData,Root,newElem,newElem1
	dim objRs,objRs1,sQuery,sorgID,iBookNo,Objrs2
	
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set objRs = Server.CreateObject("ADODB.Recordset")
	set objRs1 = Server.CreateObject("ADODB.Recordset")
	
	
	sorgID = Request("orgID")
	
	sQuery="select a.AccountHead,b.AccountHeadCode,b.AccountDescription from Acc_R_OrgGLAccountHead a,Acc_M_GLAccountHead b "&_
			"where a.OUDefinitionID='"&sorgID&"' and a.EligibleForContras=1 and a.AccountHead=b.AccountHead and a.SubLedger=0 "&_
			"and A.AmendmentExists = '0' Order By b.AccountDescription "
			
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
				
				Set newElem = OutData.createElement("Head")
				newElem.setAttribute "AccountNo", trim(objRs(0))
				newElem.setAttribute "AccountCode",trim(objRs(1))
				newElem.setAttribute "AccountDesc", trim(objRs(2))
				Root.appendChild newElem
				
				
				sQuery="select a.AccountHead,b.AccountHeadCode,b.AccountDescription from Acc_R_OrgGLAccountHead a,Acc_M_GLAccountHead b "&_
					"where a.OUDefinitionID='"&sorgID&"' and a.EligibleForContras=1 and a.AccountHead=b.AccountHead and a.SubLedger=0 "&_
					"and  a.AccountHead not in(select ToAccountHead from Acc_M_ContraEntries where  a.OUDefinitionID='"&sorgID&"' "&_
					" and FromAccountHead="&trim(objRs(0))&")"&_
					" and a.AccountHead not in(select FromAccountHead from Acc_M_ContraEntries where  a.OUDefinitionID='"&sorgID&"' "&_
					" and ToAccountHead="&trim(objRs(0))&")"&_
					" and a.AccountHead<>"&trim(objRs(0))
				
				sQuery="select a.AccountHead,b.AccountHeadCode,b.AccountDescription from Acc_R_OrgGLAccountHead a,Acc_M_GLAccountHead b "&_
					" where a.OUDefinitionID='"&sorgID&"' and a.EligibleForContras=1 and a.AccountHead=b.AccountHead and a.SubLedger=0 "&_
					" and  a.AccountHead Not in(select ToAccountHead from Acc_M_ContraEntries where  a.OUDefinitionID='"&sorgID&"' "&_
					" and FromAccountHead="&trim(objRs(0))&") and a.AccountHead<>"&trim(objRs(0)) &" "
					
				
				with objRs1
					.CursorLocation = 3
					.CursorType = 3
					.Source = sQuery
					.ActiveConnection = con
					.Open
				end with
	
				set objRs1.ActiveConnection = nothing
				if not objRs1.EOF then
					
					do while not objRs1.EOF
						Set newElem1 = OutData.createElement("ToHead")
							newElem1.setAttribute "AccountNo", trim(objRs1(0))
							newElem1.setAttribute "AccountCode",trim(objRs1(1))
							newElem1.setAttribute "AccountDesc", trim(objRs1(2))
							newElem.appendChild newElem1
						
						objRs1.MoveNext
					loop
				end if
				objRs1.Close	
				
		objRs.MoveNext
		loop

		Response.ContentType="text/xml"
		Response.Write OutData.xml
	end if
	objRs.Close
	
	
%>
