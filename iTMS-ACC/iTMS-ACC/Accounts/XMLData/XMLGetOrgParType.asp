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
	'Program Name				:	XMLGetOrgParType.asp
	'Module Name				:	ACCOUNTS (Transcation)
	'Author Name				:	SENTHIL E
	'Created On					:	June 30, 2003
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	VouCAEntry.asp
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
<!-- #include File="../../include/Accpopulate.asp" -->
<%
dim objRs,objRs1,sQuery,OutData,Root,newElem
dim sorgID,iParSubType,sParType,sParTypeName
Dim sCallTy

Set OutData = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")
set objRs1 = Server.CreateObject("ADODB.Recordset")

sorgID = Request("orgID")
sCallTy = Request("sCallTy")


IF CStr(sCallTy) <> "P" Then
	sQuery="select distinct PartyType,PartySubType,SubTypeName from vwOrgPartyType where OUDefinitionID='"&sOrgId&"'"
Else
	sQuery = "select distinct PartyType,PartySubType,SubTypeName from vwOrgPartyType where OUDefinitionID='"&sOrgId&"' "&_
			 "and ((PartyType = 'CR' and PartySubType >= 2) or PartyType = 'DR') "
End IF

	With objRs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	End with

	Set objRs.Activeconnection = nothing

	set sParType=objRs(0)
	set iParSubType = objRs(1)


	if not objRs.EOF then
		Set Root = OutData.createElement("Root")
		OutData.appendChild Root
		Do while not objRs.EOF
			sParTypeName = Replace(objRs(2)	,"&"," and ")
			sQuery="select count(1) from APP_R_OrgParty where PartyType='"&sParType&"' and PartySubType="&iParSubType&" and OUDefinitionID='"&sOrgId&"'"
			with objRs1
				.CursorLocation =3
				.CursorType =3
				.Source = sQuery
				.ActiveConnection = con
				.Open
			end with
			set objRs1.ActiveConnection=nothing
			if CDbl(objRs1(0)) then
				Set newElem = OutData.createElement("PartyType")
					newElem.setAttribute "ParType",sParType&"?"&iParSubType
					'newElem.setAttribute "ParSubType", iParSubType
					newElem.text= sParTypeName					
					Root.appendChild newElem
			end if			
			objRs1.Close
		  	objRs.MoveNext		  	
		Loop				
		Response.ContentType="text/xml"
		Response.Write OutData.xml
	end if
	objRs.Close

	set objRs=nothing
	set objRs1=nothing

%>

