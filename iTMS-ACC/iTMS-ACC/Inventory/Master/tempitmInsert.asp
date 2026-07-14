<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	tempitmInsert.asp
	'Module Name				:	Inventory (Temporary Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	January 08, 2004
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
<!--#include virtual="/include/populate.asp"-->
<%
dim oDOM,RootNode,HeaderNode,OutData,ItemNode,DeleteNode,UpdateNode,SessionNode
dim newElem,newElem1,newElem2,dCreatedDate,oNode
dim sitmCode,sitmShDesc,sitmDesc,sitmAddDesc,sitmType
dim iAppCode,iModCode,sCreationStage,sExistitmCode,sExp,sExp1,iCtr

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

oDOM.async = false
oDOM.load(Request)

Set RootNode = oDOM.documentElement

'sitmType = trim(RootNode.Attributes.getNamedItem("ITMTYPE").Value)
iAppCode = trim(RootNode.Attributes.getNamedItem("APPCODE").Value)
iModCode = trim(RootNode.Attributes.getNamedItem("MODCODE").Value)
sCreationStage = trim(RootNode.Attributes.getNamedItem("CRESTAGE").Value)

For Each HeaderNode In RootNode.childNodes
	sitmCode = trim(HeaderNode.Attributes.getNamedItem("ITMCODE").Value)
	sitmDesc = trim(HeaderNode.Attributes.getNamedItem("ITMDESC").Value)
	sitmShDesc = trim(HeaderNode.Attributes.getNamedItem("ITMSHDESC").Value)
	sitmAddDesc = trim(HeaderNode.Attributes.getNamedItem("ITMADDDESC").Value)
next

set RootNode = nothing
set oDOM = nothing
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

if oDOM.Load(server.MapPath("../xmldata/TEMPORARYITEM.xml")) then
	Set RootNode = oDOM.documentElement

	sExp="//ITEMCODE"
	set ItemNode = RootNode.Selectnodes(sExp)
	For iCtr = 0 to ItemNode.Length - 1
		sExistitmCode = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("ITMCODE").Value)
		dCreatedDate = trim(ItemNode.Item(iCtr).Attributes.getNamedItem("CREATEDDATE").Value)

		if DateDiff("d",date(),FormatDate(dCreatedDate)) < 0 then
			sExp1 ="//ITEMCODE [ @ITMCODE = '"&sExistitmCode&"']"
			Set DeleteNode = RootNode.Selectnodes(sExp1)

			if DeleteNode.Length > 0 then
				Set oNode = RootNode.RemoveChild(DeleteNode.Item(0))
			end if
		end if
	next

	sExp ="//ITEMCODE [ @ITMCODE = '"&sitmCode&"']"
	Set ItemNode = RootNode.Selectnodes(sExp)
	if 	ItemNode.Length > 0 then
		sExp1 ="//ITEMCODE [ @ITMCODE = '"&sitmCode&"']/SESSIONDETAILS [ @SESSIONID = '"&Session.SessionID&"']"
		Set SessionNode = RootNode.Selectnodes(sExp1)
		if SessionNode.Length > 0 then
			set UpdateNode = SessionNode.Item(0).childNodes(0)
			UpdateNode.setAttribute "ITMTYPE",""' sitmType
			UpdateNode.setAttribute "APPCODE", iAppCode
			UpdateNode.setAttribute "MODCODE", iModCode
			UpdateNode.setAttribute "CRESTAGE", sCreationStage
			UpdateNode.setAttribute "ITMDESC", sitmDesc
			UpdateNode.setAttribute "ITMSHDESC", sitmShDesc
			UpdateNode.setAttribute "ITMADDDESC", sitmAddDesc
		else
			' Call Function to Insert the Details into XML for Body
			AddDetails "B"
		end if
	else
		' Call Function to Insert the Details into XML for ALL
		AddDetails "A"
	end if
else
	Set RootNode = oDOM.createElement("ITEMNODE")
	oDOM.appendChild RootNode

	' Call Function to Insert the Details into XML for ALL
	AddDetails "A"
end if

oDOM.save (server.MapPath("../xmldata/TEMPORARYITEM.xml"))
%>

<%
	Function AddDetails(sWhatFor)
		'All
		if sWhatFor = "A" then
			Set newElem = oDOM.createElement("ITEMCODE")
			newElem.setAttribute "ITMCODE", sitmCode
			newElem.setAttribute "CREATEDDATE", FormatDate(date())

			Set newElem1 = oDOM.createElement("SESSIONDETAILS")
			newElem1.setAttribute "SESSIONID", Session.SessionID

			Set newElem2 = oDOM.createElement("ITEMDETAILS")
			newElem2.setAttribute "ITMTYPE", sitmType
			newElem2.setAttribute "APPCODE", iAppCode
			newElem2.setAttribute "MODCODE", iModCode
			newElem2.setAttribute "CRESTAGE", sCreationStage
			newElem2.setAttribute "ITMDESC", sitmDesc
			newElem2.setAttribute "ITMSHDESC", sitmShDesc
			newElem2.setAttribute "ITMADDDESC", sitmAddDesc

			newElem1.appendChild newElem2

			newElem.appendChild newElem1

			RootNode.appendChild newElem

		'Body
		elseif sWhatFor = "B" then
			Set newElem1 = oDOM.createElement("SESSIONDETAILS")
			newElem1.setAttribute "SESSIONID", Session.SessionID

			Set newElem2 = oDOM.createElement("ITEMDETAILS")
			newElem2.setAttribute "ITMTYPE", sitmType
			newElem2.setAttribute "APPCODE", iAppCode
			newElem2.setAttribute "MODCODE", iModCode
			newElem2.setAttribute "CRESTAGE", sCreationStage
			newElem2.setAttribute "ITMDESC", sitmDesc
			newElem2.setAttribute "ITMSHDESC", sitmShDesc
			newElem2.setAttribute "ITMADDDESC", sitmAddDesc

			newElem1.appendChild newElem2

			ItemNode.Item(0).appendChild newElem1

		end if
	end Function
%>