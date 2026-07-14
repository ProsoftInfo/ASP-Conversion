<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParGetUnitDetails.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 20,2010
	'Modified On				: 
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/accpopulate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
'XML DOM Variables
Dim oDOM,nodHeader,Root,nodUnit,iUnitId,objDom,Mainnode,sExp,UnitRoot,objRs
dim sBookValue,sTemp,newElem,iCounter,sGrType,bGrFlag,sQuery
dim sParValue,arrParType,bAgentFlag,sParCheck,sPararr,iCount,iPartyCode
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objDom = Server.CreateObject("Microsoft.XMLDOM")
set objRs = Server.CreateObject("ADODB.Recordset")

oDOM.Load server.MapPath("../temp/master/Party_Master_"&Session.SessionID&".xml")
Set UnitRoot = objDom.createElement("Root")
objDom.appendChild UnitRoot
bAgentFlag=false
'Response.Write oDOM.xml
Set Root = oDOM.documentElement
if Root.hasChildNodes() then
	For Each nodHeader In Root.childNodes
		if nodHeader.nodeName="ParCode" then
			iPartyCode= nodHeader.text
		end if
		if StrComp(nodHeader.nodeName,"Group") = 0 then
			bGrFlag=trim(nodHeader.Attributes.Item(0).nodeValue)
			sGrType=trim(nodHeader.Attributes.Item(1).nodeValue)
		end if
		if StrComp(nodHeader.nodeName,"Units") = 0 then
			for Each nodUnit in nodHeader.childNodes
				iUnitId =trim(nodUnit.Attributes.Item(0).nodeValue)
				'Response.Write "iPartyCode="& iPartyCode
				if iPartyCode<>"" then
					sQuery= "Select PartyType,PartySubType,OpeningAmount,OpeningCDIndication,OpeningMonthYear,"&_
							"ClosingMonthYear from Acc_T_PartyOpeningAmt where PartyCode ="& iPartyCode &" and OUDefinitionID = '"& iUnitId &"'"
					objRs.Open sQuery,con
					if not objRs.EOF then
						do while not objRs.EOF 
							Set newElem = objDom.createElement("Partytype")
								newElem.setAttribute "Type", objRs(0)
								newElem.setAttribute "SubType", objRs(1)
								newElem.setAttribute "OpBalance", objRs(2)
								newElem.setAttribute "OpCRDR", objRs(3)
								newElem.setAttribute "OpeningMonthYear",  objRs(4)
								newElem.setAttribute "ClosingMonthYear",  objRs(5)
								newElem.setAttribute "Unit",iUnitId
								UnitRoot.appendChild newElem
								if trim(objRs(3))="C" and cint(objRs(1))< 3 then
									bAgentFlag=true
								end if
							objRs.MoveNext 
						loop
					end if
					objRs.Close 
				end if	 'if iPartyCode<>"" then
			next	
		end if
	next
end if 'if Root.hasChildNodes() then
Set newElem = objDom.createElement("Agent")
UnitRoot.appendChild newElem
if bAgentFlag then
	newElem.setAttribute "Flag",1
else
	newElem.setAttribute "Flag",0
end if	
Response.ContentType ="text/xml"
Response.Write objDom.xml
%>