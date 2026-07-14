<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	glHeadUnitUpdate.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	SENTHIL E
	'Created On					:	December 10, 2002
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
<!--#include virtual="/include/Accpopulate.asp"-->
<%

'XML DOM Variables
Dim oDOM,nodHeader,Root,nodUnit,iUnitId,iRecCount,objfs
dim sBookValue,sTemp,newElem,iCounter,sName,sShortName,objRs
dim sParTypevalue
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objfs = CreateObject("Scripting.FileSystemObject")

oDOM.Load server.MapPath("../temp/master/glEntry_"&Session.SessionID&".xml")


Set Root = oDOM.documentElement

Set newElem = oDOM.createElement("OpeningMonthYear")
newElem.Text = getFromFinYear
Root.appendChild newElem

Set newElem = oDOM.createElement("ClosingMonthYear")
newElem.Text = getToFinYear
Root.appendChild newElem

Set newElem = oDOM.createElement("CreatedBy")
newElem.Text = getUserid()
Root.appendChild newElem

Set newElem = oDOM.createElement("ApprovedBy")
newElem.Text = getUserid()
Root.appendChild newElem
			
For Each nodHeader In Root.childNodes
	if StrComp(nodHeader.nodeName,"Description") = 0 then
		sName=nodHeader.Text
	end if
	if StrComp(nodHeader.nodeName,"ShortName") = 0 then
		sShortName=nodHeader.Text
	end if

	if StrComp(nodHeader.nodeName,"SubLedger") = 0 then
		if nodHeader.Attributes.Item(0).nodeValue=1 then
			sParTypevalue=Split(Request.Form("selPartyType"),",")
			
			for iCounter=0 to UBound(sParTypevalue)
				sTemp=Split(trim(sParTypevalue(iCounter)),"?")
			
				Set newElem = oDOM.createElement("ParType")
					newElem.setAttribute "UnitId", sTemp(0)
					newElem.setAttribute "PartyType", sTemp(1)
					newElem.setAttribute "PartySubType", sTemp(2)
					nodHeader.appendChild newElem
			next

		end if
	end if
	
	if StrComp(nodHeader.nodeName,"Units") = 0 then
			for Each nodUnit in nodHeader.childNodes
				iUnitId =trim(nodUnit.Attributes.Item(0).nodeValue)
				
				Set newElem  = oDOM.createAttribute("OpBalance")
				newElem.value = Request.Form("txtOpenBal"&iUnitId)
				nodUnit.setAttributeNode(newElem)
				Set newElem  = oDOM.createAttribute("OpCRDR")
				newElem.value = Request.Form("optOpenCD"&iUnitId)
				nodUnit.setAttributeNode(newElem)

			next	
	end if
next
oDOM.Save server.MapPath("../temp/master/glEntry_"&Session.SessionID&".xml")


sQuery="select count(1)  from Acc_M_GLAccountHead where upper(AccountHeadCode)='"&UCase(sShortName)&"'"


with objRs	
	.CursorLocation = 3
	.CursorType = 3
	.Source = sQuery
	.ActiveConnection = con
	.Open
end with
iRecCount=objRs(0)
objRs.Close
%>

<HTML>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			if (confirm("Do You want to Create GL Head")) 
				window.location.href = "glCreationMain.asp";
			else
				window.location.href = "../AccountsHome.asp";
		}
		else {
			alert(strr);
			window.location.href = "glCreationMain.asp";
		}
	}
//-->
</SCRIPT>
<%
if iRecCount <>0 then%>		
	<BODY onLoad = "msgbox('GL Head Already Exist','N')">
<%
else
	Dim adoConn,sQuery
   Set adoConn = Server.CreateObject("ADODB.Connection")
   adoConn.ConnectionString = con
   adoConn.CursorLocation = 3
   adoConn.Open

   sQuery = "Proc_GenGlHead"
   
   
   Dim adoCmd
   Set adoCmd = Server.CreateObject("ADODB.Command")
   Set adoCmd.ActiveConnection = adoConn
   adoCmd.CommandText = sQuery
   adoCmd.CommandType = 4 'adCmdStoredProc

   adoCmd.Parameters.Append adoCmd.CreateParameter("@XMLDoc",201,1,len(oDOM.xml),oDOM.xml)

   Dim adoRS
   Set adoRS = adoCmd.Execute()

	

%>
	<BODY onLoad = "msgbox('GL Head Created Successfully','Y')">
<%
end if	
objfs.DeleteFile(server.MapPath("../temp/master/glEntry_"&Session.SessionID&".xml"))

set objRs=nothing
set objfs=nothing
	
%>
</BODY>
<HTML>
