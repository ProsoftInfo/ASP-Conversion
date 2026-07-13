<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgStorageBinDetailsInsert.asp	
	'Module Name				:	Inventory (Storage Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 15, 2002
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag) {
		if (flag == "Y") {
			alert(strr);
			window.location.href = "OrgStorageDefinitionEntry.asp"
		}
	}
//-->
</SCRIPT>
<%
'XML DOM Variables
Dim oDOM,newElem,newElem1,newElem2,Root,objfs,HeaderNode,StoreNode

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")

dim dcrs,iCount,sSql,imaxBin,iVar
dim sorgID,arrLoc,sLocCode,inoBins,sBinCode,sBinName,sBinArea

sorgID = trim(Request.Form("selOrgUnit"))
arrLoc = split(trim(Request.Form("selLocName")),"|")
sLocCode = trim(arrLoc(0))
inoBins = trim(arrLoc(1))

Set dcrs = Server.CreateObject("ADODB.RecordSet")

con.beginTrans

for iVar = 1 to cint(inoBins)
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(BINNUMBER),0) + 1 FROM INV_M_ORGSLBINDETAILS"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	set imaxBin = dcrs(0)
	if not dcrs.EOF then
		imaxBin = imaxBin
	end if
	dcrs.Close

	sBinCode = trim(Request.Form(cstr(iVar)&"txtBinCode"))
	sBinName = trim(Request.Form(cstr(iVar)&"txtBinName"))
	sBinArea = trim(Request.Form(cstr(iVar)&"txtBinSize"))

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT OUDEFINITIONID FROM INV_M_ORGSLBINDETAILS WHERE OUDEFINITIONID = " & Pack(sorgID) & " AND LOCATIONNUMBER = " & sLocCode & " AND BINNUMBER = " & imaxBin & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if dcrs.EOF then
		sSql = "INSERT INTO INV_M_ORGSLBINDETAILS (OUDEFINITIONID,LOCATIONNUMBER,BINNUMBER," &_
			"BINCODE,BINNAME,BINAREA) VALUES " &_
			"(" & Pack(sorgID) & "," & sLocCode & "," & imaxBin & ", " &_
			" " & Pack(ucase(sBinCode)) & "," & Pack(ucase(sBinName)) & "," & Pack(sBinArea) & ")"
		'Response.Write sSql & "<BR>"
		con.Execute sSql

		if objfs.FileExists(Server.MapPath("../xmldata/Storagxe.xml")) then
			oDOM.Load server.MapPath("../xmldata/Storage.xml")
			Set Root = oDOM.documentElement
			For Each HeaderNode In Root.childNodes
				if StrComp(HeaderNode.Attributes.Item(0).nodeValue,sorgID) = 0 then
					For Each StoreNode In HeaderNode.childNodes
						if StrComp(StoreNode.Attributes.Item(0).nodeValue,sLocCode) = 0  then
							Set newElem = oDOM.createElement("Bin")
							newElem.setAttribute "BINNUMBER", imaxBin
							newElem.setAttribute "BINCODE", ucase(sBinCode)
							newElem.setAttribute "BINNAME", ucase(sBinName)
							newElem.setAttribute "BINAREA", sBinArea
							StoreNode.appendChild newElem					
						end if
					next
				end if
			next		
		end if
		oDOM.Save server.MapPath("../xmldata/Storage.xml")
	end if
	dcrs.Close
next	
%>
	<BODY onLoad = "msgbox('Bin Details has been Created Successfully','Y')">
<%
if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	if objfs.FileExists(Server.MapPath("../temp/master/STORAGEDEF"&Session.SessionID&".xml")) then
		objfs.DeleteFile server.MapPath("../temp/master/STORAGEDEF"&Session.SessionID&".xml")
	end if

	'con.RollbackTrans
	con.CommitTrans
end if

con.close
set con = nothing
%>
