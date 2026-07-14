<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgStorageDefinitionInsert.asp
	'Module Name				:	Inventory (Storage Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 14, 2002
	'Modified By				:	Ragavendran R
	'Modified On				:	Jan 07,2011
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->

<SCRIPT LANGUAGE=javascript>
<!--
	function msgbox(strr,flag,lName,iBin) {
		if (flag == "Y") {
			alert(strr);
			if(iBin > 0) {
				//if (confirm("Do You want to define Bin Details for the " + lName + " Storage Location Created"))
					//window.location.href = "OrgStorageBinDetailsEntry.asp"
				//	window.location.href = "STORELOCATIONS.ASP"

				//else
					//window.location.href = "OrgStorageDefinitionEntry.asp"
					window.location.href = "STORELOCATIONS.ASP"
			}
			else
				//window.location.href = "OrgStorageDefinitionEntry.asp"
				window.location.href = "STORELOCATIONS.ASP"
		}
		else {
			alert(strr);
			window.history.back(1);
		}
	}
//-->
</SCRIPT>
<%
'XML DOM Variables
Dim oDOM,newElem,newElem1,Root,objfs,HeaderNode,bFlag

' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")


dim dcrs,iCount,sSql,imaxLoc,newElem2,sFlag
dim sorgID,sLocName,sLocCode,sApplicable,sStorageType,sStorageTypeE,sStorageTypeB,iFreeArea,inoBins
dim oDOM1,oDOM2,Root1,Root2,dcrs1,imaxBin,iVar,arrLoc,sBinCode,sBinName,sBinArea,objfs1
dim BinNode,sExp,i,iBinNo,iLocNo,objfs2
Set dcrs1= Server.CreateObject("ADODB.RecordSet")
Set objfs1 = CreateObject("Scripting.FileSystemObject")
Set objfs2 = CreateObject("Scripting.FileSystemObject")
Set oDOM1 = Server.CreateObject("Microsoft.XMLDOM")
Set oDOM2 = Server.CreateObject("Microsoft.XMLDOM")
'Response.Write "Chk="&Session.SessionID &"<br><BR>"
'sorgID = trim(Request.Form("selOrgUnit"))
sorgID = Session("organizationcode")
sLocName = trim(Request.Form("txtLocationName"))
sLocCode = trim(Request.Form("txtLocationCode"))
sApplicable = trim(Request.Form("App"))
sStorageType = trim(Request.Form("ST"))
sStorageTypeE = trim(Request.Form("ST"))
sStorageTypeB = trim(Request.Form("ST2"))
iFreeArea = trim(Request.Form("txtUsable"))
inoBins = trim(Request.Form("txtBins"))
sFlag = Request("hPara")
iLocNo = Request("hLocNo")
'Response.Write "Flag="&sFlag
'Response.Write "Testing ............."
'Response.End


	if sStorageType = "F" then
		sStorageTypeB = "0"
		inoBins = "0"
		sStorageTypeE = "1"
		iFreeArea = iFreeArea
	else
		sStorageTypeE = "0"
		iFreeArea = "0"
		sStorageTypeB = "1"
		inoBins = inoBins
	end if


	con.beginTrans
If trim(sFlag) = "" then

	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(LOCATIONNUMBER),0) + 1 FROM INV_M_STORAGE"
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing

	set imaxLoc = dcrs(0)
	if not dcrs.EOF then
		imaxLoc = imaxLoc
	end if
	dcrs.Close
	'IF sFlag = "" then
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT LOCATIONNAME FROM INV_M_STORAGE WHERE LOWER(LOCATIONNAME) = " & Pack(lcase(sLocName)) & " AND OUDEFINITIONID = " & Pack(sorgID) & ""
		.ActiveConnection = con
		.Open
	end with

	set dcrs.ActiveConnection = nothing
	'Response.Write dcrs.EOF
	if dcrs.EOF then

		sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
			"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
			"(" & Pack(sorgID) & "," & imaxLoc & "," & Pack(ucase(sLocCode)) & "," & Pack(ucase(sLocName)) & ", " &_
			" " & Pack(sApplicable) & "," & Pack(sStorageTypeE) & "," & Pack(sStorageTypeB) & "," & Pack(iFreeArea) & "," & inoBins & ")"
	'	Response.Write sSql & "<BR>"
		con.Execute sSql

		bFlag = true
		if objfs.FileExists(Server.MapPath("../xmldata/Storage.xml")) then
			oDOM.Load server.MapPath("../xmldata/Storage.xml")
			Set Root = oDOM.documentElement
			For Each HeaderNode In Root.childNodes
				if StrComp(HeaderNode.Attributes.Item(0).nodeValue,sorgID) = 0 then
					bFlag = false
					Set newElem = oDOM.createElement("Storage")
					newElem.setAttribute "LOCATIONNUMBER", imaxLoc
					newElem.setAttribute "LOCATIONCODE", ucase(sLocCode)
					newElem.setAttribute "LOCATIONNAME", ucase(sLocName)
					newElem.setAttribute "APPLICABLEFOR", sApplicable
					newElem.setAttribute "STORAGETYPEFREE", sStorageTypeE
					newElem.setAttribute "STORAGETYPEBINS", sStorageTypeB
					newElem.setAttribute "USABLEFREEAREA", iFreeArea
					newElem.setAttribute "NUMBEROFBINS", inoBins
					HeaderNode.appendChild newElem
				else
					'Response.Write "AAAA"

				end if
			next
			if bFlag then
				Set newElem1 = oDOM.createElement("Organization")
				newElem1.setAttribute "OUDEFINITIONID", sorgID
				Root.appendChild newElem1

				Set newElem = oDOM.createElement("Storage")
				newElem.setAttribute "LOCATIONNUMBER", imaxLoc
				newElem.setAttribute "LOCATIONCODE", ucase(sLocCode)
				newElem.setAttribute "LOCATIONNAME", ucase(sLocName)
				newElem.setAttribute "APPLICABLEFOR", sApplicable
				newElem.setAttribute "STORAGETYPEFREE", sStorageTypeE
				newElem.setAttribute "STORAGETYPEBINS", sStorageTypeB
				newElem.setAttribute "USABLEFREEAREA", iFreeArea
				newElem.setAttribute "NUMBEROFBINS", inoBins
				newElem1.appendChild newElem
			end if
		else
			Set Root = oDOM.createElement("Root")
			oDOM.appendChild Root
			Set newElem1 = oDOM.createElement("Organization")
			newElem1.setAttribute "OUDEFINITIONID", sorgID
			Root.appendChild newElem1

			Set newElem = oDOM.createElement("Storage")
			newElem.setAttribute "LOCATIONNUMBER", imaxLoc
			newElem.setAttribute "LOCATIONCODE", ucase(sLocCode)
			newElem.setAttribute "LOCATIONNAME", ucase(sLocName)
			newElem.setAttribute "APPLICABLEFOR", sApplicable
			newElem.setAttribute "STORAGETYPEFREE", sStorageTypeE
			newElem.setAttribute "STORAGETYPEBINS", sStorageTypeB
			newElem.setAttribute "USABLEFREEAREA", iFreeArea
			newElem.setAttribute "NUMBEROFBINS", inoBins
			newElem1.appendChild newElem
		end if
		'''''''Added newly on Jul 5th 2008''''''''''''''''''''''''''''''''''
		if objfs1.FileExists(server.MapPath("../temp/Master/StorageNew"&Session.SessionID&".xml")) then
			oDOM1.Load server.MapPath("../temp/Master/StorageNew"&Session.SessionID&".xml")
			'oDOM1.save server.MapPath("../temp/Master/StorageInsert"&Session.SessionID&".xml")


			Set Root1 = oDOM1.documentElement

			sExp = "//Organization/Storage[@LOCATIONNUMBER='0']/Bin"
			Set BinNode = Root1.SelectNodes(sExp)
			IF BinNode.length <> 0 then
			 'Response.Write " BinNode.length="& BinNode.length
				for i = 0 to BinNode.length -1
					iBinNo   = BinNode.item(i).Attributes.getNamedItem("BINNUMBER").value
					sBinCode = BinNode.item(i).Attributes.getNamedItem("BINCODE").value
					sBinName = BinNode.item(i).Attributes.getNamedItem("BINNAME").value
					sBinArea = BinNode.item(i).Attributes.getNamedItem("BINAREA").value
					'Response.Write "<p>"&iBinNo
					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT ISNULL(MAX(BINNUMBER),0) + 1 FROM Inv_M_StoreBinDetails"
						.ActiveConnection = con
						.Open
					end with
					set dcrs1.ActiveConnection = nothing
					set imaxBin = dcrs1(0)
					if not dcrs1.EOF then
						imaxBin = imaxBin
					end if
					dcrs1.Close

					with dcrs1
						.CursorLocation = 3
						.CursorType = 3
						.Source = "SELECT OUDEFINITIONID FROM Inv_M_StoreBinDetails WHERE OUDEFINITIONID = " & Pack(sorgID) & " AND LOCATIONNUMBER = " & sLocCode & " AND BINNUMBER = " & iBinNo & ""
						.ActiveConnection = con
						.Open
					end with
					set dcrs1.ActiveConnection = nothing
					if dcrs1.EOF then
						 sSql = "INSERT INTO Inv_M_StoreBinDetails (OUDEFINITIONID,LOCATIONNUMBER,BINNUMBER," &_
						 	"BINCODE,BINNAME,BINAREA) VALUES " &_
						 	"(" & Pack(sorgID) & "," & imaxLoc & "," & iBinNo & ", " &_
						 	" " & Pack(ucase(sBinCode)) & "," & Pack(ucase(sBinName)) & "," & Pack(sBinArea) & ")"
						 'Response.Write "<BR>"& sSql & "<BR>"
						 con.Execute sSql
					end if
					dcrs1.close
					Set newElem2 = oDOM.createElement("Bin")
					newElem2.setAttribute "BINNUMBER", iBinNo
					newElem2.setAttribute "BINCODE", ucase(sBinCode)
					newElem2.setAttribute "BINNAME", ucase(sBinName)
					newElem2.setAttribute "BINAREA", sBinArea
					newElem.appendChild newElem2
				next
			end if
		end if
	''''''''''''''''''''''''''''''''''
	oDOM.Save server.MapPath("../xmldata/Storage.xml")
	'Response.Clear
	'Response.Write "<p><b>Insert</b>"
	'Response.End
	%>
		<BODY BGCOLOR="#336699" onLoad = "msgbox('Storage Location <%=replace(sLocName,"'","\'")%> has been Created Successfully','Y','<%=replace(sLocName,"'","\'")%>','<%=inoBins%>')">
	<%
	else
	%>
		<BODY BGCOLOR="#336699" onLoad = "msgbox('Storage Location <%=replace(sLocName,"'","\'")%> already created','N','A','A')">
	<%
	End If
Else 'Update
	dim StNode,sSLName,sSLCode,OrgNode,RemNode,BinStNode,sExp1,iChkBinNo

	'Response.Write "<b>Session ID " & Session.SessionID &"</b><br><br>"

	if objfs2.FileExists(server.MapPath("../temp/Master/StorageNew"&Session.SessionID&".xml")) then
		oDOM2.Load server.MapPath("../temp/Master/StorageNew"&Session.SessionID&".xml")
	'	oDOM2.save server.MapPath("../temp/Master/StorageNewTest"&Session.SessionID&".xml")
		Set Root2 = oDOM2.documentElement

		sExp = "//Organization/Storage"
		Set StNode = Root2.SelectNodes(sExp)
		 'Response.Write StNode.length
		IF StNode.length <> 0 then
			sSql = "DELETE FROM INV_M_StoreBINDETAILS WHERE LOCATIONNUMBER = " & iLocNo & " AND OUDefinitionID = " & Pack(sorgID) & ""
				'Response.Write sSql
				con.Execute sSql
			sSql = "DELETE FROM INV_M_STORAGE WHERE LOCATIONNUMBER = " & iLocNo & " AND OUDEFINITIONID = " & Pack(sorgID) & ""
				'  Response.Write sSql & "<BR><BR>"
				con.Execute sSql
				'Response.Write "<p>"&iLocNo &"  =  "& trim(StNode.item(i).Attributes.getNamedItem("LOCATIONNUMBER").value) &"<BR>"
				sSLName			= StNode.item(0).Attributes.getNamedItem("LOCATIONNAME").value
				sSLCode			= StNode.item(0).Attributes.getNamedItem("LOCATIONCODE").value
				sApplicable		= StNode.item(0).Attributes.getNamedItem("APPLICABLEFOR").value
				sStorageTypeE	= StNode.item(0).Attributes.getNamedItem("STORAGETYPEFREE").Value
				sStorageTypeB	= StNode.item(0).Attributes.getNamedItem("STORAGETYPEBINS").Value
				iFreeArea		= StNode.item(0).Attributes.getNamedItem("USABLEFREEAREA").Value
				inoBins			= StNode.item(0).Attributes.getNamedItem("NUMBEROFBINS").Value
				IF iFreeArea = "" then iFreeArea = 0
				IF inoBins = "" then inoBins = 0

				'Response.Write "No Of Bins="& StNode.item(0).Attributes.getNamedItem("NUMBEROFBINS").Value &"<BR><BR>"


				sSql = "INSERT INTO INV_M_STORAGE (OUDEFINITIONID,LOCATIONNUMBER,LOCATIONCODE,LOCATIONNAME," &_
					"APPLICABLEFOR,STORAGETYPEFREE,STORAGETYPEBINS,USABLEFREEAREA,NUMBEROFBINS) VALUES " &_
					"(" & Pack(sorgID) & "," & iLocNo & "," & Pack(ucase(sSLCode)) & "," & Pack(ucase(sSLName)) & ", " &_
					" " & Pack(sApplicable) & "," & Pack(sStorageTypeE) & "," & Pack(sStorageTypeB) & "," & Pack(iFreeArea) & "," & inoBins & ")"
				 '	   Response.Write sSql & "<BR>"
					con.Execute sSql
				iChkBinNo = inoBins
		End IF ' IF StNode.length <> 0 then

		if objfs.FileExists(Server.MapPath("../xmldata/Storage.xml")) then
			oDOM.Load server.MapPath("../xmldata/Storage.xml")
			Set Root = oDOM.documentElement
			For Each HeaderNode In Root.childNodes
				For each StNode in HeaderNode.childnodes
					If trim(StNode.NodeName) = "Storage" then
						IF trim(StNode.getAttribute("LOCATIONNUMBER")) = iLocNo then
							Set RemNode = StNode
							HeaderNode.removechild RemNode
						End IF
					End If

				Next
				Set newElem = oDOM.createElement("Storage")
				newElem.setAttribute "LOCATIONNUMBER",iLocNo
				newElem.setAttribute "LOCATIONCODE",ucase(sSLCode)
				newElem.setAttribute "LOCATIONNAME",ucase(sSLName)
				newElem.setAttribute "APPLICABLEFOR",sApplicable
				newElem.setAttribute "STORAGETYPEFREE",sStorageTypeE
				newElem.setAttribute "STORAGETYPEBINS",sStorageTypeB
				newElem.setAttribute "USABLEFREEAREA",iFreeArea
				newElem.setAttribute "NUMBEROFBINS",inoBins
				HeaderNode.Appendchild newElem



				'Response.Write "No Of Bins="&iChkBinNo
				If trim(iChkBinNo) <> "0" then
					'Response.Write "iLocNo="&iLocNo
					sExp1 = "//Organization/Storage[@LOCATIONNUMBER = "& iLocNo &"]/Bin"
					Set BinNode = Root2.SelectNodes(sExp1)
					'Response.Write "<p>BinNode="&BinNode.length&"<br>"
					If BinNode.length <> 0 then
						For  i = 0 to BinNode.length -1
							iBinNo = BinNode.item(i).Attributes.getNamedItem("BINNUMBER").value
							sBinCode = BinNode.item(i).Attributes.getNamedItem("BINCODE").value
							sBinName = BinNode.item(i).Attributes.getNamedItem("BINNAME").value
							sBinArea = BinNode.item(i).Attributes.getNamedItem("BINAREA").value
							'Response.Write sBinName
							sSql = "INSERT INTO Inv_M_StoreBinDetails (OUDEFINITIONID,LOCATIONNUMBER,BINNUMBER," &_
							"BINCODE,BINNAME,BINAREA) VALUES " &_
							"(" & Pack(sorgID) & "," & iLocNo & "," & iBinNo & ", " &_
							" " & Pack(ucase(sBinCode)) & "," & Pack(ucase(sBinName)) & "," & Pack(sBinArea) & ")"
					'		   Response.Write "<p>" & sSql & "<BR>"
							con.Execute sSql

							Set newElem2 = oDOM.createElement("Bin")
							newElem2.setAttribute "BINNUMBER", iBinNo
							newElem2.setAttribute "BINCODE", ucase(sBinCode)
							newElem2.setAttribute "BINNAME", ucase(sBinName)
							newElem2.setAttribute "BINAREA", sBinArea
							newElem.appendChild newElem2
						Next
					End If
				End If
			Next
		end if

	end if
	 'Response.Write "<p><b>Update</b>"
	 ' Response.End
	'Response.Clear
 	oDOM.Save server.MapPath("../xmldata/Storage.xml")
	'Response.End
	%>
		<BODY BGCOLOR="#336699" onLoad = "msgbox('Storage Location <%=replace(sLocName,"'","\'")%> has been Amended Successfully','Y','<%=replace(sLocName,"'","\'")%>','<%=inoBins%>')">
	<%


End If 'If trim(sFlag) = "" then
if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
	'con.RollbackTrans
	 con.CommitTrans
end if

con.close
set con = nothing
%>
