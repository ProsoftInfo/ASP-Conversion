<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ClassDeletionDetailsEntry.asp
	'Module Name				:	Inventory (Classification Deletion)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	March 09, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	ClassDeletionUpdate.asp
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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Classification Deletion - Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT>
function DeleteClass() {
	var itemData;
	var payload;
	var request;
	if (document.formname.hDelete.value === "N") {
		return false;
	}
	if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
		window.ITMSModernCompat.init(document);
	}
	itemData = window.ItemData || document.ItemData;
	payload = itemData && (itemData.XMLDocument || itemData._doc || itemData);
	request = new XMLHttpRequest();
	request.open("POST", "ClassDeletionUpdate.asp", false);
	request.send(payload);
	if (request.responseText === "") {
		alert("Classification has been deleted Successfully.");
		window.parent.location.href = "MasClassificationEntry.asp";
	} else {
		alert(request.responseText);
	}
	return true;
}
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<% 	
	dim dcrs,dcrs1,sWho,OutData,RootNode,newElem1,newElem2,sPGroup,sOrgName,DeleteNode
	dim iClass,sClassName,arrTemp,sCat,sCatName,sTemp,iItem,sItemName,iCount
	dim iCtr,iCtr1,sExp,sExp1,ClassNode,ItemNode,iCounter,sDeleteName,arrDelete
	iCounter = 0
	sTemp = trim(Request.Form("pGroup"))
	
	arrTemp = split(sTemp,":")
	iCount = cint(UBound(arrTemp))

	if iCount < 0 then Response.End

	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")

	Set RootNode = OutData.createElement("DETAILS")
	
	iClass = arrTemp(iCount)
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT GROUPCODE,GROUPNAME,PARENTGROUP FROM INV_M_CLASSIFICATION WHERE GROUPCODE = " & iClass & " ORDER BY GROUPCODE"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	Do While Not dcrs.EOF

		iClass = trim(dcrs(0))
		sClassName = trim(dcrs(1))
		sPGroup = trim(dcrs(2))

		set newElem1 = OutData.createElement("CLASSIFICATION")
		newElem1.setAttribute "CNAME", ucase(sClassName)
		newElem1.setAttribute "CCODE", iClass
		newElem1.setAttribute "PGROUP", sPGroup

		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT ITEMCODE FROM INV_M_ITEMGROUP WHERE CLASSIFICATIONCODE = " & iClass & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing
		if not dcrs1.EOF then
			newElem1.setAttribute "DELETE", "N"
		else
			newElem1.setAttribute "DELETE", "Y"
		end if
		dcrs1.Close
		
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT ITEMCODE,ITEMDESCRIPTION,ORGUNITSHORTDESCRIPTION FROM VWALLITEMS WHERE CLASSIFICATIONCODE = " & iClass & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs1.ActiveConnection = nothing
		if not dcrs1.EOF then
			do while not dcrs1.EOF
				iItem = trim(dcrs1(0))
				sItemName = trim(dcrs1(1))

				set newElem2 = OutData.createElement("ITEMS")
				newElem2.setAttribute "INAME", ucase(sItemName)
				newElem2.setAttribute "ICODE", iItem
				newElem2.setAttribute "CCODE", iClass
				newElem2.setAttribute "ORG", trim(dcrs1(2))
			
				newElem1.appendChild newElem2

			dcrs1.MoveNext
			loop
		end if
		dcrs1.Close

		RootNode.appendChild newElem1
			
		' Function Call
		child iClass
	
	dcrs.MoveNext
	Loop
	dcrs.Close
			
	OutData.appendChild RootNode

	OutData.Save server.MapPath("../temp/master/ClassDelete.xml")
%>
<script type="application/xml" id="ItemData" data-itms-xml-island="1" data-src="<%="../temp/master/ClassDelete.xml"%>"></script>
<form method="POST" name="formname" action="" target="bodyFrame">
	<table border="0" cellspacing="0" width="100%" cellpadding="0">
		<tr>
			<td class="ExcelHeaderCell" colspan="3"><p align="center">Classification Deletion - Details</td>
		</tr>
		<tr>
			<td width="10" colspan="3" class="MiddlePack"></td>
		</tr>
		<tr>
			<td width="5"></td>
			<td>
				<table cellpadding="0" cellspacing="0" width="100%" border="0">
					<tr>
						<td>
							<table cellpadding="0" cellspacing="0">
								<tr>
									<td valign="top" width="100%">
									<DIV class=frmBody id=frm2 style="width: 100%; height:352;">
									<table border="0" cellspacing="1" class="ExcelTable" width="100%">
										<tr>
											<td class="ExcelSerial" align="center" rowspan=2 width=5>S.No.</td>
											<td class="ExcelHeaderCell" colspan=2 align="left">Classification</td>
										</tr>
										<tr>
											<td class="ExcelHeaderCell" align="left">Item</td>
											<td class="ExcelHeaderCell" align="left">Unit</td>
										</tr>
								<%		
									iCtr = 0
									sExp ="//DETAILS/CLASSIFICATION"
									Set ClassNode = RootNode.Selectnodes(sExp)
									for iCtr =  0 to ClassNode.length - 1
										sClassName = ClassNode.Item(iCtr).Attributes.getNamedItem("CNAME").Value
										iClass = ClassNode.Item(iCtr).Attributes.getNamedItem("CCODE").Value
										sPGroup = ClassNode.Item(iCtr).Attributes.getNamedItem("PGROUP").Value

										sExp1 ="//DETAILS/CLASSIFICATION/ITEMS [ @CCODE = "&iClass&"]"
										Set ItemNode = RootNode.Selectnodes(sExp1)
										if ItemNode.Length > 0 then
								%>
										<tr>
											<td class="ExcelSerial" align="center"></td>
											<td class="ExcelDisplayCell" colspan=2 align="left"><B><%=sClassName%></B></td>
										</tr>
								<%
											for iCtr1 =  0 to ItemNode.length - 1
												iItem = ItemNode.Item(iCtr1).Attributes.getNamedItem("ICODE").Value
												sItemName = ItemNode.Item(iCtr1).Attributes.getNamedItem("INAME").Value
												sOrgName = ItemNode.Item(iCtr1).Attributes.getNamedItem("ORG").Value
												iCounter = iCounter + 1
								%>	
										<tr>
											<td class="ExcelSerial" align="center"><%=iCounter%></td>
											<td class="ExcelDisplayCell" align="left"><%=sItemName%></td>
											<td class="ExcelDisplayCell" align="left"><%=sOrgName%></td>
										</tr>
								<%
											next
										end if
									next
									iCtr1 = 0
									sExp1 ="//DETAILS/CLASSIFICATION [ @DELETE = 'Y']"
									Set DeleteNode = RootNode.Selectnodes(sExp1)
									if DeleteNode.Length > 0 then
										for iCtr1 =  0 to DeleteNode.length - 1
											sClassName = DeleteNode.Item(iCtr1).Attributes.getNamedItem("CNAME").Value
											sDeleteName = sDeleteName & sClassName & ","
										next	
									else
										sDeleteName = "NA"
									end if
									arrDelete = split(sDeleteName,",")
									iCtr = 0
									if iCounter = 0 then
										sDeleteName = mid(sDeleteName,1,len(sDeleteName)-1)
								%>
										<tr>
											<td  class="ExcelSerial" ></td>
											<td class="ExcelDisplayCell" colspan=2 align="center">No Items exists, Classification(s) <b>[<%=sDeleteName%>]</b> will be DELETED.</td>
										</tr>
										<input type=hidden name="hDelete" value="Y">
								<%	else 
										if UBound(arrDelete) > 0 then
											sDeleteName = mid(sDeleteName,1,len(sDeleteName)-1)
								%>
										<tr>
											<td  class="ExcelSerial" ></td>
											<td class="ExcelDisplayCell" colspan=2 align="center">The above listed Classification(s) has Item(s) and will not be DELETED. Remaining Classification(s) <b>[<%=sDeleteName%>]</b> will be DELETED.</td>
										</tr>
										<input type=hidden name="hDelete" value="Y">
								<%		else %>
										<tr>
											<td  class="ExcelSerial" ></td>
											<td class="ExcelDisplayCell" colspan=2 align="center"><b>The above listed Classification(s) has Item(s) and will not be DELETED.</b></td>
										</tr>
										<input type=hidden name="hDelete" value="N">
								<%
										end if
									end if 
								%>
									</table>
									</div>
									</td>
								</tr>
								<tr>
									<td align="center" class="BottomPack" width="100%" colspan="3">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>
								<tr>
									<td class=ActionCell colspan="2"> <p align="center">
								    <input type=hidden name="hCounter" value="<%=iCounter%>">
								    <input type="button" value="Delete" name="B2" class="ActionButton" onClick="DeleteClass()">
								</tr>
							</table>
						</td>
						<td width="5"></td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
</form>
</BODY>
</HTML>

<%
Private Sub child(sid)
    Dim dcrs1,dcrs2,newElem11,newElem12
    dim stGCode,stGName,sPGroup,iItem,sItemName

	set dcrs1 = Server.CreateObject("ADODB.Recordset")
	set dcrs2 = Server.CreateObject("ADODB.Recordset")
    
	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT GROUPCODE,GROUPNAME,PARENTGROUP FROM INV_M_CLASSIFICATION WHERE PARENTGROUP = " & sid & " AND GROUPCODE <> " & sid & " ORDER BY GROUPCODE"
		.ActiveConnection = con
		.Open
	end with

    Do While Not dcrs1.EOF

		stGCode = trim(dcrs1(0))
		stGName = trim(dcrs1(1))
		sPGroup = trim(dcrs1(2))
		
		set newElem11 = OutData.createElement("CLASSIFICATION")
		newElem11.setAttribute "CNAME", ucase(stGName)
        newElem11.setAttribute "CCODE", stGCode
        newElem11.setAttribute "PGROUP", sPGroup

		with dcrs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT ITEMCODE FROM INV_M_ITEMGROUP WHERE CLASSIFICATIONCODE = " & stGCode & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs2.ActiveConnection = nothing
		if not dcrs2.EOF then
			newElem11.setAttribute "DELETE", "N"
		else
			newElem11.setAttribute "DELETE", "Y"
		end if
		dcrs2.Close

		with dcrs2
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT ITEMCODE,ITEMDESCRIPTION,ORGUNITSHORTDESCRIPTION FROM VWALLITEMS WHERE CLASSIFICATIONCODE = " & stGCode & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs2.ActiveConnection = nothing
		if not dcrs2.EOF then
			do while not dcrs2.EOF
				iItem = trim(dcrs2(0))
				sItemName = trim(dcrs2(1))

				set newElem12 = OutData.createElement("ITEMS")
				newElem12.setAttribute "INAME", ucase(sItemName)
				newElem12.setAttribute "ICODE", iItem
				newElem12.setAttribute "CCODE", stGCode
				newElem12.setAttribute "ORG", trim(dcrs2(2))
			
				newElem11.appendChild newElem12

			dcrs2.MoveNext
			loop
		end if
		dcrs2.Close
        
        RootNode.appendChild(newElem11)

        child stGCode
        
    dcrs1.MoveNext
    Loop
    dcrs1.Close
End Sub

%>
