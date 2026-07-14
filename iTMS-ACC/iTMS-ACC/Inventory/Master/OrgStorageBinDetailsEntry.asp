<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	OrgStorageBinDetailsEntry.asp
	'Module Name				:	Inventory (Storage Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 14, 2002
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	OrgStorageBinDetailsInsert.asp
	'Procedures/Functions Used	:	populateUnit
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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Storage Bin Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
 <SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/orgStorageBinDetails.js"></SCRIPT>
<!--SCRIPT LANGUAGE=javascript SRC="../scripts/stoLocBinDetails.js"></SCRIPT-->

<%
'	Dim oDOM,Root,ObjFs,rs,sExp,OrgNode,sOrgID,iNoOfBins,StNode,iLocNo
'	Dim iBinNo,sBinCode,sBinName,sBinArea,BinElem,iTyFree,iTyBin,sQry
'
'	Set oDOM = server.CreateObject("Microsoft.XMLDOM")
'	Set objfs = CreateObject("Scripting.FileSystemObject")
'	Set rs = Server.CreateObject("ADODB.RecordSet")
'
'	if objfs.FileExists(server.MapPath("../Temp/Master/StorageNew"&Session.SessionID&".xml")) then
'		oDOM.load  server.MapPath("../Temp/Master/StorageNew"&Session.SessionID&".xml")
'	end if
'	Set Root = oDOM.documentElement
'
'	sExp = "//Organization"
'	Set OrgNode = Root.SelectNodes(sExp)
'	IF OrgNode.length <> 0 then
'		sOrgID = OrgNode.item(0).Attributes.getNamedItem("OUDEFINITIONID").value
'	End IF
'
'		'Response.Write iNoOfBins
'
'	IF Root.haschildnodes then
'		For each OrgNode in Root.childnodes
'			For each StNode in OrgNode.childnodes
'				IF trim(StNode.NodeName) = "Storage" then
'					iLocNo = StNode.getAttribute("LOCATIONNUMBER")
'					IF trim(StNode.getAttribute("STORAGETYPEBINS")) <> "0" and trim(StNode.getAttribute("STORAGETYPEFREE")) = "0" then
'						sQry = "Select BinNumber,BinCode,BinName,BinArea from Inv_M_OrgSLBinDetails where  "&_
'								"OUDefinitionID = '"& sOrgID &"' and LocationNumber = "& iLocNo &" "
'						'Response.Write sQry
'						rs.Open sQry,con
'
'						do while not rs.EOF
'							iBinNo	  = rs(0)
'							sBinCode  = rs(1)
'							sBinName  = rs(2)
'							sBinArea   = rs(3)
'
'							set BinElem	 = oDOM.createElement("Bin")
'							BinElem.setAttribute "BINNUMBER",iBinNo
'							BinElem.setAttribute "BINCODE",sBinCode
'							BinElem.setAttribute "BINNAME",sBinName
'							BinElem.setAttribute "BINAREA",sBinArea
'							StNode.appendchild BinElem
'
'							rs.MoveNext
'						loop
'						rs.Close
'					End If
'
'				End IF 'IF trim(StNode.NodeName) = "Storage" then
'			Next 'For each StNode in OrgNode.childnodes
'		Next ' For each OrgNode in Root.childnodes
'	End IF 'IF Root.haschildnodes then
'	oDOM.save server.MapPath("../Temp/Master/StorageNew"&Session.SessionID&".xml")
%>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="FnInit()">

<form method="POST" name="formname" action="">
<input type="hidden" name="hBinNo" value="">
<input type="hidden" name="hLocNo" value="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Storage Location Creation
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%"  >
				<TR>
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" width="100%">

									</table>
								</td>
								<td >
									<table border="0" cellpadding="0" cellspacing="0" width="100%" >

									</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    <p align="center"><font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font>
								</td>
							</tr>
						</table>
					</td>
				</tr >
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td class=FieldCell> Organization</td>
											<td class='FieldCellSub'><span id="UnitID" class="Dataonly"></span>
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell> Location Name</td>
											<td class='FieldCellSub'>
											<span id="LocID" class="Dataonly"></span>

                                            </td>
										</tr>
									</table>
                                    </div>
								</td>
								<td align="center" width="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" class="MiddlePack">
                                    <table border="0" cellspacing="1" Id ="tblBin" name="tblBin" class="ExcelTable" width="100%"></table>
								</td>
								<td align="center" width="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
													<input type="button" value="Add New" name="B4" class="ActionButtonX" onClick="AddNew()">
                                                    <input type="button" value="Done" name="B2" class="ActionButton" onClick="CheckSubmit()">
													<input type="reset" value="Reset" name="B1" class="ActionButton" onClick="ClearAll()" >
													<input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Window_close()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>

