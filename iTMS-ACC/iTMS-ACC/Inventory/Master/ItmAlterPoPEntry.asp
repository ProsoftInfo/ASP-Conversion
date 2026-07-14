<%@ Language=VBScript%>
<%option explicit%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ItmAlterPoPEntry.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	Ragavendran
	'Created On					:	Jul 12,2011
	'Modified By				:
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/ItemDisplay.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Alternate Item</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%

Response.Write "<Font color=red>"
'XML DOM Variables
	Dim oDOM,Root,objfs,PGNode,rsTemp
	dim iItmCode,sQuery
	dim sOrgName,sClassName,sClassCode,sItmDescr,sOrgCode,iCount,sItemType

	iCount = 0
	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	Set rsTemp = Server.CreateObject("ADODB.Recordset")

	iItmCode = Request.QueryString("iItmCode")

	sOrgCode = Session("organizationcode")
	sOrgName = Session("OrgShortName")

	'sQuery = "Select ItemTypeID,ClassificationCode from VWITEM V where ItemCode = "& iItmCode
	sQuery = "Select ClassificationCode from VWITEM V where ItemCode = "& iItmCode
    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
       ' sItemType = Trim(rsTemp(0))
        'sClassCode = trim(rsTemp(1))
        sClassCode = Trim(rsTemp(0))
    end if
    rsTemp.Close

%>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itmCtrlBoM.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemAlternate.js"></SCRIPT>

</HEAD>
<BODY leftMargin=0 topMargin=0>
<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Alternate Items
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
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td width="100%">
                                    <table border="0" cellspacing="0" cellpadding="0" class="TableOutlineOnly">
									    <tr>
											<td class="FieldCellSub" width="80">Item Name</td>
											<td>
											<span class="DataOnly" id="txtItemName">&nbsp;</span>
											</td>
											<td class="FieldCell" width="15"></td>
											<td class="FieldCell" width="82">Classification</td>
											<td>
											<span class="DataOnly" id="txtClassName">&nbsp;</span>
											&nbsp;</td>
											<td></td>
									    </tr>
                                    </table>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td width="100%">
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td class="FieldCell">Enter few characters</td>
											<td class="FieldCellSub">
												<input type="text" name="txtSearch" size="11" class="Formelem"  ONKEYUP="javascript:selectTheItem(this,'selItem')">
											</td>
										</tr>
										<tr>
										    <td class="FieldCell" valign="top">Classification / Item</td>
										    <td class="FieldCellSub">
												<select size="5" name="selItem" class="FormElem">
											<%	'Calling the Function which populates Classification and Item List
												populateClassItem
											%>
												</select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Priority</td>
											<td class="FieldCellSub">
												<input type="text" name="txtPriority" size="4" maxlength=3 class="FormElem">&nbsp;
												<input type="button" value=" Add " name="B3" class="AddButtonX" onClick="CheckEntry('<%=iCount%>')">
											</td>
										</tr>
									</table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td width="100%">
									<div class="frmBody" id="frm2" style="width:100%;height:90;">
										<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Item Name</td>
												<td class="ExcelHeaderCell" align="center" width="60">Priority</td>
											</tr>
										</table>
									</div>
								</td>
								<td align="center"></td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="button" value="Cancel" name="B2" class="ActionButton" onClick="window.close()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
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
<%
	' Function to populate Classification and Item
	Function populateClassItem()
		' Declaration of variables
		Dim dcrs,sItemDesc,sItemShDesc,sClassDesc,iTempItmCode,sTempClassCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT ITEMDESCRIPTION,SHORTDESCRIPTION,GROUPNAME,IM.CLASSIFICATIONCODE,IM.ITEMCODE FROM INV_M_ITEMMASTER IM,INV_M_CLASSIFICATION IC WHERE IM.CLASSIFICATIONCODE = IC.GROUPCODE AND IM.ITEMCODE <> " & iItmCode & " AND IM.CLASSIFICATIONCODE <> " & sClassCode & " AND IC.ITEMTYPEID = " & Pack(sItemType) & " AND IM.ORGANISATIONCODE = " & Pack(sOrgCode) & " ORDER BY IM.ITEMCODE"
			.Source = "SELECT ITEMDESCRIPTION,SHORTDESCRIPTION,GROUPNAME,IM.CLASSIFICATIONCODE,IM.ITEMCODE FROM INV_M_ITEMMASTER IM,INV_M_CLASSIFICATION IC WHERE IM.CLASSIFICATIONCODE = IC.GROUPCODE AND IM.ITEMCODE <> " & iItmCode & " AND IM.CLASSIFICATIONCODE <> " & sClassCode & " AND IM.ORGANISATIONCODE = " & Pack(sOrgCode) & " ORDER BY IM.ITEMCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sItemShDesc = dcrs(1)
		set sClassDesc = dcrs(2)
		set sTempClassCode = dcrs(3)
		set iTempItmCode = dcrs(4)

		Do While Not dcrs.EOF
			iCount = iCount + 1
			sItemDesc = ItemDisplay(iTempItmCode,sTempClassCode)
			Response.Write("<OPTION VALUE="""&trim(sTempClassCode)&":"&trim(iTempItmCode)&""">"&trim(sItemDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
