<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	ItmOpUoMSalPoPEntry.asp
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
	'Connects To				:	ItmOpUoMSalPoPInsert.asp
	'Procedures/Functions Used	:	populateUoMList
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
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Optional UoM</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<%  	'XML DOM Variables
	Dim oDOM,Root,objfs,PGNode,rsTemp
    dim sOrgName,sClassName,sClassCode,sItmDescr,sOrgCode,sUoM,sQuery,iItmCode,sCallFrom
	' Create our DOM Document Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	set rsTemp = Server.CreateObject("ADODB.Recordset")

	sOrgCode = Session("organizationcode")
	sOrgName = Session("OrgShortName")

	iItmCode = Request.QueryString("iItmCode")
    sCallFrom = Request("CallFrom")
%>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
function  selectTheItem(obj,srcCombo){
var i;
		objSel = document.forms[0].elements[srcCombo];
		i = 0;
		if (obj.value == "") {
			for(i=0; i < objSel.options.length; i++){
				objSel.options[i].selected = false;
			}
		}
		i = 0;
		for(i=0; i < objSel.options.length; i++){
			if (obj.value != "" && objSel.options[i].text.toUpperCase().indexOf(obj.value.toUpperCase()) >=0 ){
				objSel.options[i].selected = true;
				return;
			}
		}
		if (obj.value == "") {
			objSel.selectedIndex = -1;
		}
}
</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itemOptionalUom.js"></SCRIPT>
</HEAD>
<%
    sQuery = "Select StoresUOM,ClassificationCode from VWITEM V where ItemCode = "& iItmCode

    rsTemp.Open sQuery,con
    if not rsTemp.EOF then
        sUoM = rsTemp(0)
        sClassCode = rsTemp(1)
    end if
    rsTemp.Close
%>

<BODY leftMargin=0 topMargin=0 >
<form method="POST" name="formname" action="">
<input type="hidden" name="hCallFrom" value="<%=sCallFrom%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Optional UoM
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
											<td class="FieldCell">Sales UoM</td>
											<td class="FieldCellSub"><span class="DataOnly"><%=sUoM%>&nbsp;</span></td>
										</tr>
										<tr>
											<td class="FieldCell">Enter few characters</td>
											<td class="FieldCellSub">
												<input type="text" name="txtSearch" size="11" class="Formelem"  ONKEYUP="javascript:selectTheItem(this,'selItem')">
											</td>
										</tr>
										<tr>
										    <td class="FieldCell" valign="top"> Optional UoM</td>
										    <td class="FieldCellSub">
												<select size="5" name="selItem" class="FormElem">
											<%	'Calling the Function which populates UoM List
												populateUoMList sUoM
											%>
												</select>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Factor</td>
											<td class="FieldCellSub">
												<input type="text" name="txtFactor" size="8" maxlength=9 class="FormElem">
											</td>
										</tr>
										<tr>
										    <td class="FieldCell">Operator</td>
										    <td class="FieldCellSub">
												<select size="1" name="selOpe" class="FormElem">
													<option value="select">Select</option>
													<option value="0">*</option>
													<option value="1">/</option>
												</select>&nbsp;
												<input type="button" value=" Add " name="B3" class="AddButtonX" onClick="CheckEntry()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
                            <tr>
								<td align="center"></td>
								<td width="100%">
									<div class="frmBody" id="frm2" style="width: 100%; height:78;">
										<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Alternate UoM</td>
												<td class="ExcelHeaderCell" align="center" width="100">Factor</td>
												<td class="ExcelHeaderCell" align="center" width="60">Operator</td>
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
	' Function to populate UoM List
	Function populateUoMList(sUoM)
		' Declaration of variables
		Dim dcrs,sUomDesc,sUomShDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMDESCRIPTION,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE NOT IN (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItmCode & " AND CLASSIFICATIONCODE = " & sClassCode & " AND ORGANISATIONCODE = " & Pack(sOrgCode) & ") AND UOMCODE <> " & Pack(sUoM) & " ORDER BY UOMCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUoMCode = dcrs(0)
		set sUomDesc = dcrs(1)
		set sUomShDesc = dcrs(2)

		Do While Not dcrs.EOF
			Response.Write("<OPTION VALUE="""&trim(sUoMCode)&""">"&trim(sUomShDesc)&"</OPTION>" &vbcrlf)
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
