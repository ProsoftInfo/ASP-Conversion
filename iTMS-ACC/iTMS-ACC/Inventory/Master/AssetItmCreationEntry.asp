<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	AssetItmCreationEntry.asp
	'Module Name				:	Inventory (Asset Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 16, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	AssetItmCreationInsert.asp
	'Procedures/Functions Used	:	populateUoM
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
<!--#include virtual="/include/GetSettings.asp"-->
<%
	dim sIP
	sIP = GetSettings("IP")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Asset Item Creation</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="Data"><root/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutDataO"><root/></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData"><Output/></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/ModalReturnCompat.js"></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/Cancel.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/assetItemCreation.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript>
	function trimTrue(val){
		var ltrim = /^\s+/g;
		var rtrim = /\s+$/g;
		return val.replace(ltrim,'').replace(rtrim,'');
	}

	/*function Init(val){
		if(!val == "") {
			var frm = window.frames;
			frm(0).ctlCategoryTree.IType = val+":"+document.forms[0].selItmType(document.forms[0].selItmType.selectedIndex).text+":NO"
		}
	}
	function classSelected(){
		var frm = window.frames;
		sTemp = frm(0).ctlCategoryTree.classification;
		return sTemp;
	}*/

</SCRIPT>


</HEAD>
<%
	dim dcrs,iAssetCode,sDesc,sAddDesc,sShDesc,sIType,sCode,sOrgID,sUOM,sTempArr,sCallFrom
	Dim iItmCode,sTempMonYr,sMonYr
	'Declaration of Objects

	'iAssetCode = trim(Request.Form("selItem"))  'Blocked by UmaMaheswari S
	'iAssetCode = Request.QueryString("AssetCode")

	sTempArr   = Request.QueryString("sTemp")
	iAssetCode = split(sTempArr,":")(0)
	sCallFrom  = split(sTempArr,":")(1)
	sDesc	   = trim(Request.Form("hTempItemname"))

	'Response.Write "<p>sTempArr = "& sTempArr
	sTempMonYr = mid(FormatDate(date),4,2)
	sMonYr = sTempMonYr&Year(date())

	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT ASSETDESCID,ASSETDESCRIPTION,OUDEFINITIONID FROM VWASSETSITEM WHERE ASSETDESCID = " & iAssetCode & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sCode = trim(dcrs(0))
		sDesc = trim(dcrs(1))
		sOrgID = trim(dcrs(2))
	end if
	dcrs.Close

	'taking UOM
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT Uom From Far_T_AssetDetails WHERE ASSETDESCID = " & iAssetCode & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		sUOM = trim(dcrs(0))
	end if
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT ISNULL(MAX(ITEMCODE),0) + 1 FROM INV_M_ITEMMASTER WHERE (ITEMCODE = (SELECT ISNULL(MAX(ITEMCODE), 0) FROM INV_M_ITEMMASTER))"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iItmCode = trim(dcrs(0))
	end if
	dcrs.close

	Dim iQty,iValue

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT COUNT(TAGNUMBER) FROM FAR_T_ASSETDETAILS WHERE ASSETDESCID = " & iAssetCode & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iQty = dcrs(0)
	end if
	dcrs.Close

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT SUM(SUPPLIERINVOICEVALUE) FROM FAR_T_ASSETSUPPLIER WHERE ASSETNUMBER IN (SELECT ASSETNUMBER FROM FAR_T_ASSETDETAILS WHERE ASSETDESCID = " & iAssetCode & ")"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	if not dcrs.EOF then
		iValue = dcrs(0)
	end if
	dcrs.Close
%>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="Init('PLA')">
<form method="POST" name="formname" action="" TARGET="bodyFrame">
<input type=hidden name="hIType" value="PLA">
<input type=hidden name="hOrgID" value="<%=sOrgID%>">
<input type=hidden name="hAssetCode" value="<%=iAssetCode%>">
<INPUT TYPE=HIDDEN NAME="hClassSelected" VALUE="">
<input type=hidden name="hCallFrom" value="<%=sCallFrom%>">
<input type=hidden name="hItemCode" value="<%=iItmCode%>">
<input type=hidden name="hQty" value="<%=iQty%>">
<input type=hidden name="hMonthYr" value="<%=sMonYr%>">
<input type=hidden name="hValue" value="<%=iValue%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Asset Item Creation</p>
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
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td align="center" class="MiddlePack" width="100%">
											</td>
										</tr>
										<tr>
											<td class='FieldCell'>Item Type &nbsp;
												<select size="1" name="selItmType" class="FormElem" disabled>
													<option value="select">Select</option>
													<%	'Calling the Function which populates the Item Type list
														'populateItemTypeSelected "PLA"
														popItemTypes 6
													%>
												</select>
                                            </td>
										</tr>
                                        <tr>
											<td align="center" class="MiddlePack" width="100%">
											</td>
                                        </tr>
                                        <tr>
											<td>
                                                <table border="0" class="TableOutlineOnly" width="589">
                                                  <tr>
                                                    <td>
														<!--<IFRAME NAME="ifr "ID="iframe" FRAMEBORDER=0 SCROLLING=AUTO SRC="comItemClassificationTree.asp" NORESIZE="RESIZE" STYLE="WIDTH=100%; HEIGHT=350"></IFRAME>-->
														 <td>
                                                        <div id="ctlCategoryTree" data-itms-tree-control data-tree-kind="item-classification"
	                                                    data-dsn="../Components/GetCategoryGroup.asp" data-itype="NAP:NO:NO"
	                                                    data-width="552px" data-height="340px"></div>
                                                        </td>
                                                  </tr>
                                                </table>
											</td>
                                        </tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%">
											</td>
										</tr>
										<tr>
											<td width="100%">
												<table cellpadding="0" cellspacing="0">
													<tr>
														<td class=FieldCell> Item Code</td>
														<td class='FieldCellSub'>
															<input type="text" name="txtItmCode" size="19" maxlength=15 class="Formelem" value="<%=sCode%>">
                                                            <input type="button" value="Existing Items" class="AddButtonX" onClick="DisplayItemCode()" id=button1 name=button1>
                                                        </td>
													</tr>
													<tr>
														<td class=FieldCell> Name</td>
														<td class='FieldCellSub'><input type="text" name="txtItmDesc" size="50" value="<%=sDesc%>" maxlength=60 class="Formelem"></td>
													</tr>
													<tr>
														<td class=FieldCell> UoM</td>
														<td class='FieldCellSub'>
															<select size="1" name="selUoMStores" class="FormElem">
																<option value="select">Select</option>
																<%	'Calling the Function which populates the UoM list
																	populateUoM
																%>
															</select>
														</td>
													</tr>
													<tr>
														<td class=FieldCell valign="top"> Storage Location</td>
														<td class='FieldCellSub' colspan="4">
															<select size="5" name="selStorage" class="FormElem" multiple>
															<%	'Calling the Function which populates the Store list
																populateStores sOrgID
															%>
															</select>
													    </td>
													</tr>
													<tr>
														<td class=FieldCell> Applicable For</td>
														<td class='FieldCell'>
                                                            <input type="checkbox" name="chkAppP" value="P" class="FormElem"> Production &nbsp;
                                                            <input type="checkbox" name="chkAppM" value="M" class="FormElem"> Maintenance
                                                        </td>
													</tr>
													<tr>
														<td class=FieldCell> BoM</td>
														<td class='FieldCellSub'>
                                                            <input type="button" value="Yes" class="AddButton" onclick="GetDetails()">
                                                        </td>
													</tr>
												</table>
											</td>
										</tr>
                                        </tr>
										<tr>
											<td align="center" class="MiddlePack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
										</tr>
										<tr>
											<td width="100%">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
                                                                <input type="button" value="Create" name="B1" class="ActionButton" onClick="CheckSubmit()">
																<input type="reset" value="Reset" name="B2" class="ActionButton">
																<input type="button" value="Cancel" name="B3" class="ActionButton" onClick="Cancel('../welcome_Inventory.asp')">
														</td>
													</tr>
												</table>
											</td>
										</tr>
                                        <tr>
											<td align="center" class="BottomPack" width="100%">
												<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											</td>
                                        </tr>
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
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
	' Function to populate Store
	Function populateStores(sOrgID)
		' Declaration of variables
		Dim dcrs,dcrs1,sLoc,sBin,sBinName,sLocName,sLocCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT DISTINCT LOCATIONNUMBER,LOCATIONNAME,APPLICABLEFOR FROM INV_M_ORGSTORAGE WHERE OUDEFINITIONID = " & Pack(sOrgID) & " ORDER BY 1"
			.Source = "SELECT DISTINCT LOCATIONNUMBER,LOCATIONNAME,APPLICABLEFOR FROM INV_M_STORAGE WHERE OUDEFINITIONID = " & Pack(sOrgID) & " ORDER BY 1"
			'Response.Write dcrs.source
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		if not dcrs.EOF then
			Do While Not dcrs.EOF
				sLoc = trim(dcrs(0))
				sLocName = trim(dcrs(1))
				sLocCode = trim(dcrs(2))

				with dcrs1
					.CursorLocation = 3
					.CursorType = 3
					'.Source = "SELECT BINNUMBER,BINNAME,BINCODE FROM INV_M_ORGSLBINDETAILS WHERE LOCATIONNUMBER = " & sLoc & " ORDER BY BINNUMBER"
					.Source = "SELECT DISTINCT BINNUMBER,BINNAME,BINCODE FROM INV_M_STOREBINDETAILS WHERE OUDEFINITIONID = '" & trim(sOrgID) & "' AND LOCATIONNUMBER = " & trim(sLoc) & " ORDER BY 1"
					.ActiveConnection = con
					.Open
				end with
				set dcrs1.ActiveConnection = nothing

				if not dcrs1.EOF then
					do while not dcrs1.EOF
						'Response.Write("<OPTION VALUE="""&sLoc&"-"&trim(dcrs1(0))&"-"&sLocCode&"-"&sLocName&""">"&sLocName&" -- "&trim(dcrs1(1))&"</OPTION>" &vbcrlf)
						Response.Write("<OPTION VALUE="""&sLoc&"~"&trim(dcrs1(0))&"~"&sLocCode&"~"&sLocName&""">"&sLocName&" -- "&trim(dcrs1(1))&"</OPTION>" &vbcrlf)
					dcrs1.MoveNext
					loop
				else
					Response.Write("<OPTION VALUE="""&sLoc&"~NULL~"&sLocCode&"~"&sLocName&""">"&sLocName&"</OPTION>" &vbcrlf)
				end if
				dcrs1.Close

			dcrs.MoveNext
			Loop
		end if
		dcrs.Close
	End Function

	' Function to populate the UoM list
	Function populateUoM()
		' Declaration of variables
		Dim oDom,fs,Root,PGNode
		dim sUoMID,sUoMName,sUoMShName

		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		Set fs = CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(Server.MapPath("../../Inventory/xmldata/UoM.xml")) then
			oDOM.Load server.MapPath("../../Inventory/xmldata/UoM.xml")
			Set Root = oDOM.documentElement
			if Root.HaschildNodes() then
				For Each PGNode In Root.childNodes
					sUoMID = trim(PGNode.Attributes.Item(0).nodeValue)
					sUoMName = trim(PGNode.Attributes.Item(1).nodeValue)
					sUoMShName = trim(PGNode.Attributes.Item(2).nodeValue)
					IF CStr(sUoMID) = CStr(sUOM) Then
						Response.Write("<OPTION VALUE="""&trim(sUoMID)&""" Selected>"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
					Else
						Response.Write("<OPTION VALUE="""&trim(sUoMID)&""">"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
					End IF
					'Response.Write("<OPTION VALUE="""&trim(sUoMID)&""">"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
				next
			end if
		end if
	End Function
%>


