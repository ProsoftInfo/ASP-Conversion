<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItmTypeAttributeEntry.asp
	'Module Name				:	Inventory (Item Control Definition)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	January 29, 2003
	'Modified By				:	Ragavendran R
	'Modified On				:	March 06,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	ItmTypeAttributeInsert.asp
	'Procedures/Functions Used	:	populateItemType
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
<%
	Dim sItemType,sHeader,sAttName,sDataType,sLength,sDecimal,sClassCode,sClassName
	Dim sMode ,sValue,sAttributeID,sSql
	Dim oDOM,dcrs,stypID,stypName,sItmType
	Dim ndRoot,ndTypeHeader
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")


    sItemType = Request.QueryString("ItemType")
    sClassCode = Request.QueryString("ClassCode")
    
    sSql = "Select GroupName from INV_M_Classification where GroupCode = "& sClassCode 
    dcrs.Open sSql,con
    if not dcrs.EOF then
        sClassName = trim(dcrs(0))
    end if
    dcrs.Close 

	sMode = Request.QueryString("Mod")
	if sMode ="" or IsNull(sMode) then sMode= "S"

		sValue = Request.QueryString("sValue")
		'Response.Write "sValue = "& sValue
		if sValue="" or IsNull(sValue) then sValue= 0
    if sValue<>"0" and trim(sValue)<>"" then
        'sSql ="Select ItemTypeAttributeID,ItemTypeID,HeaderID,ItemTypeAttributeName,ItemTypeAttributeType,ItemTypeAttributeDataLength,ItemTypeAttributeDecimal from INV_M_ITEMTYPEATTRIBUTES where ItemTypeAttributeID = "& sValue &" and ClassificationCode = " &sClassCode
        sSql ="Select ItemTypeAttributeID,HeaderID,ItemTypeAttributeName,ItemTypeAttributeType,ItemTypeAttributeDataLength,ItemTypeAttributeDecimal from INV_M_ITEMTYPEATTRIBUTES where ItemTypeAttributeID = "& sValue &" and ClassificationCode = " &sClassCode
	        'Response.Write sSql & vbCrLf
            dcrs.Open sSql,con
	        if not dcrs.EOF then
    	        sAttributeID = dcrs(0)
		        'sItemType = dcrs(1)
		        sHeader = trim(dcrs(1))
		        sAttName = dcrs(2)
		        sDataType = trim(lcase(dcrs(3)))
		        sLength = dcrs(4)
		        sDecimal = dcrs(5)
	        end if 'if not dcrs.EOF then
	        dcrs.Close
    end if
    'Response.Write "sDataType = "& sDataType
    'Response.Write "sHeader = "& sHeader
if sDataType= "string" then sDataType = "String"
if sDataType= "numeric" then sDataType="Numeric"
if sDataType= "options" then sDataType= "Options"
if isnull(sItemType) or sItemType="" then sItemType="select"
if IsNull(sHeader) or sHeader="" then sHeader = "select"
if IsNull(sAttName) or sAttName="" then sAttName=""
if sDataType="" or IsNull(sDataType) then sDataType="select"
'Response.Write "sDataType = "& sDataType
		set ndRoot = oDOM.createElement("Root")
		oDOM.appendChild ndRoot

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			'.Source = "SELECT DISTINCT HEADERID,ITEMTYPEHEADERNAME,ITEMTYPEID FROM INV_M_ITEMTYPEHEADER ORDER BY HEADERID"
			.Source = "SELECT DISTINCT HEADERID,ITEMTYPEHEADERNAME FROM INV_M_ITEMTYPEHEADER ORDER BY HEADERID"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set stypID = dcrs(0)
		set stypName = dcrs(1)
		'set sItmType = dcrs(2)

		If not dcrs.EOF then
			Do While Not dcrs.EOF

				set ndTypeHeader =oDOM.createElement("Type")
				ndTypeHeader.setAttribute "ID", stypID
				ndTypeHeader.setAttribute "Name", stypName
				ndTypeHeader.setAttribute "IType", sItmType
				ndRoot.appendChild ndTypeHeader

				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		oDOM.save (Server.MapPath("../temp/transaction/ItemTypeHeader.xml"))

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Item Type Attributes</TITLE>
<base target="_self"></base>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="OutData"><root ITYPE="" HEADER="" ATTRNAME="" DATATYPE="" DATALENGTH="" DECIMAL=""></root></script>
<script type="application/xml" data-itms-xml-island="1" id="EditData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="NewData"><Root/></script>
<script type="application/xml" data-itms-xml-island="1" id="AttData"><Root></Root></script>
<script type="application/xml" data-itms-xml-island="1" id="ItemTypeHeader" data-src="<%="../temp/transaction/ItemTypeHeader.xml"%>"><Root/></script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/itmTypeAttribute.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="Init();populateUpdate();Init2()">
<form method="POST" name="formname">
<input type=hidden name="hRow" value="0">
<input type=hidden name="hItemType" value="<%=sItemType%>">
<input type=hidden name="hHeader" value="<%=sHeader%>">
<input type=hidden name="hAttribute" value="<%=sAttName%>">
<input type=hidden name="hType" value="<%=sDataType%>">
<input type=hidden name="hLength" value="<%=sLength%>">
<input type=hidden name="hDecimal" value="<%=sDecimal%>">
<input type=hidden name="hValue" value="<%=sValue%>">
<input type=hidden name="hOldAttID" value="<%=sAttributeID%>">
<input type=hidden name="hRowCtr" value="0">
<input type=hidden name="hClassCode" value="<%=sClassCode%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class="pagetitle" height="20"><p align="center">Attributes
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
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
									    <tr>
									        <td>
									            <table border="0" cellpadding="0" cellspacing="0">
									                <tr>
									                    <td class="fieldcellsub" colspan="2">
									                        <input type="button" name="btnManage" class="actionbutton" value="Manage" onclick="ManageAttribute('M')" />
									                        <input type="button" name="btnSelect" class="actionbutton" value="Select" onclick="ManageAttribute('S')"/>
									                    </td>
									                </tr>
									                <tr>
											            <td align="center" class="MiddlePack" width="100%" colspan="2">
												            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											            </td>
										            </tr>
									                <tr>    
									                    <td class=FieldCellSub> Classification Name</td>
											            <td class='FieldCellSub'>
												            <span id="txtClass" class="DataOnly"><%=sClassName%></span>
											            </td>
									                </tr>
									                 <tr>
											            <td align="center" class="MiddlePack" width="100%" colspan="2">
												            <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
											            </td>
										            </tr>   
									            </table>
									        </td>
										</tr>
										<tr>
											<td width="100%">
											    <div id="divManageAtt" style="display:none">
											        <table cellpadding="0" cellspacing="0" border="0" width="100%">
											            <tr>
											                <td>
											                    <table cellpadding="0" cellspacing="0" border="0">
													                <tr>
														                <td class="FieldCellSub"> Header Name</td>
														                <td class="FieldCellSub">
															                <select size="1" name="selHeader" class="FormElem">
																                <option value="select">Select</option>
															                </select>
														                </td>
													                </tr>
													                <tr>
														                <td class=FieldCellSub> Attribute</td>
														                <td class='FieldCellSub'>
															                <input type="text" name="txtAttribute" size="50" maxlength=50 class="Formelem">
														                </td>
													                </tr>
													                <tr>
														                <td class=FieldCellSub> Data Type</td>
														                <td class='FieldCellSub'>
															                <%if sMode="S" then%>
																                <select size="1" name="selDataType" class="FormElem" onChange="checkSelect()">
															                <%else%>
																                <select size="1" name="selDataType" class="FormElem" onChange="popOptionSel()">
															                <%end if%>
																                <option value="select">Select</option>
																                <option value="String">String</option>
																                <option value="Numeric">Numeric</option>
																                <option value="Options">Options</option>
															                </select>
															                &nbsp;&nbsp;&nbsp;
															                <%if sMode<>"S" then%>
															                <img border="0" onClick="popOptionSel()" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor: hand" width="11" height="11">
															                <%end if%>
														                </td>
													                </tr>
													                <tr>
														                <td class=FieldCellSub> Data Length</td>
														                <td class='FieldCellSub'>
															                <input type="text" name="txtDataLen" size="5" maxlength=4 class="Formelem">
														                </td>
													                </tr>
													                <tr>
														                <td class=FieldCellSub> Decimals</td>
														                <td class='FieldCellSub'>
															                <input type="text" name="txtDecimals" size="5" maxlength=4 class="Formelem">
														                </td>
													                </tr>
												                </table>
											                </td>
											            </tr>
											            <tr>
											                <td>
											                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
													                <tr>
														                <td valign="middle" class="ActionCell">
															                <p align="center">
																                <%if sMode="S" then%>
															                        <input type="button" value="Save" name="B1" class="ActionButton" onClick="CheckSubmit()">
															                    <%else%>
																	                <input type="button" value="Update" name="B4" class="ActionButton" onClick="CheckUpdate()">
																                <%end if%>
																                <input type="reset" value="Reset" name="B2" class="ActionButton">
																                <input type="button" value="Cancel" name="B3" class="ActionButton" onClick="window.close()">
														                </td>
													                </tr>
												                </table>
												            </td>
											            </tr>
											            <tr>
											                <td>
											                    <table border="0" cellpadding="0" cellspacing="1" width="100%" id="tblDisplay" class="ExcelTable">
									                                <tr>
										                                <td class="ExcelHeaderCell" align="center">S.No.</td>
										                                <td class="ExcelHeaderCell" align="center"></td>
										                                <td class="ExcelHeaderCell" align="center">
										                                    <img src="../../assets/images/iTMS%20icons/Deleteicon.gif" onClick="DeleteAttribute()">
										                                </td>
										                                <td class="ExcelHeaderCell" align="center">Header Name</td>
										                                <td class="ExcelHeaderCell" align="center">Attribute Name</td>
									                                </tr>
								                                </table>
											                </td>
											            </tr>
											        </table>
											    </div>
											    <div id="divSelectAtt" style="display:block">
											        <table border="0" cellpadding="0" cellspacing="0" width="100%">
											            <tr>
											                <td>
											                    <table border="0" cellpadding="0" cellspacing="1" width="100%" id="tblAttSelect" class="ExcelTable">
									                                <tr>
										                                <td class="ExcelHeaderCell" align="center">S.No.</td>
										                                <td class="ExcelHeaderCell" align="center">
										                                </td>
										                                <td class="ExcelHeaderCell" align="center">Header Name</td>
										                                <td class="ExcelHeaderCell" align="center">Attribute Name</td>
									                                </tr>
								                                </table>
											                </td>
											            </tr>
											            <tr>
											                <td>
											                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
													                <tr>
														                <td valign="middle" class="ActionCell">
															                <p align="center">
																                <input type="button" value="Done" name="btnDone" class="ActionButton" onClick="SelectVal()">
															                    <input type="reset" value="Reset" name="B2" class="ActionButton">
																                <input type="button" value="Cancel" name="B3" class="ActionButton" onClick="window.close()">
														                </td>
													                </tr>
												                </table>
												            </td>
											            </tr>
											        </table>
											    </div>
											</td>
										</tr>
									   
									</table>
								</td>
								<td align="center">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
							    <td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							    <td>
								    
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



