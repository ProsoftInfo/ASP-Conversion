<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	mrsPRSchedulePoP.asp
	'Module Name				:	Inventory (MRS PR Schedule)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	December 05, 2003
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<%
dim oDom,Root,PageNode,HeaderNode,PGNode,objfs

dim dcrs,sql,sItemTypeName,sUnitName,sUsageName,sno,iItem,iClass
dim i,ShortDesc,sSchType,iQty,bFlag,iCounter,selected,EntryNode
dim arrTemp,iMRSNo,sOrgID,sOrgName,dMRSDate,sItemName,iEntNo

set dcrs = server.CreateObject("Adodb.recordset")
set oDom = server.CreateObject("Microsoft.xmlDom")
Set objfs = CreateObject("Scripting.FileSystemObject")

sSchType="D"
arrTemp = split(trim(Request.QueryString("sTemp")),":")
iQty	= arrTemp(0)
iClass	= arrTemp(2)
iItem	= arrTemp(1)
iMRSNo = arrTemp(3)
iEntNo = arrTemp(4)

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objfs = CreateObject("Scripting.FileSystemObject")
if objfs.FileExists(Server.MapPath("../temp/transaction/MRS"&iMRSNo&".xml")) then
	oDOM.Load server.MapPath("../temp/transaction/MRS"&iMRSNo&".xml")
	Set Root = oDOM.documentElement
	if Root.HaschildNodes() then
		For Each HeaderNode In Root.childNodes
			if StrComp(HeaderNode.nodeName,"MRSHeader") = 0 then
				dMRSDate = HeaderNode.Attributes.Item(1).nodeValue
				sOrgID = HeaderNode.Attributes.Item(2).nodeValue
				sOrgName = HeaderNode.Attributes.Item(3).nodeValue
			end if
		next
	end if
end if
sItemName = ItemDisplay(iItem,iClass)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - PR Schedule Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/MRCreateSch.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=iItem%>','<%=iClass%>','<%=iQty%>','<%=iEntNo%>','PRSchedule')">

<form method="POST" name="formname">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr>
		<td align="center" class="TopPack">
		</td>
    </tr>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Purchase Requsition - Schedule
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
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
								</td>
								<td valign="top" class="FieldCell" width="100%">
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
                                            <td class="FieldCellSub">MRS Date</td>
                                            <td class="FieldCellSub"><span class="DataOnly"><%=dMRSDate%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Unit Name</td>
                                            <td class="FieldCellSub"><span class="DataOnly"><%=sOrgName%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub"> Description</td>
                                            <td class="FieldCellSub"><span class="DataOnly" id="idItemName"><%=sItemName%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Item Quantity</td>
                                            <td class="FieldCellSub"><span class="DataOnly"><%=iQty%>&nbsp;</span> - <span class="DataOnly"><%=DisplayUoM(sOrgID,iItem,iClass)%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Schedule Type</td>

                                            <td class="FieldCellSub">
												<SELECT NAME="selSchtype" class="formelem" size="1" onChange="setMax(this)">
													<option value="select">Select</option>
													<option value="DS">Date</option>
													<option value="MS">MonthYear</option>
													<option value="WS">WeekYear</option>
													<option value="YS">MonthWeekYear</option>
												</SELECT>
											</td>
                                        </tr>
							       </table>
								</td>
								<td align="center" class="ClearPixel">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
                        <tr>
							<td align="center" class="ClearPixel"></td>
							<td valign="top" class="FieldCell" width="100%" align="center">
								<DIV class=frmBody id=frm3 style="width: 220; height:275;">
                                    <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                        <tr>
											<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
											<td class="ExcelHeaderCell" align="center">Need by</td>
											<td class="ExcelHeaderCell" align="center">Quantity</td>
                                        </tr>
										<%for iCounter=1 to 12%>
                                        <tr>
											<td class="ExcelSerial" align="center" width="10"><%=iCounter%></td>
											<td class="ExcelInputCell"><input type="text" name="txtD<%=iCounter%>" size="12" maxlength=10 class="FormElem"></td>
											<td class="ExcelInputCell"><input type="text" name="txtQ<%=iCounter%>" size="12" maxlength=10 class="FormElem"></td>
                                        </tr>
                                        <%next%>
                                    </table>
								</div>
							</td>
							<td align="center" class="ClearPixel"></td>
                        </tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
							</tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" class="FieldCell">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
                                                    <input type="button" value="Done" name="B4" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date())%>')">
                                                    <input type="reset" value="Reset" name="B5" class="ActionButton">
											</td>
										</tr>
									</table>
								</td>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" class="BottomPack" colspan="3">
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
	Function DisplayUoM(sOrgID,iItem,iClass)
		' Declaration of variables
		Dim dcrs,sUoMDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMORGMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMDesc
		end if
		dcrs.Close
	End Function
%>
