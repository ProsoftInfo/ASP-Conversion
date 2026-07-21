<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	mrsIssueQtyParaPoP.asp
	'Module Name				:	Inventory (MRS Issue)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 18, 2003
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
' Declaration of variables
Dim dcrs,dcrs1,iCtr,bexists
'Declaration of Objects
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")

dim oDom,Root,PageNode,HeaderNode,PGNode,objfs,sIssueCode

dim sql,sItemTypeName,sUnitName,sUsageName,iItem,iClass,sItemName
dim arrTemp,iMRSNo,sOrgID,sOrgName,dMRSDate,sAttrValue,sAttrOpValue,sItmType
dim sOptName,iEntNo
set oDom = server.CreateObject("Microsoft.xmlDom")
Set objfs = CreateObject("Scripting.FileSystemObject")
'Response.Write Request.QueryString("sTemp")
arrTemp = split(trim(Request.QueryString("sTemp")),":")
iClass	= arrTemp(1)
iItem	= arrTemp(0)
iMRSNo = arrTemp(2)
iEntNo = arrTemp(3)
sOptName = arrTemp(4)

'sItemName = ItemDisplay(iItem,iClass)
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT ITEMDESCRIPTION FROM VWITEM WHERE ITEMCODE = " & iItem & ""
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if not dcrs.EOF then
	sItemName = trim(dcrs(0))
end if
dcrs.close
sItemName = sItemName & sOptName
with dcrs
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT DISTINCT MRSFORUNIT,ORGUNITSHORTDESCRIPTION,CONVERT(CHAR,MRSDATE,103),ISNULL(MRSCODE,MRSNUMBER) FROM VWMRSLIST WHERE MRSNUMBER = " & iMRSNo & " AND ISNULL(ICOUNTER,0) = "& iEntNo &" "
	.ActiveConnection = con
	.Open
end with
set dcrs.ActiveConnection = nothing

if not dcrs.EOF then
	sOrgID = trim(dcrs(0))
	sOrgName = trim(dcrs(1))
	dMRSDate = trim(dcrs(2))
	'sItmType = trim(dcrs(3))
	sIssueCode = trim(dcrs(3))
end if
dcrs.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS : MR Issue - Quality Parameters</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrsIssueQtyPara.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=iItem%>','<%=iClass%>')">

<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
    <tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Material Requisition - Quality Parameters
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
                                            <td class="FieldCell">MR No. - Date&nbsp;</td>
                                            <td class="FieldCellSub" colspan="3"><span class="DataOnly"><%=sIssueCode%>&nbsp;</span> - <span class="DataOnly"><%=dMRSDate%></span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Unit Name</td>
                                            <td class="FieldCellSub" colspan="3"><span class="DataOnly"><%=sOrgName%>&nbsp;</span></td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell"> Description</td>
                                            <td class="FieldCellSub" colspan="3"><span class="DataOnly" id="idItemName"><%=sItemName%>&nbsp;</span></td>
                                        </tr>
                                    </table>
								</td>
								<td align="center" class="ClearPixel">
								</td>
							</tr>
							<tr>
								<td align="center" class="MiddlePack" colspan="3">
								</td>
							</tr>
                            <tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
									<DIV class=frmBody id=frm1 style="width: 100%; height:180">
                                        <table border="0" cellspacing="1" class="ExcelTable" width="100%">
                                            <tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center">Quality Controls</td>
												<td class="ExcelHeaderCell" align="center">Value</td>
                                            </tr>
											<%
												iCtr = 0
												with dcrs
													.CursorLocation = 3
													.CursorType = 3
													.Source = "SELECT ATTRIBUTEDATA,OPTIONVALUE,ITEMTYPEATTRIBUTEID FROM INV_T_MRSITEMSPECS WHERE MRSNUMBER = " & iMRSNo & " AND ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
													.ActiveConnection = con
													.Open
												end with
												set dcrs.ActiveConnection = nothing

												if not dcrs.EOF then
													do while not dcrs.EOF
														iCtr = iCtr + 1
														sAttrValue = trim(dcrs(0))
														sAttrOpValue = trim(dcrs(1))

														with dcrs1
															.CursorLocation = 3
															.CursorType = 3
															.Source = "SELECT ITEMTYPEATTRIBUTENAME,ITEMTYPEATTRIBUTETYPE FROM INV_M_ITEMTYPEATTRIBUTES WHERE HEADERID = 1 AND ITEMTYPEID = " & Pack(sItmType) & " AND ITEMTYPEATTRIBUTEID = " & trim(dcrs(2)) & " ORDER BY ITEMTYPEATTRIBUTENAME"
															.ActiveConnection = con
															.Open
														end with
														set dcrs1.ActiveConnection = nothing
														if not dcrs1.EOF then
											%>
                                            <tr>
												<td class="ExcelSerial" align="center"><%=iCtr%></td>
												<td class="ExcelDisplayCell"><%=trim(dcrs1(0))%></td>
												<td class="ExcelDisplayCell">
													<%	if lcase(trim(dcrs1(1))) = "options" then %>
														<%=populateOptionList(trim(dcrs(2)),sAttrOpValue)%>
													<%	else %>
														<%=sAttrValue%>
													<%	end if %>
												</td>
                                            </tr>
											<%			end if
														dcrs1.Close
													dcrs.MoveNext
													loop
												else
											%>
                                            <tr>
												<td class="ExcelDisplayCell" align=center colspan=3><b>No Quality Parameters Defined.</b></td>
											</tr>
											<%
												end if
												dcrs.Close
												con.Close
												set con = nothing
											%>
                                        </table>
									</div>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top" width="100%">
									<%	'if bexists then %>
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                    <input type="button" value="OK" name="B3" class="ActionButton" onClick="window.close()">
											</td>
										</tr>
									</table>
									<%	'end if %>
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
	' Function to populate Option List
	Function populateOptionList(iAID,iOpVal)
		' Declaration of variables
		Dim dcrs,iOptVal,sOptName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT OPTIONVALUE,OPTIONNAME FROM INV_M_ITEMTYPEOPTIONS WHERE ITEMTYPEATTRIBUTEID = " & iAID & " AND OPTIONVALUE = " & iOpVal & ""
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set iOptVal = dcrs(0)
		set sOptName = dcrs(1)

		if Not dcrs.EOF then
			populateOptionList = trim(sOptName)
		else
			populateOptionList = ""
		end if
		dcrs.Close

	End Function
%>
