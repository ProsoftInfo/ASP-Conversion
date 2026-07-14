<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	AddEntryDetails.asp
	'Module Name				:	Inventory (Issue Additional Details)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	February 06, 2004
	'Modified By				:	S.Maheshwari
	'Modified On				:	October 22, 2007
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
<!--#include virtual="/include/UoMDecimal.asp"-->
<!--#include virtual="/include/ItemDisplay.asp"-->
<%
' Declaration of variables
Dim dcrs,dcrs1,dcrs2,iCtr,bexists
'Declaration of Objects
iCtr = 0
Set dcrs = Server.CreateObject("ADODB.RecordSet")
Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
Set dcrs2 = Server.CreateObject("ADODB.RecordSet")

dim iItem,iClass,sItemName,sUsage,sClassName,iEntNo,iAttribList,sOptName
dim arrTemp,sOrgID,sOrgName,iQty,arrUoM,sUoMCode,sUoM

arrTemp = split(trim(Request.QueryString("sTemp")),"|")
iClass = arrTemp(0)
iItem = arrTemp(1)
sOrgID = arrTemp(2)
sUsage = arrTemp(3)
iQty = arrTemp(4)
iEntNo = arrTemp(5)
iAttribList = arrTemp(6)
If iAttribList <> "" then
	sOptName = FunAttribName(iAttribList)
else
	sOptName = ""
End IF
with dcrs2
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT DISTINCT GROUPNAME,ITEMDESCRIPTION,ORGUNITSHORTDESCRIPTION FROM VWITEM WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ""
	.ActiveConnection = con
	.Open
end with
'Response.Write dcrs2.Source
set dcrs2.ActiveConnection = nothing

if not dcrs2.EOF then
	sClassName = trim(dcrs2(0))
	sItemName =  trim(dcrs2(1))
	sOrgName = trim(dcrs2(2))
end if
dcrs2.Close

'sItemName = ItemDisplay(iItem,iClass)
sItemName = sItemName & sOptName
arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
sUoMCode = arrUoM(0)
sUoM = arrUoM(1)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - Additional Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="Data">
<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="MCData">
<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="WGData">
<Root/>
</script>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ModalReturnCompat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/AddDetails.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="fnInit('<%=sOrgID%>','<%=iClass%>','<%=iItem%>','<%=sUsage%>')">
<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0" class=PopupTable>
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Additional Details
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
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td class="FieldCell">Item Name</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sItemName%>&nbsp;</span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Unit Name</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sOrgName%></span>
											</td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCell">Quantity</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly" id="idQty"><%=iQty%>&nbsp;</span>&nbsp;
												<span class="DataOnly"><%=sUoM%></span>
											</td>
										</tr>
										<tr>
										    <td class="FieldCell" valign="top"> Work Group</td>
										    <td class="FieldCellSub">
												<select size="1" name="selWG" class="FormElem" onChange="GetWC(this)">
													<option value="select">Select</option>
											<%	'Calling the Function which populates Work Center List
												populateWorkGroups
											%>
												</select>
											</td>
										</tr>
										<tr>
										    <td class="FieldCell" valign="top"> Work Center</td>
										    <td class="FieldCellSub">
												<select size="1" name="selWC" class="FormElem" onChange="GetMC(this)">
													<option value="select">Select</option>
											<%	'Calling the Function which populates Work Center List
												'populateWorkCenters
											%>
												</select>
											</td>
										</tr>
									<%'	Response.Write sUsage
									if sUsage = "MAT" then %>
										<tr>
										    <td class="FieldCell" valign="top"> Machine Center</td>
										    <td class="FieldCellSub">
												<select size="1" name="selMC" class="FormElem" onChange="GetMCDetails(this)">
													<option value="select">Select</option>
												</select>
											</td>
										</tr>
										<tr>
										    <td class="FieldCell" valign="top"> Machine Details</td>
										    <td class="FieldCellSub">
												<span class="DataOnly" ID="MCModel"></span>&nbsp<span class="DataOnly" ID="MCSNo"></Span>
											</td>
										</tr>
										<tr>
											<td class="FieldCell">Quantity</td>
											<td class="FieldCellSub">
												<input type="text" name="txtQty" size="12" onkeypress="DoKeyPress('<%=UoMDecimal(sUoMCode)%>',7,3)" class="FormElem">
												<input type="button" value=" Add " name="B3" class="AddButtonX" onClick="CheckEntry()">
											</td>
										</tr>
									<%	end if %>
									</table>
								</td>
								<td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
						<%	if sUsage = "MAT" then %>
                            <tr>
								<td align="center"></td>
								<td width="100%">
									<div class="frmBody" id="frm2" style="width: 100%; height:130;">
										<table border="0" cellspacing="1" id="tblData" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="10">
													<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" alt="Select the Item (s) to be rejected" height="15" Onclick="DeleteEntry()"></a>
												</td>
												<td class="ExcelHeaderCell" align="center">Work / Machine Center</td>
												<td class="ExcelHeaderCell" align="center" width="100">Quantity</td>
											</tr>
										</table>
									</div>
								</td>
								<td align="center"></td>
                            </tr>

							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
					<%	end if %>
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
	' Function to populate Store
	Function DisplayUoM(sOrgID,iClass,iItem)
		' Declaration of variables
		Dim dcrs,sUoMDesc,sUoMCode
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")

		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT UOMCODE,UOMSHORTDESCRIPTION FROM MS_UNITOFMEASUREMENT WHERE UOMCODE = (SELECT STORESUOM FROM INV_M_ITEMMASTER WHERE ITEMCODE = " & iItem & " AND CLASSIFICATIONCODE = " & iClass & " AND ORGANISATIONCODE = " & Pack(sOrgID) & ")"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
		set sUoMCode = dcrs(0)
		set sUoMDesc = dcrs(1)
		if Not dcrs.EOF then
			DisplayUoM = sUoMCode&":"&sUoMDesc
		end if
		dcrs.Close
	End Function
%>

<%
	' Function to populate the Work Center list
	Function populateWorkCenters()
		' Declaration of variables
		Dim dcrs,sWCID,sWCName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT WORKCENTERCODE,WORKCENTERNAME FROM VWWORKMACHINECENTER WHERE ORGANISATIONCODE = " & Pack(sOrgID) & " ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sWCID = dcrs(0)
		set sWCName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sWCID)&""">"&trim(sWCName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close

	End Function
%>
<%
	' Function to populate the Work Groups list
	Function populateWorkGroups()
		' Declaration of variables
		Dim dcrs,sWCID,sWCName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT DISTINCT WORKGROUPCODE,WORKGROUPNAME FROM PRD_M_WORKGROUP ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
'Response.Write dcrs.Source
		set sWCID = dcrs(0)
		set sWCName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				Response.Write("<OPTION VALUE="""&trim(sWCID)&""">"&trim(sWCName)&"</OPTION>" &vbcrlf)
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close

	End Function
%>
