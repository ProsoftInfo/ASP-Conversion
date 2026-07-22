<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MRApprovalEntry.asp
	'Module Name				:	Inventory (Transaction)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	July 20, 2005
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
<!-- #include File="../../include/DatabaseConnection.asp" -->
<!-- #include File="../../include/populate.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/CommonFunctions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>Material Requisition Approval</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<%
	Dim oDom,Root,HeaderNode,newElem,newElem1,newElem2,newElem3,newElem4,newElem5
	Dim iItemcode,iCtr,iEntNo,iAttribList,sCreatedBy,iuserid,rsUser,i,iClassCode,sAttList
	Dim dcrs,dcrs1
	dim sUnit,iMRNo,dMRDate,sMRType,sLotCardNo,sMachineNo,sCC
	dim sFinPeriod,Arr,dFrmDate,dToDate,sArrRefDetails,sAppRefType,sAppRefNoDate,sAppRefName,sAppRefNo,sAction
	Dim sIssToType,sIssToCode,sIssToSubCode,sIssToStr,sIssType
	sFinPeriod = session("Finperiod")
	Arr = split(sFinPeriod,":")
	dFrmDate = "01/04/"& Arr(0)
	dToDate = "31/03/"& Arr(1)
	'Declaration of Objects
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set dcrs1 = Server.CreateObject("ADODB.RecordSet")
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	sUnit = Session("organizationcode")
	'Response.Write "sUnit = "& sUnit
	iMRNo = Request.Form("mrs")
	iuserid = Session("userid")
    sAction = Request.Form("hAction")
	'Response.Write "iMRNo="&iMRNo
	'Response.Write "sUnit=:"& sUnit
	'Response.Write "sAction = "& sAction
	'To get User name
	sCreatedBy = Session("username")
	
	if trim(iMRNo)="" or IsNull(iMRNo) then
	%>
	    <script>
	        alert("Please selec any MR Action in List tab");
	        window.history.back(-1) 
	    </script>
	<%
	end if
	
	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT CONVERT(CHAR,MRSDATE,103),'',MRSTYPE,ISNULL(LOTCARDNO,''),ISNULL(MACHINENO,''),ISNULL(COSTCENTERHEAD,0),'',AppRefType,AppRefNo,ISSTOTYPE,ISSTOCODE,ISSTOSUBCODE,IsNull(IssueTypeCode,'GEN') FROM VWMRSLIST WHERE MRSFORUNIT = " & Pack(sUnit) & " AND MRSNUMBER = " & iMRNo & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	if not dcrs.EOF then
		dMRDate = trim(dcrs(0))
		sMRType = trim(dcrs(2))
		sLotCardNo = trim(dcrs(3))
		sMachineNo = trim(dcrs(4))
		sCC = trim(dcrs(5))
		sAppRefType = trim(dcrs(7))
		sAppRefNo = trim(dcrs(8))
		sIssToType = trim(dcrs(9))
		sIssToCode = trim(dcrs(10))
		sIssToSubCode = trim(dcrs(11))
		sIssToStr = IssuedToString(sIssToType,sIssToCode,sIssToSubCode)
		sIssType = trim(dcrs(12))
	end if
	dcrs.Close
    'Response.Write "Value = "& 	GetRefNoDate(sAppRefType,sAppRefNo)
    if trim(sAppRefType)<>"" then
        sArrRefDetails = split(GetRefNoDate(sAppRefType,sAppRefNo),",")
        sAppRefName = sArrRefDetails(0)
        sAppRefNoDate = sArrRefDetails(1)
    end if


	'Declaration of Objects
	Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

	Set Root = oDOM.createElement("root")
	oDOM.appendChild Root

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT DISTINCT ITEMDESCRIPTION,ITEMCODE,QUANTITYREQUESTED,STORESUOM,REQUIREDBY,ISNULL(REQUIREDVALUE,''),ISNULL(ITEMATTRIBUTES,''),ISNULL(ICOUNTER,0),ISNULL(ITEMREMARKS,''),ClassificationCode FROM VWMRSITEMDETAILS WHERE MRSNUMBER = " & iMRNo & ""
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	iCtr = 0
	if not dcrs.EOF then
		do while not dcrs.EOF
			iCtr = iCtr + 1
			iEntNo = dcrs(7)
			'Response.Write "iEntNo="&iEntNo
			IF cint(iEntNo)= 0 then iEntNo = iCtr
			iItemcode = dcrs(1)
			iAttribList = dcrs(6)
			iClassCode = dcrs(9)
			Set newElem = oDOM.createElement("ITEMDETAILS")
			newElem.setAttribute "ENTRYNO",iEntNo
			newElem.setAttribute "ITEMCODE",iItemcode
			newElem.setAttribute "CLASSCODE", iClassCode
			newElem.setAttribute "UNIT", sUnit
			newElem.setAttribute "ITEMNAME", ""
			newElem.setAttribute "UOM", dcrs(3)
			newElem.setAttribute "DECIMAL", ""
			newElem.setAttribute "DISPLAYED", "N"
			newElem.setAttribute "QTY", ""
			newElem.setAttribute "REQUIREDBY", dcrs(4)
			newElem.setAttribute "REQUIREDVALUE", ""
			newElem.setAttribute "ATTRIBUTELIST",iAttribList
			newElem.setAttribute "REMARKS", dcrs(8)

		'Added on Oct 23rd 2007	by Maheshwari to fetch Addspec values from additionaldetail table
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " SELECT I.WORKCENTERCODE,I.MACHINECENTERCODE,I.MIXCODE,I.QUANTITYISSUED,V.WORKCENTERNAME,V.MACHINECENTERNAME FROM "&_
					  " INV_T_MRSADDITIONALDETAILS AS I,VWWORKMACHINECENTER AS V WHERE I.MRSNUMBER = " & iMRNo & " AND ITEMCODE = "& iItemcode &" "&_
					  " AND V.WORKCENTERCODE = I.WORKCENTERCODE AND V.MACHINECENTERCODE = I.MACHINECENTERCODE"
			.ActiveConnection = con
			.Open
		end with

		set dcrs1.ActiveConnection = nothing
		if not dcrs1.EOF then
			Set newElem1 = oDOM.createElement("AddDet")
			newElem.appendchild newElem1
			do while not dcrs1.EOF
				Set newElem2 = oDOM.createElement("WorkCenter")
				newElem2.setAttribute "WCODE",dcrs1(0)
				newElem1.appendchild newElem2

				Set newElem3 = oDOM.createElement("MachineCenter")
				newElem3.setAttribute "MCODE",dcrs1(1)
				newElem3.setAttribute "QTY",dcrs1(3)
				newElem3.setAttribute "NAME",dcrs1(4)&" / "& dcrs1(5)
				newElem2.appendchild newElem3

				dcrs1.MoveNext
				loop
			end if
			dcrs1.Close

		'Added on 3rd April 2008 by Maheshwari to fetch Schedule Details  from Inv_T_MRSItemSchedules table
		with dcrs1
			.CursorLocation = 3
			.CursorType = 3
			.Source = " SELECT SCHEDULENO,SCHEDULETYPE,SCHEDULEDON,SCHEDULEDQTY FROM INV_T_MRSITEMSCHEDULES  WHERE "&_
					  " MRSNUMBER = " & iMRNo & " AND ITEMCODE = "& iItemcode &" AND ITEMENTRYNO = "& iEntNo &" AND "&_
					  " ORGANISATIONCODE = "& sUnit &" "
			.ActiveConnection = con
			.Open
		end with
	'	Response.Write "dcrs1(0)="&dcrs1.Source

		set dcrs1.ActiveConnection = nothing


		if not dcrs1.EOF then
			Set newElem4 = oDOM.createElement("Schedule")
			newElem4.setAttribute "STYPE",dcrs1(1)
			newElem4.setAttribute "SVALUE", dcrs1(2)
			newElem4.setAttribute "ITEMCODE", iItemcode
			newElem4.setAttribute "CLASSCODE","0"
			newElem4.setAttribute "SCHENTRYNO",iEntNo
			newElem.appendchild newElem4

			do while not dcrs1.EOF

				'Have to create Scheduledetails node
				Set newElem5 = oDOM.createElement("ScheduleDetails")
				newElem5.setAttribute "SNO", dcrs1(0)
				newElem5.setAttribute "NEED", dcrs1(2)
				newElem5.setAttribute "QTY", dcrs1(3)
				newElem5.setAttribute "TYPE", dcrs1(1)
				newElem4.appendchild newElem5


				dcrs1.MoveNext
			loop

		end if
		dcrs1.close
		Root.appendChild newElem
		dcrs.MoveNext
		loop
	end if
	dcrs.Close

	oDOM.Save server.MapPath("../temp/transaction/MRAPPROVAL"&Session.SessionID&".xml")
%>

<script type="application/xml" data-itms-xml-island="1" id="ItemData"></script>
<script type="application/xml" data-itms-xml-island="1" id="UoMData" data-src="../../inventory/xmldata/Uom.xml"></script>
<script type="application/xml" data-itms-xml-island="1" id="OutData" data-src="<%="../temp/transaction/MRAPPROVAL"&Session.SessionID&".xml"%>"></script>
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<script type="application/xml" data-itms-xml-island="1" id="PartyData"><Root/></script>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/RoundOff.js"></script>
<script LANGUAGE=javascript SRC="../scripts/Date.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/TempItem.js"></script>
<script Language="javascript" Src="../../scripts/RefTypePop.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/mrEntryModern.js"></SCRIPT>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="Init('<%=dMRDate%>')">
	<form method="POST" name="formname">
	<input type=hidden name="hCreatedBy" value="<%=Session("userid")%>">
	<input type=hidden name="hUnit" value="<%=sUnit%>">
	<input type=hidden name="hMRNo" value="<%=iMRNo%>">
	<input type=hidden name="hFrmDate" value="<%=dFrmDate%>">
	<input type=hidden name="hToDate" value="<%=dToDate%>">
	<input type=hidden name="hMRDate" value="<%=dMRDate%>">
	<input type=hidden name="hAction" value="<%=sAction%>">
	
	<input type="hidden" name="hIssueToType" value="<%=sIssToType%>">
	<input type="hidden" name="hIssueToCode" value="<%=sIssToCode%>">
	<input type="hidden" name="hIssueToSubCode" value="<%=sIssToSubCode%>">
	
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="center" class="PageTitle" height="20">
				<p align="center">Material Requisition
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%"  >
				    		    <tr>
						<td height="20" valign="bottom">
							<table border="0" cellpadding="0" cellspacing="0" >
								<tr>
								   	<td class="TabCell" valign="bottom" width="50">
									    <table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
										    <tr><a href="MRSMGMTLIST.asp">
											    <td align="center">List
											    </td>
										    </tr>
									    </table>
								    </td>
									<td class="TabCell" valign="bottom" width="90">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr><a href="MRGENERATIONENTRY.asp">
												<td align="center">Basic
												</td></a>
											</tr>
										</table>
									</td>
									
								    <td class="TabCurrentCell" valign="bottom" align="center" width="145">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="TabCurrentTable" onMouseOver="tabrollover(this)" onMouseOut="tabrollout(this)">
											<tr>
												<td align="center">Edit/Approval
												</td>
											</tr>
										</table>
									</td>
									<td class="TabCellEnd" valign="bottom" align="left">
										    &nbsp;
								    </td>
								</tr>
							</table>
						</td>
                	</tr>
					<tr>
						<td class="TabBody">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>
								<tr>
									<td align="center">
									</td>
									<td width="100%" colspan="2">
										<div align="left">
											<table border="0" cellspacing="0" cellpadding="0" width="100%">
											    <tr>
												    <td class="FieldCellSub" style="width:125px">Requested By</td>
												    <td class="FieldCellSub" valign="top">
													    <select size="1" name="selIssueTo" class="FormElem"  onChange="popIssueTo()">
														    <option value="select">Select</option>
													    <%	'Calling the Function which populates Issue TO
														    populateIssueToSel(sUnit)
													    %>
													    </select>
												    </td>
												    <td class="FieldCellSub" >MR Date</td>
												    <td class="FieldCellSub" valign="middle">
													    <object id="ctlCDDate" onBlur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD"     codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89" height="20" class="formelem" viewastext>
														    <param name="_ExtentX" value="2355">
														    <param name="_ExtentY" value="529">
													    </object>
												    </td>
											    </tr>
											    <tr>
                                                    <td class="FieldCellSub" style="width:125px">Reference Name</td>
												    <td class="FieldCellSub">
													    <span class="DataOnly" align=center>
													        <%
													            if trim(sAppRefName)<>"" then
													                Response.Write sAppRefName
													            else
													                Response.Write "None"
													            end if
													        %>
													    </span>
											        </td>
											        <td class="FieldCellSub">Created By</td>
												    <td class="FieldCellSub">
													    <span class="dataonly"><%=sCreatedBy%></span>
												    </td>
											    </tr>
                                                <tr>
                                                    <td class="FieldCellSub" style="width:125px">Reference No - Date</td>
												    <td class="FieldCellSub">

													    <span class="DataOnly" align=center>
													    <%
													        if trim(sAppRefNoDate)<>"" then
													            Response.Write sAppRefNoDate
													        else
													            Response.Write "NA"
													        end if
													    %>
													    </span>
										            </td>
												    <td class="FieldCellSub">Cost Center</td>
												    <td class="FieldCellSub" valign="top">
													    <select size="1" name="selCC" class="FormElem">
														    <option value="select">Select</option>
													    <%	'Calling the Function which populates Cost Center List
														    populateCostCenter
													    %>
													    </select>
												    </td>
                                                </tr>
                                                <tr>
                                                    <td class="FieldCellSub">Issue Type</td>
                                                    <td class="FieldCellSub">
                                                        <select id="cmbIssType" class="FormElem">
                                                            <option value="SEL" <%if sIssType="SEL" then Response.write "Selected" %>>Select</option>
                                                            <%
                                                                sQuery = "Select ReceiptIssueTypeCode,ReceiptIssueTypeDesc from APP_M_ReceiptIssueTypes where ApplicableFor in ('B','I')"
                                                                dcrs.open sQuery,con
                                                                if not dcrs.eof then
                                                                    do while not dcrs.eof
                                                                        if trim(sIssType)=trim(dcrs(0)) then
                                                                            response.write "<option value="& trim(dcrs(0)) &" selected>"& trim(dcrs(1)) &"</option>"
                                                                        else
                                                                            response.write "<option value="& trim(dcrs(0)) &">"& trim(dcrs(1)) &"</option>"
                                                                        end if 
                                                                        dcrs.movenext
                                                                    loop
                                                                end if
                                                                dcrs.close
                                                            %>
                                                        </select>
                                                    </td>
                                                </tr>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack"></td>
								</tr>

								<tr>
									<td align="center"></td>
									<td width="100%" colspan="2">
										<div class="frmBody" id="frm1" style="width: 100%; height:300;">
											<table border="0" cellspacing="1" class="ExcelTable" width="100%" id=tblLot>
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10"><p align="center">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center" >
														<a href="#"><img border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" alt="Select the Item (s) to be rejected" height="15"></a>
													</td>
													<td class="ExcelHeaderCell" align="center" width=200>Item Description</td>
													<td class="ExcelHeaderCell" align="center" >Quantity</td>
													<td class="ExcelHeaderCell" align="center" width="50">UoM</td>
													<td class="ExcelHeaderCell" align="center" >Required By</td>
													<!--<td class="ExcelHeaderCell" align="center" width="50">Add Spec</td>-->
													<td class="ExcelHeaderCell" align="center" width="50">Stock</td>
												</tr>
											<%Dim sOptName,sItemName,sArrList,sQuery,rsTemp
											set rstemp = Server.CreateObject("ADODB.Recordset")
											iCtr = 0
												with dcrs
													.CursorLocation = 3
													.CursorType = 3
													.Source = "SELECT DISTINCT ITEMDESCRIPTION,ITEMCODE,QUANTITYREQUESTED,STORESUOM,REQUIREDBY,REQUIREDVALUE,ISNULL(ITEMATTRIBUTES,''),ISNULL(ICOUNTER,0),ClassificationCode FROM VWMRSITEMDETAILS WHERE MRSNUMBER = " & iMRNo & " "
													.ActiveConnection = con
													.Open
												end with
												'Response.Write dcrs.source
												set dcrs.ActiveConnection = nothing
												iAttribList = ""

													do while not dcrs.EOF
														iCtr = iCtr + 1
														sAttList = dcrs(6)
														iEntNo = dcrs(7)
														IF cint(iEntNo) = 0 then iEntNo = iCtr
														
														if trim(sAttList)<>"" and trim(sAttList)<>"0" and Trim(sAttList)<>"NULL" then
															iAttribList = split(sAttList,":")
															'Response.Write "iAttribList="&iAttribList(0)
															IF trim(iAttribList(0)) <> "" then
																sArrList = split(iAttribList(0),"#")
																if UBound(sArrList)>0 then
																	sOptName = FunAttribName(iAttribList(0))
																else
																	sQuery = "Select OptionName from Inv_M_ItemTypeOptions where OptionValue = "& iAttribList(0)
																	rsTemp.open squery,con
																	if not rstemp.eof then
																		sOptName = " ["&trim(rstemp(0))&"]"
																	end if
																	rsTemp.Close 
																end if
															Else
																sOptName =""
															End IF
														else
															sOptName =""
														end if 'if trim(sAttList)<>"" and Trim(sAttList)<>"NULL" then
														if trim(sOptName)<>"" then
														    sItemName = trim(dcrs(0)) & sOptName
														else
														    sItemName = trim(dcrs(0))
														end if
											%>
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10"><p align="center"><%=iCtr%></td>
													<td class="ExcelInputCell" align="center" width="10">
														<input type="checkbox" name="chkDeleteA<%=trim(dcrs(1))%>A<%=trim(dcrs(8))%>A<%=iEntNo%>" value="<%=iEntNo%>" class="Formelem" style="text-align=right">
													</td>
													<td class="ExcelDisplayCell" align="left" width=200><%=trim(dcrs(0))%> <%=sOptName%></td>
													<td class="ExcelInputCell" align="right" width="50"><input type="text" name="txtQtyZ<%=trim(dcrs(1))%>Z<%=trim(dcrs(8))%>Z<%=iEntNo%>" size="12" value="<%=trim(dcrs(2))%>" class="Formelem" style="text-align=right" onkeypress="DoKeyPress('<%=UoMDecimal(trim(dcrs(3)))%>',7,3)"></td>
													<td class="ExcelDisplayCell" align="center" width="50"><%=trim(dcrs(3))%></td>
													<td class="ExcelFieldCell" align="left" width=30>
													
													<%if trim(sAttList)<>"" and trim(sAttList)<>"0" and Trim(sAttList)<>"NULL" then%>
													    <select size="1" name="selSchZ<%=trim(dcrs(1))%>Z<%=trim(dcrs(8))%>Z<%=iEntNo%>" class="FormElem" onChange="CheckSch(this,'<%=FormatDate(date())%>','<%=trim(dcrs(3))%>','<%=iAttribList(0)%>')">
													<%else%>
														<select size="1" name="selSchZ<%=trim(dcrs(1))%>Z<%=trim(dcrs(8))%>Z<%=iEntNo%>" class="FormElem" onChange="CheckSch(this,'<%=FormatDate(date())%>','<%=trim(dcrs(3))%>','')">
													<%END IF%>
															<option value="select">Select</option>
														<%if trim(dcrs(4)) = "I" then %>
															<option value="I" SELECTED>Immediate</option>
															<option value="W">Within x Days</option>
															<option value="D">Specific Date</option>
															<option value="S">Scheduled</option>
														<%	elseif 	trim(dcrs(4)) = "W" then %>
															<option value="I">Immediate</option>
															<option value="W" SELECTED>Within x Days</option>
															<option value="D">Specific Date</option>
															<option value="S">Scheduled</option>
														<%	elseif 	trim(dcrs(4)) = "D" then %>
															<option value="I">Immediate</option>
															<option value="W">Within x Days</option>
															<option value="D" SELECTED>Specific Date</option>
															<option value="S">Scheduled</option>
														<%	elseif 	trim(dcrs(4)) = "S" then %>
															<option value="I">Immediate</option>
															<option value="W">Within x Days</option>
															<option value="D">Specific Date</option>
															<option value="S" SELECTED>Scheduled</option>
														<%	end if %>
													    </select>
													</td>
													<!--<td class="ExcelFieldCell" align="center">
														<%if trim(sAttList)<>"" and trim(sAttList)<>"0" and Trim(sAttList)<>"NULL" then%>
														<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor:hand" alt="Additional Specs" width="11" height="11" onClick="GetAddDetails('<%=trim(dcrs(1))%>','<%=trim(dcrs(8)) %>','<%=sUnit%>','<%=iEntNo%>','<%=trim(iAttribList(0)) %>')">
														<%ELSE%>
														<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor:hand" alt="Additional Specs" width="11" height="11" onClick="GetAddDetails('<%=trim(dcrs(1))%>','<%=trim(dcrs(8)) %>','<%=sUnit%>','<%=iEntNo%>','')">
														<%END IF%>
													</td>-->
													<td class="ExcelFieldCell" align="center">
													<%if trim(sAttList)<>"" and trim(sAttList)<>"0" and Trim(sAttList)<>"NULL" then%>
														<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor:hand" alt="Stock Details" width="11" height="11" onClick="DisplayStock('<%=trim(dcrs(1))%>','<%=trim(dcrs(8)) %>','<%=sUnit%>','<%=iEntNo%>','<%=replace(replace(sItemName,"'",""),chr(34),"~~")%>','<%=trim(iAttribList(0)) %>')">
													<%ELSE%>
														<img border="0" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" style="cursor:hand" alt="Stock Details" width="11" height="11" onClick="DisplayStock('<%=trim(dcrs(1))%>','<%=trim(dcrs(8)) %>','<%=sUnit%>','<%=iEntNo%>','<%=replace(replace(sItemName,"'",""),chr(34),"~~")%>','')">
													<%END IF%>
														<!--input type="button" value="View" name="B6" class="ActionButtonX" onClick="DisplayStock('<%=trim(dcrs(1))%>','0','<%=sUnit%>')"-->
													</td>
												</tr>

											<%
													dcrs.MoveNext
													loop
													dcrs.Close
											%>
											</table>
										</div>
									</td>
									<td align="center"></td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>

								<tr>
								    <td align="center"></td>
								    <td class="FieldCellSub"> Remarks</td>
									<td class="FieldCellSub">
										<textarea name="txtRemarks" cols="100" class="Formelem"></textarea>
									</td>
									<td align="center"></td>
								</tr>
								<tr>
									<td align="center" colspan="4" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
									<td valign="top" colspan="2">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<p align="center">
													<input type="button" value="Save" name="BtnSubmit" class="ActionButton" onClick="CheckSubmit('<%=FormatDate(date)%>')">
 													<input type="reset" value="Reset" name="B1" class="ActionButton">
 													&nbsp;
												</td>
											</tr>

										</table>
									</td>
									<td align="center">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" colspan="4" class="BottomPack">
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
</body>
</html>

<%
	' Function to populate Usage
	Function populateUsage(sIssuedFor)
		' Declaration of variables
		Dim dcrs,sUsageCode,sUsageDesc
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT ISSUEDFORCODE,ISSUEDFORDESCRIPTION FROM INV_M_ISSUEDFOR WHERE ISSUEDFORCODE <> 'INV' ORDER BY ISSUEDFORCODE"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUsageCode = dcrs(0)
		set sUsageDesc = dcrs(1)

		Do While Not dcrs.EOF
			if sIssuedFor = trim(sUsageDesc) then
				Response.Write("<OPTION VALUE="""&trim(sUsageCode)&""" SELECTED>"&trim(sUsageDesc)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(sUsageCode)&""">"&trim(sUsageDesc)&"</OPTION>" &vbcrlf)
			end if
			dcrs.MoveNext
		Loop
		dcrs.Close

	End Function
%>
<%
	' Function to populate the Cost Center list
	Function populateCostCenter()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT COSTCENTERHEAD,CCACCOUNTDESCRIPTION FROM VWORGCOSTCENTER WHERE OUDEFINITIONID = " & Pack(sUnit) & " AND USEABLE = 1 ORDER BY COSTCENTERHEAD"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
			if sCC = trim(stypID) then
				Response.Write("<OPTION VALUE="""&trim(stypID)&""" SELECTED>"&trim(stypName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
			end if
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing

	End Function
%>
