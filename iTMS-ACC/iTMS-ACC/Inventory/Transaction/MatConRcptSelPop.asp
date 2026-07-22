<%@ LANGUAGE =VBSCRIPT%>
<% OPTION EXPLICIT %>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	MatConRcptSelPop.asp
	'Module Name				:	Inventory (Issues- Consumption)
	'Author Name				:	Ragavendran
	'Created On					:	May 24,2013
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
<!-- #include file="../../include/DatabaseConnection.asp" -->
<!-- #include file="../../include/populate.asp" -->
<!-- #include File="../../include/UoMDecimal.asp" -->
<!-- #include File="../../include/ItemDisplay.asp" -->
<!-- #include File="../../include/CommonFunctions.asp" -->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<%
	dim dcrs,dcrs1,iItem,iClass,sOrgID,iMRSNo,arrTemp,iIssNo
	dim sItmName,iQty,sTemp,sAlt,iLot
	dim sRead,iCtr,dIssDate,sTitle,sVar,iLineNo
	dim arrUoM,sUoMDesc,sUoMCode,sType,iDINo,sSQL,iAccHead,sRecptNum
	Dim sAttributeList,sCallFrom,sIssueCode,iNoOfPacks
	Dim iIntRcptNo,iIntItemCode,sIntAttList,sRcptNum,iOutPutQty,iRcptQty
	Dim sItemCode,sClassCode,sCatCode,sItemName,sClassName,iRowCount
	Dim sFromDate,sToDate,sFinPeriod,sFinArr

	Const iPageSize=10	'How many records to show
    Dim iCurrentPage	'Current Page No.
    Dim iTotPage		'Total No. of pages if iPageSize records are displayed = per page.
    Dim iPageCtr		'Counter
	Dim lnPage

    iCurrentPage = Request.Form("hPageSelection")
    if iCurrentPage = "" or iCurrentPage = "0" then iCurrentPage = "1"
    iCurrentPage = CInt(iCurrentPage)

	iAccHead = 0
	set dcrs = Server.CreateObject("ADODB.Recordset")
	set dcrs1 = Server.CreateObject("ADODB.Recordset")

	iCtr = 0
	Response.Write "<font color=red>"
	'Response.Write Request.QueryString("sTemp")

	arrTemp = split(trim(Request.QueryString("sTemp")),":")
	sType   = arrTemp(0)
	iIssNo  = arrTemp(1)
	iItem	= arrTemp(2)
	iClass	= arrTemp(3)
	sOrgID	= arrTemp(4)
	sTitle  = arrTemp(5)
	sVar    = arrTemp(6)
	sAttributeList = arrTemp(7)
	if UBound(arrTemp)>7 then
	    sCallFrom = arrTemp(8)
	    'Response.Write "CallFrom="& sCallFrom
	end if
	if trim(iItem)<>"" then
	    sSQL = "Select ReceiptNumbering from VwItem where ItemCode ="& iItem &" and ClassificationCode = "& iClass
	    'Response.Write sSQL
	    dcrs.open sSQL,con
	    if not dcrs.eof then
	        sRecptNum = trim(dcrs(0))
	    end if
	    dcrs.close
	end if 'if trim(iItem)<>"" then

	sItemCode = Request.QueryString("ItemCode")
	sClassCode = Request.QueryString("ClassCode")
	sCatCode = Request.QueryString("CatCode")
	sItemName = Request.QueryString("ItemName")
	iRowCount = Request.QueryString("hRowCount")
	if Trim(Request.QueryString("hSubmit"))<>"" and IsNumeric(Request.QueryString("hSubmit")) then
	    iCurrentPage = cint(Request.QueryString("hSubmit"))
	end if
	sFromDate = Request.QueryString("FromDate")
	sToDate = Request.QueryString("ToDate")
	if Trim(sClassCode)<>"" then
	    sClassName = getClassName(sClassCode)
	end if
	if Trim(sFromDate)="" then
	    sFinPeriod = Session("FinPeriod")
	    sFinArr = Split(sFinPeriod,":")
	    sFromDate = "01/04/"&sFinArr(0)
	    sToDate = "31/03/"&sFinArr(1)
	end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>iTMS - <%=sTitle%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<script type="application/xml" data-itms-xml-island="1" id="ItemData">
<Root/>
</script>
<script type="application/xml" data-itms-xml-island="1" id="ItemXML"><Root></Root></script>
<%
    if trim(iItem)<>"" then
        sItmName = ItemDisplay(iItem,iClass)
	    arrUoM = split(DisplayUoM(sOrgID,iClass,iItem),":")
	    'Response.Write  sOrgID&iClass
	    sUoMCode = arrUoM(0)
	    sUoMDesc = arrUoM(1)
	end if 'if trim(iItem)<>"" then

	sSQL = "Select IsNull(IssueEntryCode,IssueEntryNo),Convert(varchar,IssueDate,103) from Inv_T_MaterialIssueHeader where IssueEntryNo = "& iIssNo
	dcrs.open sSQL,con
	if not dcrs.eof then
	    sIssueCode= dcrs(0)
	    dIssDate = dcrs(1)
	end if
	dcrs.close
%>
<script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/ValidateFormat.js"></SCRIPT>
<script language="javascript" src="../../scripts/PrintWindow.js"></script>
<script LANGUAGE="javascript" SRC="../../scripts/GetPopUpWindowSize.js"></script>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/matConReceiptSelectionPop.js"></SCRIPT>
</head>
<BODY leftMargin=0 topMargin=0  onLoad="fnInit('<%=trim(Request.QueryString("sTemp"))%>')">

<form method="POST" name="formname">
<input type=hidden name="hCallFrom" value="<%=sCallFrom%>">
<input type="hidden" name="hObjVal" value="<%=Request.QueryString("sTemp")%>" />
<input type="hidden" name="hOrgID" value="<%=sOrgID%>" />
<input type="hidden" name="hClassCode" value="<%=sClassCode%>" />
<input type="hidden" name="hCatCode" value="<%=sCatCode%>" />
<input type="hidden" name="hItemCode" value="<%=sItemCode%>" />
<input type="hidden" name="hFromDate" value="<%=sFromDate%>" />
<input type="hidden" name="hToDate" value="<%=sToDate%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class="PageTitle">Receipt Details
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table border="0" cellpadding="0" cellspacing="0">
                                        <%if trim(sCallFrom)="ITEM" then %>
                                        <tr>
                                            <td class="FieldCell">Item Description</td>
                                            <td class="FieldCellSub">
                                                <span class="DataOnly" id="idItemName"><%=sItmName%>&nbsp;</span>
                                            </td>
                                        </tr>
                                        <%end if 'if trim(sCallFrom)<>"" then %>
                                        <tr>
                                            <td class="FieldCell">Issue No - Date</td>
                                            <td class="FieldCellSub">
												<span class="DataOnly"><%=sIssueCode%>-<%=dIssDate%>&nbsp;</span>
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
                                <td>
                                    <table class="BodyTable" width="100%">
                                        <tr>
                                            <td class="FieldCellSub">Date From</td>
                                            <td class="FieldCellSub">
                                                <% Response.Write InsertDatePicker("ctlFrom")%>
                                            </td>
                                            <td class="FieldCellSub">To</td>
                                            <td class="FieldCellSub">
                                                <% Response.Write InsertDatePicker("ctlTo")%>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="FieldCellSub">Classification</td>
                                            <td class="FieldCellSub"><span class="DataOnly" id="txtClass"><%=sClassName%>&nbsp;</span>&nbsp;&nbsp;<img style="cursor: hand" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="top" width="11" height="11" alt="Select Classification" onclick="SelectClassifcation()"></td>
                                             <td class="FieldCellSub">Item</td>
                                            <td class="FieldCellSub"><span class="DataOnly" id="txtItem"><%=sItemName%>&nbsp;</span>&nbsp;&nbsp;<img style="cursor: hand" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="top" width="11" height="11" alt="Select Classification" onclick="SelectItem()">&nbsp;&nbsp;<input type="button" name="btnGo" value=" GO " class="ActionButtonX" onclick="NextSelection('1')" /></td>
                                        </tr>
                                    </table>
                                </td>
                                <td align="center"></td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
                            </tr>
							<tr>
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<DIV class=frmBody id=frm6 style="width: 100%; height:380;">
										<table border="0" cellspacing="1" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
												<td class="ExcelHeaderCell" align="center" width="10"></td>
												<td class="ExcelHeaderCell" align="center" width="80">Receipt No-<br />Date</td>
												<td class="ExcelHeaderCell" align="center">Item</td>
												<td class="ExcelHeaderCell" align="center">Qty[Packs]</td>
												<td class="ExcelHeaderCell" align="center">Output</td>
												<td class="ExcelHeaderCell" align="center">By Product</td>
											</tr>
											<%
											    iCtr = 0
											    sSQL = "Select H.InternalReceiptNo,Convert(varchar,ReceivedOn,103),D.ItemCode,ItemDescription,SUM(QuantityReturn),IsNull(InvRecNo,0),IsNull(D.AttributeList,'') "
											    sSQL = sSQL & "from APP_T_InternalReceiptHeader H Join APP_T_InternalReceiptDetails D on "
											    sSQL = sSQL & "H.InternalReceiptNo = D.InternalReceiptNo Join Inv_M_ItemMaster I on D.ItemCode = I.ItemCode and (InvRecNo is not null or InvRecNo <>0) "
											    sSQL = sSQL & " and ReceivedOn between Convert(datetime,'"& sFromDate&"',103) and Convert(datetime,'"& sToDate &"',103)"

											    if Trim(sClassCode)<>"" then
											        sSQL = sSQL & " and D.ClassificationCode in ("& sClassCode  &")"
											    end if

											    if Trim(sItemCode)<>"" then
											        sSQL = sSQL & " and D.ItemCode in ("& sItemCode  &")"
											    end if
											    sSQL = sSQL & " Group By H.InternalReceiptNo,ReceivedOn,D.ItemCode,ItemDescription,InvRecNo,D.AttributeList"
											    'Response.write "<textarea>"& sSql &"</textarea>"
											    With dcrs
											        .activeConnection = con
											        .CursorLocation = 3
											        .CursorType = 3
											        .Source = sSQL
											        .open
											    End With
											    if not dcrs.eof then
											        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
   														dcrs.PageSize = iPageSize
														If iCurrentPage = 0 then iCurrentPage = 1	'initially make current page first page
														dcrs.AbsolutePage = iCurrentPage			'specifies that current = record resides in CPage
														iTotPage = dcrs.PageCount					'stores total no. of pages
													'''''''''''''''''''''''''''''''''''''''''''''''''''''''
														For iPageCtr = 1 to dcrs.PageSize

											                iNoOfPacks = 0
											                iIntRcptNo = dcrs(0)
											                iIntItemCode = dcrs(2)
											                sIntAttList = dcrs(6)
											                iRcptQty = dcrs(4)
											                sSQL = "Select Count(*) from INV_T_LocationLot where InventoryReceiptNo = "& dcrs(5) &" and ItemCode ="& dcrs(2)
											                sSql = sSql & " and SerialNumber not in (Select Serialno from INV_T_MaterialConsumptionOutput H join INV_T_MaterialConsumptionOutputDet D on H.ConsumptionNo = D.ConsumptionNo  and H.IssueEntryNo = D.IssueEntryNo and H.LineNumber = D.LineNumber)"
											                'Response.write "<textarea>"& sSql &"</textarea>"
											                dcrs1.open sSQL,con
											                if not dcrs1.eof then
											                    iNoOfPacks = dcrs1(0)
											                end if
											                dcrs1.close

											                sSql = "Select IsNull(Sum(OutputQuantity),0) from INV_T_MaterialConsumptionOutput where AppRefNo = "& iIntRcptNo &" and AppRefType = 39"
											                dcrs1.open sSql,con
											                if not dcrs1.eof then
											                    iOutPutQty = dcrs1(0)
											                end if
											                dcrs1.close

											               sRcptNum = GetItemRcptNum(iIntItemCode)

								                            if iRcptQty<>"" then
								                                if cdbl(iRcptQty)>cdbl(iOutPutQty) then
								                                    iRcptQty =  cdbl(iRcptQty)-cdbl(iOutputQty)
								                                else
								                                    iRcptQty = 0
								                                end if
								                            end if


											                if iRcptQty>0 then
											                iCtr = iCtr + 1
											                iRowCount = iRowCount + 1
											                %>
											                    <tr>
												                    <td class="ExcelSerial" align="center"><%=iCtr%></td>
												                    <td class="ExcelDisplayCell" align="center" width="10">
												                        <input type="checkbox" name="ChkZ<%=iCtr%>" value="<%=iIntRcptNo%>:<%=iIntItemCode%>:<%=sIntAttList%>:<%=sRcptNum%>" onclick="GenXML('<%=iCtr%>')"/>
												                    </td>
												                    <td class="ExcelDisplayCell" align="center" width="80"><%=dcrs(0)%>-<%=dcrs(1)%></td>
												                    <td class="ExcelDisplayCell" align="left"><%=dcrs(3)%>
												                    </td>
												                    <td class="ExcelDisplayCell" align="center"><%=iRcptQty%>
												                    <%if iNoOfPacks>0 then
												                        Response.write "["&iNoofPacks&"]"
												                      end if'if iNoOfPacks>0 then %>
												                    </td>
												                    <td class="ExcelDisplayCell" align="center">
												                        <%if trim(sRcptNum)="N" then%>
												                        <input type="text" name="txtOPQtyZ<%=iIntRcptNo%>Z<%=iIntItemCode%>Z<%=sIntAttList%>" value="0" class="FormElem" style="text-align:right" size="12" >
												                        <%else%>
												                        <input type="text" name="txtOPQtyZ<%=iIntRcptNo%>Z<%=iIntItemCode%>Z<%=sIntAttList%>" value="0" class="FormElem" style="text-align:right" size="12" disabled="true">
												                        <img src="../../assets/images/iTMS%20icons/Entryicon.gif" alt="Pick Details" style="cursor:hand" id="imgZ<%=iIntRcptNo%>Z<%=iIntItemCode%>Z<%=sIntAttList%>" onClick="PackDisplay('<%=iIntRcptNo%>','<%=iIntItemCode%>','<%=sIntAttList%>','<%=dcrs(4)%>','<%=sRcptNum%>','<%=iCtr%>')" >
												                        <%end if 'if trim(sRcptNum)="N" then%>
												                    </td>
												                    <td class="ExcelDisplayCell" align="center">
												                        <input type="checkbox" name="ChkBPZ<%=iCtr%>" value="0" class="FormElem" style="text-align:right">
												                    </td>
											                    </tr>
											                <%
											                end if 'if iRcptQty>0 then
											            dcrs.movenext
											            If dcrs.EOF Then Exit For
											        Next
											    end if
											    dcrs.close
											%>
										</table>
									</div>
								</td>
								<input type="hidden" name="hCtr" value="<%=iCtr%>">
								<input type="hidden" name="hRowCount" value="<%=iRowCount%>" />
								<td align="center" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" colspan="3" class="MiddlePack"></td>
							</tr>
							<tr>
									    <td colspan="2" align="right">
								                        <Input Type=Hidden name="hCurrentPage" Value="<%=iCurrentPage%>" >
								                        <Input Type=Hidden name="hPageSelection" Value="1" >
														<%	If iTotPage >= 2 Then
																if iCurrentPage = 1 then
														%>
														<input type="button" value=" |< " class="ActionButtonX" id=button1 name=button1>
														<input type="button" value=" << " class="ActionButtonX" id=button2 name=button2>
														<%		else	%>
														<input type="button" value=" |< " class="ActionButtonX" onclick="NextSelection('1')" id=button3 name=button3>
														<input type="button" value=" << " class="ActionButtonX" onclick="NextSelection('<%=iCurrentPage - 1%>')" id=button4 name=button4>
    													<%		end if	%>
    													<SELECT class="FormElem" onChange="NextSelection(this.options[this.selectedIndex].value)" id="selPage">
    													<%
															For lnPage = 1 To iTotPage
																If lnPage = iCurrentPage Then
														%>
															<OPTION value="<%=lnPage%>" selected>Page <%=lnPage%> of <%=iTotPage%></OPTION>
														<%		else	%>
															<OPTION value="<%=lnPage%>">Page <%=lnPage%></OPTION>
    													<%		end if
    														next
    													%>
    													</SELECT>
    													<%
    															if iCurrentPage = iTotPage then
    													%>
														<input type="button" value=" >> " class="ActionButtonX" id=button5 name=button5>
														<input type="button" value=" >| " class="ActionButtonX" id=button6 name=button6>

    													<%		else	%>
														<input type="button" value=" >> " class="ActionButtonX" onclick="NextSelection('<%=iCurrentPage + 1%>')" id=button7 name=button7>
														<input type="button" value=" >| " class="ActionButtonX" onclick="NextSelection('<%=iTotPage%>')" id=button8 name=button8>
    													<%		end if
															End If
														%>
												</td>
												<td></td>

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
                                                    <input type="button" value="Done" name="B1" class="ActionButton" onClick="CheckSubmit()">
                                                    <input type="button" value="Close" name="B2" class="ActionButton" onclick="Func_Close()">
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
		'Response.Write dcrs.source
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
	' Function to populate the Account Head list
	Function populateAccountHead()
		' Declaration of variables
		Dim dcrs,stypID,stypName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT DISTINCT ACCOUNTHEAD,ACCOUNTHEADCODE FROM VWORGGLHEADS WHERE OUDEFINITIONID = " & Pack(sOrgID) & " AND ACCOUNTHEAD IN (SELECT ACCOUNTHEAD FROM ACC_R_GLACCAPPLICATIONS WHERE AVAILABLEINAPPLN IN (4,5,6) AND OUDEFINITIONID = " & Pack(sOrgID) & ") ORDER BY 2"
			.ActiveConnection = con
			.Open
		end with
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				stypID = dcrs(0)
				stypName = dcrs(1)
				if cint(iAccHead) = cint(stypID) then
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
