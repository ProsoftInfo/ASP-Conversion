<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParDisplayGrid.asp
	'Module Name				:	Accounts
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 09,2010
	'Modified By				:
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
<%
Dim rsObj
Dim nSlNo
Dim iPartyCode,iStartRec,iEndRec
Dim iPrevPage,iTotalPages,nPageCtr,iNextPage,iPageSize,iPageNo
Dim iTotalRecords
Dim sQuery,sParName,sCity,sTINNumber,sSearch,sTabValue

set rsObj = Server.CreateObject("ADODB.Recordset")

sParName = Request("hParName")
sCity = Request("hCity")
sTINNumber = Request("hTINNumber")
sSearch = Request("hSearch")
sTabValue = Request("TAB")

iPageSize = 15

iPageNo = trim(Request("hPage"))
if trim(iPageNo) = "" then iPageNo = 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS - Accounts</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DivClick.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<Script Language="VBScript">
'**********************************************
Function ViewContactDeatils(sPartyCode)
	showModalDialog "ParDisplayContactDetails.asp?PartyCode="&sPartyCode,"","dialogHeight:350px;Status:no"
End Function
'************************************************
Function CreateNewParty()
	document.formname.action = "ParCreate_Edit_Entry.asp"
	document.formname.submit
End Function
'**********************************
Function AssignPage(nPage)
	document.formname.hPage.value = nPage
	document.formname.submit()
End Function
'************************************************
Function CheckSubmit()
	document.formname.hParName.value = document.formname.txtPartyName.value
	document.formname.hCity.value = document.formname.txtCity.value
	document.formname.hTINNumber.value = document.formname.txtTINNumber.value
	document.formname.hSearch.value = document.formname.selParSearchType(document.formname.selParSearchType.selectedIndex).value
	document.formname.submit
End Function
'************************************************
Function EditParty(sPartyCode,sTabValue)
    if sTabValue = "C" then
	    document.formname.action = "PartyControlData.asp?PartyCode="&sPartyCode
	else
	    document.formname.action = "ParCreate_Edit_Entry.asp?PartyCode="&sPartyCode
	end if
	document.formname.submit
End Function
'************************************************
Function DelParty()
	Dim sRow,iCtr,sPartyVal,objHttp,Returnvalue,sUnDeleteParty,iDelItemCount,iSelItemCount,iCountUnit,iUnitCode
	set objhttp = CreateObject("Microsoft.XMLHTTP")

	iDelItemCount = 0
	iSelItemCount = 0


	For iCtr=1 to cint(document.formname.hCnt.value)
	    if cint(document.formname.hCnt.value)=1 then
	        if eval("document.formname.radButton").checked = true then
			    iSelItemCount = iSelItemCount + 1
			    sPartyVal = eval("document.formname.radButton").value
			    document.formname.hPartyCode.value = "0?0?Selected Party?"& sPartyVal

			    Objhttp.Open "POST","PartyDelCheck.asp?sCallType=P?"&document.formname.hPartyCode.value , false
			    Objhttp.send

			    IF Objhttp.responsetext = "T" Then
				    sUnDeleteParty = sUnDeleteParty & "," & eval("document.formname.hPartyName"&iCtr).value
			    else

				    objHttp.Open "GET","GetUnitsForParty.asp?ParCode="& sPartyVal,false
				    objHttp.send
				    'alert(objHttp.responseText)
				    if trim(objHttp.responseText)<>"" then
					    'alert(objHttp.responseText)
					    iCountUnit = split(objHttp.responseText,":")(0)
					    iUnitCode = Split(objHttp.responseText,":")(1)
					    'alert(iCountUnit)
					    if trim(iCountUnit)<>"" then
						    if iCountUnit = 1 then
							    Returnvalue = iUnitCode
						    else
							    Returnvalue = showModalDialog ("ParDelSelectType.asp?Type="&Ftype ,"","dialogHeight:250px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
							    if trim(Returnvalue)="" then
								    exit function
							    end if
						    end if ' if Returnvalue > 0 then

						    Objhttp.Open "POST","ParDeleteEntry.asp?hPartyCode="&document.formname.hPartyCode.value&"&hCallTy=D&hDelTy="& Returnvalue , false
						    Objhttp.send
						    if objHttp.responsetext<>"" then
							    alert(objHttp.responseText)
						    else
							    alert("Party Deleted Successfully")
							    document.formname.submit
						    end if
					    end if 'if trim(iCountUnit)<>"" then
				    end if ' if trim(objHttp.responseText)<>"" then
			    End if
		    end if ' if eval("document.formname.CheckBox"&iCtr).checked = true then
	    else
	        if eval("document.formname.radButton")(iCtr-1).checked = true then
			    iSelItemCount = iSelItemCount + 1
			    sPartyVal = eval("document.formname.radButton")(iCtr-1).value
			    document.formname.hPartyCode.value = "0?0?Selected Party?"& sPartyVal

			    Objhttp.Open "POST","PartyDelCheck.asp?sCallType=P?"&document.formname.hPartyCode.value , false
			    Objhttp.send

			    IF Objhttp.responsetext = "T" Then
				    sUnDeleteParty = sUnDeleteParty & "," & eval("document.formname.hPartyName"&iCtr).value
			    else

				    objHttp.Open "GET","GetUnitsForParty.asp?ParCode="& sPartyVal,false
				    objHttp.send
				    'alert(objHttp.responseText)
				    if trim(objHttp.responseText)<>"" then
					    'alert(objHttp.responseText)
					    iCountUnit = split(objHttp.responseText,":")(0)
					    iUnitCode = Split(objHttp.responseText,":")(1)
					    'alert(iCountUnit)
					    if trim(iCountUnit)<>"" then
						    if iCountUnit = 1 then
							    Returnvalue = iUnitCode
						    else
							    Returnvalue = showModalDialog ("ParDelSelectType.asp?Type="&Ftype ,"","dialogHeight:250px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No")
							    if trim(Returnvalue)="" then
								    exit function
							    end if
						    end if ' if Returnvalue > 0 then

						    Objhttp.Open "POST","ParDeleteEntry.asp?hPartyCode="&document.formname.hPartyCode.value&"&hCallTy=D&hDelTy="& Returnvalue , false
						    Objhttp.send
						    if objHttp.responsetext<>"" then
							    alert(objHttp.responseText)
						    else
							    alert("Party Deleted Successfully")
							    document.formname.submit
						    end if
					    end if 'if trim(iCountUnit)<>"" then
				    end if ' if trim(objHttp.responseText)<>"" then
			    End if
		    end if ' if eval("document.formname.CheckBox"&iCtr).checked = true then
		end if
	Next

	if iSelItemCount=0 then
		alert("Select the Party to Delete")
		exit function
	end if ' if iSelItemCount=0 then

	if trim(sUnDeleteParty)<>"" then
		sUnDeleteParty = mid(sUnDeleteParty,2)
		MsgBox( sUnDeleteParty &" is having Transactions Could not be Deleted " )
	end if
End Function
'************************************************
</Script>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" action="">
	<input type="hidden" name="hPage" value="<%=iPageNo%>">
	<input type="hidden" name="hParName" value="<%=sParName%>">
	<input type="hidden" name="hCity" value="<%=sCity%>">
	<input type="hidden" name="hTINNumber" value="<%=sTINNumber%>">
	<input type="hidden" name="hPartyCode" value="">
	<input type="hidden" name="hSearch" value="<%=sSearch%>">
	<input type="hidden" name="hDelTy" value="">
	<input type="hidden" name="hCallTy" value="">

	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Party Master
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack" height="7px">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
								</tr>

								<tr>
									<td align="center" width="5px" class="ClearPixel">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
									<td valign="top" width="100%">
										<table border="0" cellpadding="0" cellspacing="0" width="100%" class="ExcelTable">
											<tr>
												<td>
													<div>
														<table class="CollapseBand" cellspacing="0" cellpadding="0" class="BodyTable">
															<tr>
																<td valign="center"><a style="width: 1em; height: 1em;" title href onclick="Div_OnClick(idUnprocessed,'')" >
																	<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" border="0" src="../../assets/images/plus.gif" width="10px" height="10px" alt="Expands this section for more search criteria.">
																	</a>
																</td>
																<td valign="center" class="SubTitle">&nbsp;&nbsp;

																</td>
															</tr>

														</table>
														<table border="0" cellpadding="0" cellspacing="0" width="100%">
															<tr>
																<td width="100%">
																	<div id="idUnprocessed" style="display: none">
																		<table cellpadding="0" cellspacing="0" class="BodyTable" width="100%">
																			<tr>
																				<td class="MiddlePack">
																				</td>
																				<td class="MiddlePack" colspan="4">
																				</td>
																			</tr>
																			<tr>
																				<td width="100px"></td>
																				<td class="FieldCell">Party Name</td>
																				<td class="FieldCellSub" colspan=3>
																					<Select name="selParSearchType" class="FormElem">
																						<Option value="SB" <%if sSearch="SB" then Response.Write "Selected"%> >Start With</Option>
																						<Option value="WA" <%if sSearch="WA" then Response.Write "Selected"%> >Any Where</Option>
																					</Select>
																					<input type="text" name="txtPartyName" value="<%=sParName%>" class="FormElem">
																				</td>
																			</tr>
																			<tr>
																				<td width="100px"></td>
																				<td class="FieldCell">City</td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtCity" value="<%=sCity%>" class="FormElem">
																				</td>
																				<td class="FieldCell">TIN Number</td>
																				<td class="FieldCellSub">
																					<input type="text" name="txtTINNumber" value="<%=sTINNumber%>" class="FormElem">
																					<input type="button" name="btnGo" value="  GO  " class="ActionButtonX" onClick="CheckSubmit()">
																				</td>
																			</tr>
																			<tr>
																				<td class="Middlepack" colspan="5">

																				</td>
																			</tr>

																		</table>
																	</div>
																</td>
															</tr>

														</table>
													</div>
												</td>
											</tr>

										</table>
									</td>
									<td align="center" class="ClearPixel" width="5px">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
								</tr>

								<tr>
									<td align="center" class="MiddlePack" colspan="3">
									</td>
								</tr>

								<tr>
									<td align="center" width="5px" class="ClearPixel">
									</td>
									<td valign="top">
										<!--div class="frmBody" id="frm4" style="width: 585; height:140;"-->
										<table border="0" cellspacing="1px" class="ExcelTable" width="100%">
											<tr>
												<td class="ExcelHeaderCell" width="10px" >S.No.
												</td>
												<td class="ExcelHeaderCell"><a style="width: 1em; height: 1em;" title href onclick="DelParty()" itms_state="0">
													<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" alt="Click here to delete the Party" width="15px" height="15px">
													</a>
												</td>
												<td class="ExcelHeaderCell">Party Code
												</td>
												<td class="ExcelHeaderCell">Party Name
												</td>
												<td class="ExcelHeaderCell">City
												</td>
												<td class="ExcelHeaderCell">Phone No
												</td>
												<td class="ExcelHeaderCell">TIN Number
												</td>
												<td class="ExcelHeaderCell">Status
												</td>
											</tr>

											<%
												sQuery = "Select PartyCode,PartyName,City,isNull(TINNumber,'') TINNumber,OrgnPartyCode,isNull(PhoneNos,'') from APP_M_PartyMaster where isNull(PartyName,'')<>'' "

												if trim(sParName)<>"" then
													if sSearch = "SB" then
														sQuery = sQuery & " and PartyName like '"& sParName &"%' "
													else
														sQuery = sQuery & " and PartyName like '%"& sParName &"%' "
													end if
												end if

												if trim(sCity)<>"" then
													sQuery = sQuery & " and City like '"& sCity & "%'"
												end if

												if trim(sTINNumber)<>"" then
													sQuery = sQuery & " and TINNumber like '"& sTINNumber &"%'"
												end if

												'Response.Write "<p>" & sQuery

												with rsObj
													.CursorLocation = 3
													.CursorType = 3
													.ActiveConnection = con
													.PageSize = iPageSize
													.Source = sQuery
													.Open
												end with

												nSlNo = 1
												If not rsObj.EOF Then
													iTotalPages = rsObj.PageCount
													iTotalRecords = rsObj.RecordCount
													rsObj.AbsolutePage = iPageNo
												Else
													iTotalPages = 0
													iTotalRecords = 0

													iStartRec = 0
													iEndRec = 0
												End If

												if trim(iPageNo) = 1 then
													iPrevPage = 0
												else
													iPrevPage = iPageNo - 1
												end if


												if iTotalPages >= iPageNo + 1 then
													iNextPage = iPageNo + 1
												else
													iNextPage = 0
												end if

												do while not rsObj.EOF and nSlNo <= iPageSize
													iPartyCode = rsObj(0)
											%>

											<tr>
												<td class="ExcelSerial"><%=nSlNo%>
												</td>
												<td class="ExcelDisplayCell" width="10px">
													<input type="radio" name="radButton" value="<%=iPartyCode%>">
												</td>
												<td class="ExcelDisplayCell" align="left"><%=rsObj(4)%></td>
												<td class="ExcelDisplayCell" align="left">
												<a href="#" class="ExcelDisplayLink" onClick="EditParty('<%=iPartyCode%>','<%=sTabValue%>')"><%=rsObj(1)%></a>
													<input type=hidden name="hPartyName<%=nSlNo%>" value="<%=rsObj(1)%>">
												</td>
												<td class="ExcelDisplayCell" align="left"><a href="#" onClick="ViewContactDeatils('<%=iPartyCode%>')" class="ExcelDisplayLink"><%=rsObj(2)%></a>
													<!--<img border="0" src="../../assets/images/iTMS%20icons/Details.gif" width=15 height=15 alt="View the Contact Details" onClick="ViewContactDeatils('<%=iPartyCode%>')">-->
												</td>
												<td class="ExcelDisplayCell" align="left"><%=rsObj(5)%></td>
												<td class="ExcelDisplayCell" align="left"><%=rsObj(3)%></td>
												<td class="ExcelDisplayCell" align="left"></td>
											</tr>
											<%
													nSlNo=nSlNo + 1
													rsObj.MoveNext
												loop
											%>
											<input type="hidden" name="hCnt" value="<%=nSlNo-1%>" >
										</table>
										<!--/div-->
									</td>
									<td align="center" class="ClearPixel" width="5px">
									</td>
								</tr>

								<tr>
									<td align="center" class="MiddlePack" colspan="3">
									</td>
								</tr>

								<tr>
									<td align="center" width="5px" class="ClearPixel">
									</td>
									<td valign="top" align="right">

									<input type="button" value=" |< " class="ActionButtonX" id=ButFirst name=ButFirst onClick="AssignPage('1')">

									<%if trim(iPrevPage) = "0" then  %>
										<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev >
									<%else%>
										<input type="button" value=" << " class="ActionButtonX" id=ButPrev name=ButPrev onClick="AssignPage('<%=iPrevPage%>')">
									<%end if %>


									<SELECT class="FormElem" onChange="AssignPage(this.value)"  id="mCmbPage" name="mCmbPage">

									<%for nPageCtr= 1 to iTotalPages %>
										<option value="<%=nPageCtr%>" <%if trim(iPageNo) = trim(nPageCtr) then Response.Write "Selected" %> >Page <%=nPageCtr%> of <%=iTotalPages %></option>
									<%next%>

									</SELECT>
									<%if trim(iNextPage) = "0" then  %>
										<input type="button" value=" >> " class="ActionButtonX" id=ButNext name=ButNext >
									<%else%>
										<input type="button" value=" >> " class="ActionButtonX" onclick="AssignPage('<%=iNextPage%>')" id=ButNext name=ButNext >
									<%end if%>

									<input type="button" value=" >| " class="ActionButtonX" id=ButLast name=ButLast OnClick="AssignPage('<%=iTotalPages %>')">

									</td>
									<td align="center" class="ClearPixel" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>

								<tr>
									<td align="center" class="MiddlePack" colspan="3">
									</td>
								</tr>

								<tr>
									<td align="center" width="5px" class="ClearPixel">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td class="ActionCell">
													<input type="button" value="Create New" name="BtnNewParty" class="ActionButtonX" onclick="CreateNewParty()">
												</td>
											</tr>

										</table>
									</td>
									<td align="center" class="ClearPixel" width="5px">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
									</td>
								</tr>

								<tr>
									<td align="center" class="BottomPack" colspan="3">
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
