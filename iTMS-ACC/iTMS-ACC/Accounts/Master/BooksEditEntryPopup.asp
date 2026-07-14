<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	BooksEditEntry.asp
	'Module Name				:	ACCOUNTS (Master Amendment)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 10,2010
	'Modified On				:	Dec 08,2010
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
dim objRs,objRs1,objFs,sQuery,sSelBookName,sAmnType
Dim oDOM,Root,newElem,newElem1,nodUnit
dim sUnitLName,sUnitSName,sBookID,sBookName
Dim sBookCode,sBookNo,iUnitno,sSelBookVal,Temparr,sSelDayBook
Dim sSelBookCode,sSelBookNo
Dim sMon,sYear,sMonYr,sFinYear,sFinFrom,sFinTo,saTemp,sCrValue,sDrValue
Dim sLogFinPer,sTemp
Dim bUseable
Dim iDrSeriesNo,iDrSeriesCode,iCrSeriesNo,iCrSeriesCode,sSeriesName,sCounterType
dim sAccName,iToAccCode,sToAccName,iSno,objRs2
Dim sFromAccNo,sFromAccName
Dim iBookCode,iBookNumber,iRecordsCount

sLogFinPer = Session("FinPeriod")
sTemp = Split(sLogFinPer,":")

sYear = sTemp(0)

sMon = Month(Date)
'sYear = Year(Date)



IF CInt(sMon) <=9 Then
	sMon = 0&sMon
End IF
sMonYr = sMon&sYear
sFinYear = GetFinancialYear(sMonYr)
saTemp = Split(sFinYear,":")
sYear = Right(saTemp(0),4)
sMon = Mid(saTemp(0),4,2)
sFinFrom = sYear&sMon

sYear = Right(saTemp(1),4)
sMon = Mid(saTemp(0),4,2)
sFinTo = sYear&sMon

sSelBookName = ""
iUnitno = Request.QueryString("OrgCode")
sSelDayBook = Request.QueryString("BookCode")
sSelBookVal = Request.QueryString("BookNumber")
IF CStr(iUnitno) = "" Then
	iUnitno = 0
End IF

IF CStr(sSelDayBook) = "" Then
	sSelDayBook = "0"
End IF

IF CStr(sSelBookNo) = "" Then
	sSelBookNo = 0
End IF

'Response.Write sSelDayBook

sSelDayBook = Trim(sSelDayBook)

sFinFrom = Trim(sTemp(0))&"04"
sFinTo = Trim(sTemp(1))&"04"


Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs1 = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")
Set objfs = CreateObject("Scripting.FileSystemObject")

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
sFromAccNo=Request.QueryString("FromAcc")

if trim(sFromAccNo)<>"" then
	sQuery = "Select b.AccountDescription from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
			  "where a.OUDefinitionID='"& iUnitno &"' and b.AccountHead=a.FromAccountHead and "&_
			  " a.FromAccountHead = "& sFromAccNo &" Group by a.FromAccountHead,b.AccountDescription"
			 ' Response.Write sQuery
	objRs.Open sQuery,con
	if not objRs.EOF then
		sFromAccName = trim(objRs(0))
	end if
	objRs.Close
end if ' if trim(sFromAccNo)<>"" then
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<base target="_self">
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<script type="application/xml" data-itms-xml-island="1" ID="SeriesNoData"><Root /></script>
<script type="application/xml" data-itms-xml-island="1" ID="AccData"><Root /></script>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<SCRIPT SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<script>
window.__itmsPopupCompat = { type: "booksEditEntryPopup" };
</script>
<script src="/Scripts/itms-modern-compat.js"></script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="CreateXML();popSeriesNo()">

<form method="POST" name="formname">
<input type=hidden name="hSeriesType" value="">
<input type=hidden name="hSeriesLen" value="">
<input type=hidden name="hOrgCode" value="<%=iUnitno%>">
<input type=hidden name="hBookCode" value="<%=sSelDayBook%>">
<input type=hidden name="hBookNo" value="<%=sSelBookVal%>">
<input type=hidden name="hCallType" value="">
<input type=hidden name="hAction" value="">
<input type=hidden name="hFromHead" value="<%=sFromAccNo%>">

<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<%
			With objRs
				.CursorLocation = 3
				.CursorType = 3
				.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 and OUDEFINITIONID ='"&iUnitno&"'"
				.ActiveConnection = con
				.Open
			End With
			if not objRs.EOF then
				Response.Write objRs(1)
			end if
			objrs.Close
		%><br>
		Books Edit Entry</p>
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
						<table border="0" cellpadding="0" cellspacing="0" width="100%" >
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
							<td colspan=3>
								<table class=ExcelTable cellpadding="0" cellspacing="0" width="100%" class="bodytable">
								<tr>
								<td colspan=3>
									<div>
										<table class="CollapseBand" cellspacing="0" cellpadding="0">
											<tr>
												<td valign="center"><a style="width: 1em; height: 1em;" title href="#" onclick="DivClick('DivBasic'); return false;" >

													<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" id="imgBasic" border="0" src="../../assets/images/minus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
													</a>
												</td>
												<td valign="center" class="SubTitle">&nbsp;&nbsp;
													<b>Basic Details</b>
												</td>
											</tr>

										</table>
									</div>
								</tr>
								<tr>
								<td>
								<div id=DivBasic class="frmBody" >
								<table  cellspacing=0 cellpadding=0 width=100% class="bodytable">
								<tr>
									<td align="center">
									</td>
									<td valign="top" width="100%">
									<center>
										<table cellpadding="0" cellspacing="0" width="100%" >
											<tr>
												<td width=5>
												<td>
												<table cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td class='GroupTitleLeft' width="10">&nbsp;
								                        </td>
														<td class='GroupTitle' width="60"><p align="center">Basic
								                        </td>
													</center>
														<td class='GroupTitleRight'><p align="left">&nbsp;
								                        </td>
													</tr>
												</table>
								                    </td>
								                    <td width=5>
													</tr>
													<tr>
														<td width=5>
														<td class=GroupTable>
												<center>
								                 <div align="left">
													<table cellpadding="0" cellspacing="0" width="100%" >
													<tr>
														<td class=FieldCell width="200">Day Book Type</td>
														<td class=FieldCell ><span class="DataOnly">
															<%
																sQuery = "Select BookCode,BookName from ACC_M_DayBooks"
																objRs.Open sQuery,con
																if not objRs.EOF then
																	do while not objRs.EOF
																		if cstr(objRs(0))=CStr(sSelDayBook) then
																			Response.Write objRs(1)
																		end if
																		objRs.MoveNext
																	loop
																end if
																objRs.Close

																sQuery = "Select Useable From Acc_R_ApplicableAccountHeads "&_
																		 "Where OUDefinitionID = '"& iUnitno &"' and BookCode = '"& sSelDayBook &"' and BookNumber ="& sSelBookVal &" "
																		 'Response.Write sQuery
																objRs.Open sQuery,con
																if not objRs.EOF then
																	bUseable = objRs(0)
																end if
																objrs.Close
															%></span> &nbsp;&nbsp;&nbsp; Useable
															<%
															if trim(bUseable)="0" then%>
																<input type=radio name=optUseable value="0" class=FormElem onClick=FunUseable() Checked  >Yes
																<input type=radio name=optUseable value="0" class=FormElem onClick=FunUseable() >No
															<%else%>
																<input type=radio name=optUseable value="0" class=FormElem onClick=FunUseable() >Yes
																<input type=radio name=optUseable value="0" class=FormElem onClick=FunUseable() Checked >No
															<%end if%>
										                </td>
													</tr>


													<tr>
														<td class=FieldCell width="200">Book Name</td>
														<td><span class="DataOnly">
														<%
															sQuery = "Select BookCode,BookNumber,BookName From Acc_R_ApplicableAccountHeads "
															sQuery = sQuery &"Where OUDefinitionID = '"&iUnitno&"'  "
															sQuery = sQuery &"and BookCode = '"&sSelDayBook&"' and BookNumber = "& sSelBookVal & " Order By BookName "
															'Response.Write sQuery


															With objRs
																.CursorLocation = 3
																.CursorType = 3
																.ActiveConnection = Con
																.Source = sQuery
																.Open
															End With

															Set objRs.ActiveConnection = Nothing
															IF Not objRs.EOF Then
																sSelBookCode = objRs(0)
																sSelBookNo = objRs(1)
																sSelBookName = objRs(2)

																Response.Write objRs(2)
															End IF
															objRs.Close
														%></span>
										                </td>
													</tr>
													<%
														Dim sCssTab,sCssFrm

														IF Len(sAmnType) = 0 Then
															sQuery = "Select isNull(BookCode,0) From Acc_T_CreatedVoucherHeader Where isNull(BookCode,0) = '"&sSelBookCode&"' "&_
																	 "and isNull(BookNumber,0) = "&sSelBookNo&" and  OUDefinitionID = '"&iUnitno&"' "

															objRs.Open sQuery,Con
															IF Not objRs.EOF Then
																sAmnType = "readonly"
																sCssTab = "ExcelDisplayCell"
																sCssFrm = "FormElemRead"
															Else
																sAmnType = ""
																sCssTab = "ExcelInputCell"
																sCssFrm = "FormElem"
															End IF
															objRs.Close

															sAmnType = ""
															sCssTab = "ExcelInputCell"
															sCssFrm = "FormElem"

														End IF


													%>
										            <tr>
										            	<td align="center" class="MiddlePack" colspan="2">
										            	</td>
										            </tr>

										            <tr>
										            	<td align="center" class="MiddlePack" colspan="2">
										            	</td>
										            </tr>
										            <%
														Dim sOtherUnit,sAccountHead,sRecType
														sQuery="select BookAccountHead,OtherUnitTransaction,Useable from Acc_R_ApplicableAccountHeads "&_
														"where OUDefinitionID='"&iUnitno&"' and BookCode='"&sSelBookCode&"' and BookNumber= "&sSelBookNo

														'Response.Write sQuery

														with objRs
															.CursorLocation = 3
															.CursorType = 3
															.Source = sQuery
															.ActiveConnection = con
															.Open
														end with
														set objRs.ActiveConnection = nothing
														IF Not objRs.EOF Then
															if objRs(1)=0 then
																sOtherUnit="0"
															else
																sOtherUnit="1"
															end if
															sAccountHead=objRs(0)
														End IF

														objRs.Close

														sQuery = "select DrSeriesNo,DrSeriesCode,CrSeriesNo,CrSeriesCode from Acc_M_BookNumberSeries "&_
																 "where OUDefinitionID='"&iUnitno&"' and BookCode='"&sSelBookCode&"' and BookNumber="&sSelBookNo



														'Response.write sQuery

														with objRs
															.CursorLocation = 3
															.CursorType = 3
															.Source = sQuery
															.ActiveConnection = con
															.Open
														end with

														set objRs.ActiveConnection = nothing
														IF Not objRs.EOF Then
															iDrSeriesNo=objRs(0)
															iDrSeriesCode=objRs(1)
															iCrSeriesNo=objRs(2)
															iCrSeriesCode=objRs(3)
														End IF
														objRs.Close



														sRecType = "Y"
														IF CStr(iDrSeriesCode) = CStr(iCrSeriesCode) Then
															IF CStr(iDrSeriesNo) = CStr(iCrSeriesNo) Then
																sRecType = "N"
															End IF
														End IF


										            %>
										            <Input type="hidden" name="hCounterType" value="<%=iDrSeriesNo%>">
													<tr>
														<td class=FieldCell width="200">Allow Other&nbsp;Units Transaction
														</td>
														<td>
															<table border="0" cellpadding="0" cellspacing="0">
																<tr>
																	<%IF CStr(sOtherUnit) = "1" Then %>
																	<td width="20"><input type="radio" value="1" name="optEligible" checked class="formelem"></td>
																	<td class="FieldCell" width="30">Yes </td>
																	<td width="20"><input type="radio" value="0" name="optEligible" class="formelem"></td>
																	<td class="FieldCell">No</td>
																	<%Else%>
																	<td width="20"><input type="radio" value="1" name="optEligible" class="formelem"></td>
																	<td class="FieldCell" width="30">Yes </td>
																	<td width="20"><input type="radio" value="0" name="optEligible" checked class="formelem"></td>
																	<td class="FieldCell">No</td>
																	<%end if %>

																</tr>
															</table>
														</td>
													</tr>
										            <tr>
										            	<td align="center" class="MiddlePack" colspan="2">
										            	</td>
										            </tr>

													<tr>
														<td class=FieldCell width="200"> Book Name</td>
														<td><input type="text" class="Formelem" maxlength="50" name="txtName" size="45" value="<%=sSelBookName%>"></td>
													</tr>
										            <tr>
										            	<td align="center" class="MiddlePack" colspan="2">
										            	</td>
										            </tr>
										            </table>
										        </div>
										     </td>
										     <td width=5>
										   </tr>
										  </table>
									   <center>
										<table cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td width=5>
											<td>
											<table cellpadding="0" cellspacing="0" width="100%">
												<tr>
													<td class='GroupTitleLeft' width="10">&nbsp;
                                                    </td>
													<td class='GroupTitle' width="60"><p align="center">No Series
                                                    </td>
												</center>
													<td class='GroupTitleRight'><p align="left">&nbsp;
                                                    </td>
												</tr>
											</table>
                                                </td>
                                                <td width=5>
												</tr>
												<tr>
													<td width=5>
													<td class=GroupTable>
											<center>
                                             <div align="left">
                                             <table>
															<tr>
																<td class=FieldCell width="200"> Separate Payment / Receipt No&nbsp;</td>
																<td>
																<%IF CStr(sSelDayBook) = "01" or CStr(sSelDayBook) = "02" Then %>
																<select size="1" name="selPayRecNo" class="FormElem" onChange="DisplayBook()"  >
																<%else%>
																<select size="1" name="selPayRecNo" class="FormElem" onChange="DisplayBook()" >
																<%end if %>
																<%IF CStr(sRecType) = "Y" Then %>
																	<OPTION value="Y" Selected>Yes</option>
																	<OPTION value="N">No</option>
																<%else%>
																	<OPTION value="Y">Yes</option>
																	<OPTION value="N" Selected>No</option>
																<%end if %>
															    </select></td>
															</tr>
															<tr>
																<td align="center" class="MiddlePack" colspan="2">
																</td>
																		<!--tr>
																			<td class=FieldCell width="200"> Select
															                  No Series</td>
																			<td>
																			<Input type="text" name="txtNumSerName" class="FormElemRead" readonly size="40">
																			</td>
																		</tr-->
															<tr>
																<td class=FieldCell width="200"> Select
																  No Series</td>
																<td><select size="1" name="selNoSeries" class="FormElem" onChange="DisplayBook()">
																<OPTION value="0">Select Number
																Series</option>
																</select></td>
															</tr>
															<tr>
																<td align="center" colspan="3" class="BottomPack">
																</td>
															</tr>
															<tr>
																<td align="center" valign="top" colspan=2>
															             <table id="tblBook" border="0" cellspacing="1" class="ExcelTable">
															             <%IF CStr(sRecType) = "N" Then %>
															                <tr>
																			<td class="ExcelHeaderCell" align="center" width="10"><p align="center">S.No.</td>
																			<td class="ExcelHeaderCell" align="center" width="75">Period</td>
																			<td class="ExcelHeaderCell" align="center" width="50">Start No</td>
																			<td class="ExcelHeaderCell" align="center" width="100">Prefix</td>
																			<td class="ExcelHeaderCell" align="center" width="100">Suffix</td>
															                </tr>
															             <%Else%>
																			<tr>
																			<td class="ExcelHeaderCell" align="center" width="10"><p align="center">S.No.</td>
																			<td class="ExcelHeaderCell" align="center" width="50">Period</td>
																			<td class="ExcelHeaderCell" align="center" width="25">Cr Start No</td>
																			<td class="ExcelHeaderCell" align="center" width="75">Cr Prefix</td>
																			<td class="ExcelHeaderCell" align="center" width="75">Cr Suffix</td>
																			<td class="ExcelHeaderCell" align="center" width="25">Dr Start No</td>
																			<td class="ExcelHeaderCell" align="center" width="75">Dr Prefix</td>
																			<td class="ExcelHeaderCell" align="center" width="75">Dr Suffix</td>
															                </tr>
															             <%End IF %>

															                <%
																				Dim iCount
																				iCount = 1
																				sQuery = "Select Distinct A.EntryNo,A.Period,A.Number,A.Prefix,A.Suffix From "&_
																						 "APP_R_NoSeriesModuleEntry A, Acc_M_BookNumberSeries B "&_
																						 "Where B.OUDefinitionID = '"&iUnitno&"' and B.BookCode = '"&sSelBookCode&"' "&_
																						 "and B.BookNumber = "&sSelBookNo&" and B.DrSeriesNo = A.SeriesNo "&_
																						 "and B.DrSeriesCode = A.SeriesCode and A.OUDefinitionID = '"&iUnitno&"' "&_
																						 "and Cast(A.Period As Numeric) >= "&sFinFrom&" and  Cast(A.Period As Numeric) <= "&sFinTo&" "

																				'Response.Write sQuery

																				With objRs
																					.CursorLocation = 3
																					.CursorType = 3
																					.ActiveConnection = con
																					.Source = sQuery
																					.Open
																				End With

																				Set objRs.ActiveConnection = Nothing
																				Do While Not objRs.EOF
																					IF CStr(sRecType) = "N" Then
															                %>

																			<tr>
																			<td class="ExcelHeaderCell" align="center" width="10"><p align="center"><%=iCount%></td>
																			<td class="<%=sCssTab%>" align="center" width="75"><%=objRs(1)%></td>
																			<td class="<%=sCssTab%>" align="center" width="50">
																			<input type="text" class="<%=sCssFrm%>" value="<%=objRs(2)%>" name="txtStartNo<%=objRs(0)%>" <%=sAmnType%>>
																			</td>
																			<td class="<%=sCssTab%>" align="center" width="100">
																			<input type="text" class="<%=sCssFrm%>" value="<%=objRs(3)%>" name="txtPrefix<%=objRs(0)%>" <%=sAmnType%> MAXLENGTH="11">
																			</td>
																			<td class="<%=sCssTab%>" align="center" width="100">
																			<input type="text" class="<%=sCssFrm%>" value="<%=objRs(4)%>" name="txtSuffix<%=objRs(0)%>" <%=sAmnType%>>
																			</td>
															                </tr>
															                <%else
																				Dim sCrSuff,sCrPre,sCrStNo
																				sQuery = "Select Suffix,Prefix,Number From APP_R_NoSeriesModuleEntry Where OUDefinitionID = '"&iUnitno&"' "&_
																						 "and SeriesNo = "&iCrSeriesNo&" and SeriesCode = "&iCrSeriesCode&" and Cast(Period As Numeric) >= "&sFinFrom&" "&_
																						 "and  Cast(Period As Numeric) <= "&sFinTo&" and EntryNo = "&objRs(0)&" "


																				With objRs1
																					.CursorLocation = 3
																					.CursorType = 3
																					.ActiveConnection = Con
																					.Source = sQuery
																					.Open
																				End With
																				Set objRs1.ActiveConnection = Nothing
																				IF Not objRs1.EOF Then
																					sCrSuff = objRs1(0)
																					sCrPre = objRs1(1)
																					sCrStNo = objRs1(2)
																				End IF
																				objRs1.Close


															                %>
															                <tr>
																				<td class="ExcelHeaderCell" align="center"><p align="center"><%=iCount%></td>
																				<td class="<%=sCssTab%>" align="center"><%=objRs(1)%></td>
																				<td class="<%=sCssTab%>" align="center">
																					<input type="text" class="<%=sCssFrm%>" value="<%=sCrStNo%>" name="txtCrStartNo<%=objRs(0)%>" size="5" <%=sAmnType%>>
																				</td>
																				<td class="<%=sCssTab%>" align="center">
																					<input type="text" class="<%=sCssFrm%>" value="<%=sCrPre%>" name="txtCrPrefix<%=objRs(0)%>" size="12" <%=sAmnType%> MAXLENGTH="11">
																				</td>
																				<td class="<%=sCssTab%>" align="center">
																					<input type="text" class="<%=sCssFrm%>" value="<%=sCrSuff%>" name="txtCrSuffix<%=objRs(0)%>" size="7" <%=sAmnType%>>
																				</td>
																				<td class="<%=sCssTab%>" align="center">
																					<input type="text" class="<%=sCssFrm%>" value="<%=objRs(2)%>" name="txtDrStartNo<%=objRs(0)%>" size="7" <%=sAmnType%>>
																				</td>
																				<td class="<%=sCssTab%>" align="center">
																					<input type="text" class="<%=sCssFrm%>" value="<%=objRs(3)%>" name="txtDrPrefix<%=objRs(0)%>" size="12" <%=sAmnType%> MAXLENGTH="11">
																				</td>
																				<td class="<%=sCssTab%>" align="center">
																					<input type="text" class="<%=sCssFrm%>" value="<%=objRs(4)%>" name="txtDrSuffix<%=objRs(0)%>" size="7" <%=sAmnType%>>
																				</td>

															                </tr>
															                <%end if %>

																			<%
																				objRs.MoveNext
																				iCount = iCount + 1
																				loop
																				objRs.Close
																			%>
															 </table>

														</td>
														<Input type="hidden" name="hSelBookno" value="<%=sBookNo%>" >
														<Input type="hidden" name="hRowCnt" value="<%=iCount%>" >
														<td align="center">
														</td>
										            </tr>
										            <tr>
														<td align="center" colspan="3" class="MiddlePack">
														</td>
										            </tr>
										            <tr>
													<!--	<td align="center">
															<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
														</td>-->
														<%IF Len(sAmnType) = 0 Then %>
															<Input type="hidden" name="hEditType" value="E">
														<%else%>
															<Input type="hidden" name="hEditType" value="N">
														<%end if %>

														<td valign="top" colspan=2>
																		<table border="0" cellpadding="0" cellspacing="0" width="100%">
																			<tr>
																				<td valign="middle" class="ActionCell">
																					<p align="center">
																					<input type="button" value="Save" name="B4" class="ActionButton" onClick="validateForm()" >
																					<input type="button" value="Delete " name="B3" class="ActionButton" <%=sAmnType%> onClick="DelBook()">
																					<input type="button" value="Close" name="B2" class="ActionButton" onClick="FormClose()">
																				</td>
																			</tr>
																		</table>
														</td>
													<!--	<td align="center">
															<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
														</td>-->
								            </tr>
											<tr>
												<td align="center" colspan="3" class="BottomPack">
												</td>
											</tr>
											</table>
											</div>
										</td>
										<td width=5>
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
				</tr>
				<% if trim(sFromAccNo)<>"0" then %>
				<tr>
					<td align="center" colspan="3" class="MiddlePack">
						<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
					</td>
				</tr>
				<tr>
				<td colspan=3>
				<table class=ExcelTable cellspacing=0 cellpadding=0 width=100%>
					<tr>
					<td colspan=3>
						<div>
							<table class="CollapseBand" cellspacing="0" cellpadding="0">
								<tr>
									<td valign="center"><a style="width: 1em; height: 1em;" title href="#" onclick="DivClick('DivContra'); return false;" >
										<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: pointer;" id=imgContra  border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
										</a>
									</td>
									<td valign="center" class="SubTitle">&nbsp;&nbsp;
										<b>Contra Mapping</b>
									</td>
								</tr>
							</table>
						</div>
					</tr>
					<tr>
					<td>
						<div id=DivContra class="frmBody" style="display:none">
						<table class="bodytable">
							<tr>
								<td align=center></td>
								<td valign=top width=100%>
									<center>
										<table cellspacing=0 cellpadding=0 width=100%>
											<tr>
											<td class="GroupTitleLeft" width=10>&nbsp;
											</td>
											<td class="GroupTitle" width=60 align=center >Contra
											</td>
											</center>
											<td class="GroupTitleRight"><p align=left>&nbsp;
											</td>
											</tr>
											<tr>
												<td	class="GroupTable" colspan=3>
													<center>
													<table>
														<tr>
															<td align="center" colspan="3" class="MiddlePack" height="7">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
														</tr>
														<tr>
															<td align="center" width="5">
															</td>
															<td class="FieldCellSub">
															<p align=center><B>Contra Mapping For <%=sFromAccName%></B></p>
															</td>
														</tr>

														<tr>
															<td align="center" width="5">
															</td>
															<td class="FieldCellSub" height=10>
															</td>
														</tr>

														<tr>
															<td align="center" width="5">
															</td>
															<td class="FieldCellSub" valign=top>
																<%
																	sQuery = "Select a.AccountHead,b.AccountHeadCode,b.AccountDescription from Acc_R_OrgGLAccountHead a,"&_
																			 "Acc_M_GLAccountHead b Where a.OUDefinitionID='"& iUnitno &"' and a.EligibleForContras=1 and "&_
																			 "a.AccountHead=b.AccountHead and a.SubLedger=0 and A.AmendmentExists = '0' and a.AccountHead<>"& sFromAccNo &" and a.AccountHead Not in ("&_
																			 " select a.ToAccountHead from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
																			"where a.OUDefinitionID='"& iUnitno &"' and b.AccountHead=a.FromAccountHead and a.FromAccountHead = "& sFromAccNo  &") "&_
																			" and a.AccountHead in (Select BookAccountHead from Acc_R_ApplicableAccountHeads where BookAccountHead = b.AccountHead and Useable=0)"&_
																			"Order By b.AccountDescription "
																			' Response.Write sQuery
																	objRs.Open sQuery,con
																	if not objRs.EOF then
																		Response.Write "<input type=hidden name=hToAccHead value='Y'>"
																		Response.Write "Map To Book  : "
																		Response.Write "<Select name=selToAccHead class=FormElem size=5 multiple>"

																		do while not objRs.EOF
																			Response.Write "<option value="& trim(objRs(0)) &">"&trim(objrs(2))&"</option>"
																			objRs.MoveNext
																		loop
																		Response.Write "</Select>"
																	else
																		Response.Write "<input type=hidden name=hToAccHead value='N'>"
																		Response.Write "Map To Book : No Books Available For Mapping"
																	end if
																	objRs.Close
																%>

															</td>
														</tr>

														<tr>
															<td align="center" height=10 colspan="3">
															</td>
														</tr>

														<tr>
															<td align="center" class="BottomPack" colspan="3">
																<table border="0" cellpadding="0" cellspacing="0" width="100%">
																	<tr>
																		<td valign="middle" class="ActionCell">
																			<p align="center">
							                                                    <input type="button" value="Done" name="B2" class="ActionButton" onclick="CheckSubmit()" >&nbsp;
							                                                    <input type="reset" value="Reset" name="B1" class="ActionButton" >
																		</td>
																	</tr>
																</table>
															</td>
														</tr>

														<tr>
															<td align="center" height=10 colspan="3">
															</td>
														</tr>

														<tr>
															<td align="center" width="5">
															</td>
															<td valign="top" width="100%">

																<DIV class=frmBody id=frm1 style="width: 415; height:200">
							                                        <table id="tblMap"  border="0" cellspacing="1" class="ExcelTable" width="100%">
							                                            <tr>
																			<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
																			<td class="ExcelHeaderCell" align="center">
																				<img style="cursor: pointer;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onClick="DelMapBook()">
																			</td>
																			<td class="ExcelHeaderCell" align="center" width="100%">Books Already Mapped</td>
							                                            </tr>
							<%
									iSno=0
									sQuery="select b.AccountDescription,a.ToAccountHead from Acc_M_ContraEntries a,Acc_M_GLAccountHead b "&_
											"where a.OUDefinitionID='"& iUnitno &"' and b.AccountHead=a.FromAccountHead and a.FromAccountHead="& sFromAccNo
									'Response.Write sQuery
									with objRs
										.CursorLocation = 3
										.CursorType = 3
										.Source = sQuery
										.ActiveConnection = con
										.Open
									end with

									set objRs.ActiveConnection = nothing

									set sAccName=objRs(0)
									set iToAccCode=objRs(1)

									if not objRs.EOF then
										do while not objRs.EOF
										iRecordsCount = 0
											iSno=cint(iSno)+1
											sQuery="select AccountDescription from Acc_M_GLAccountHead "&_
													"where AccountHead="&iToAccCode
											with objRs1
												.CursorLocation = 3
												.CursorType = 3
												.Source = sQuery
												.ActiveConnection = con
												.Open
											end with
											set objRs1.ActiveConnection = nothing

											if not objRs1.EOF then
												sToAccName=objRs1(0)
											end if
											objRs1.Close

											sQuery = "Select BookCode,BookNumber from Acc_R_ApplicableAccountHeads where BookAccountHead = "& iToAccCode
											'Response.Write sQuery
											objRs1.Open sQuery,con
											if not objRs1.EOF then
												iBookCode = objRs1(0)
												iBookNumber = objrs1(1)
												sQuery = "Select Count(CreatedTransNo) from ACC_T_CreatedVoucherheader where BookCode = "& iBookCode  &" and BookNumber = "& iBookNumber
												'Response.Write sQuery
												objRs2.Open sQuery,con
												if not objRs2.EOF then
													iRecordsCount = objRs2(0)
												end if
												objRs2.Close
											end if
											objRs1.Close

										'	Response.Write "iRecordsCount = "& iRecordsCount

							%>
							                <tr>
												<td class="ExcelSerial" align="center"><%=iSno%></td>
												<td class="ExcelDisplayCell" align="center">
												<% IF iRecordsCount=0 THEN %>
													<input type="checkbox" name="chkBox<%=iSno%>" class=FormElem value="<%=iToAccCode%>">
												<% Else%>
													<input type="checkbox" name="chkBox<%=iSno%>" class=FormElem value="<%=iToAccCode%>" disabled>
												<% End IF%>
												</td>
												<td class="ExcelDisplayCell"><%=sToAccName%></td>
							                </tr>
							<%
										objRs.MoveNext
										loop
									end if
							%>
								<input type=hidden name="hRowContraCnt" value="<%=iSno%>">
							                                        </table>
																</div>
															</td>
															<td align="center">
															</td>
							                            </tr>
							                                <tr>
															<td align="center" class="MiddlePack" colspan="3">
															</td>
							                                </tr>
														<tr>
															<td align="center" width="5" class="ClearPixel">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
															<td valign="top">
																			<table border="0" cellpadding="0" cellspacing="0" width="100%">
																				<tr>
																					<td valign="middle" class="ActionCell">
							                                                            <p align="center">
							                                                                <input type="button" value="Close" name="B7" onclick="window.close()" class="ActionButton" >&nbsp;
																					</td>
																				</tr>
																			</table>
															</td>
															<td align="center" class="ClearPixel" width="5">
																<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
															</td>
														</tr>
							                                <tr>
															<td align="center" class="BottomPack" colspan="3">
															</td>
							                                </tr>
													</table>
													</center>
												</td>
											</tr>
										</table>
								</td>
							</tr>
						</table>
						</td>
						</tr>
						</table>
					</div>
				</td>
				</tr>
				<%end if 'if trim(sFromAccNo)<>"" then%>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</HTML>
