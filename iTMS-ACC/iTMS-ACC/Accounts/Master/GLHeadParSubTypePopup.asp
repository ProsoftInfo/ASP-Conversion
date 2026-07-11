<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GLHeadParSubTypePopup.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 22, 2011
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
<!--#include file="../../include/accpopulate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->

<%
dim sGlHeadName,objRs
'XML DOM Variables
Dim oDOM,nodHeader,Root,nodBook,nodUnit,nodTemp,sGroupName,bSubLedger,iAcchead,ndUN,ndSub
Dim sSubExp,Lednode,iSubSNo,sUnitName
sUnitName = Request.QueryString("UnitName")
Set objRs = Server.CreateObject("ADODB.RecordSet")
' Create our DOM Document Objects
Set oDOM = Server.CreateObject("Microsoft.XMLDOM")

oDOM.Load Server.MapPath("../temp/master/GLAccount_Head_"&Session.SessionID&".xml")

Set Root = oDOM.documentElement
For Each nodHeader In Root.childNodes
    if StrComp(nodHeader.nodeName,"Description") = 0 then
		    sGlHeadName=nodHeader.text
	end if
    if StrComp(nodHeader.nodeName,"GroupCode") = 0 then
	    sGroupName=nodHeader.getAttribute("Name")
    end if
    if StrComp(nodHeader.nodeName,"AccHeadNo") = 0 Then
		iAcchead = nodHeader.text
	End IF
    if nodHeader.nodeName="Units" then
        set nodUnit=nodHeader
        For Each ndUN in nodUnit.childNodes
            if StrComp(ndUN.nodeName,"UN")=0 then
                for each ndSub in ndUN.childNodes
                    if StrComp(ndSub.nodeName,"SubLedger") = 0 then
		                bSubLedger=ndSub.getAttribute("Flag")
	                end if
                next
            end if
	    Next
    end if
next

Set Root = oDOM.documentElement
sSubExp = "//UN/SubLedger[@Flag="&"1"&"]"
Set Lednode = Root.selectNodes(sSubExp)

IF CStr(Lednode.length) = "0" Then
	bSubLedger = "0"
Else
	bSubLedger = "1"
End IF
'Response.Write Lednode.length

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Party Sub Type Selection</TITLE>
<base target="_self"></base>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<xml id="GLHeadData" src="<%="../temp/master/GLAccount_Head_"&Session.SessionID&".xml"%>"></xml>
<xml id="SubTypeData"><Root action="Cancel"/></xml>
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<script src="../../scripts/itms-modern-compat.js"></script>
<script src="../../scripts/ModalReturnCompat.js"></script>
<script>
window.__itmsPopupCompat = { type: "glHeadPartySubTypePopup" };
window.ITMSModalReturnCompat.install(function () {
	return window.ITMSModalReturnCompat.xmlIsland("SubTypeData");
});
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center"><%=sUnitName%><br>Party Sub Type

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
								<td align="center" rowspan="6" width="5" class="ClearPixel">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
                                    <table cellpadding="0" cellspacing="0" width="100%">

                                <tr>
                            <td class="FieldCell" width="145" valign="top">GL Classification</td>
                            <td><span class="DataOnly"><%=sGroupName%></span>
                            </td>
                                </tr>
                                 <tr>
                            <td class="FieldCell" width="145">GL Account Head Name</td>
                            <td>
                            <span class="DataOnly"><%=sGlHeadName%>  </span>
                            </td>
                                </tr>
                            <!--      <tr>
                            <td class="FieldCell" width="145" valign="top">Opening MonthYear</td>
                            <td>

                            <input type="text" value="<%=getFromFinYear%>" name="txtOpenYear" maxlength="6" size="7" class="FormElem">

                            </td>
                                </tr>-->
                                    </table>
								</td>
								<td align="center" rowspan="6" class="ClearPixel" width="5">
                                    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td valign="top" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td valign="top" align="center">
												<DIV class=frmBody id=frm1 style="width: 550; height:300;">
<%

if bSubLedger="1" then



%>
                                                <table border="0" cellspacing="1" class="ExcelTable" width="550">
                                        <tr>
                                            <td class="ExcelHeaderCell" align="center"> Party Sub Type</td>
                                        </tr>
<%
dim sQuery,iOpenAmt,iCounter,sCdindi
dim sUnitId,iSno,bFirst,iParTypeCount,sExp
Dim dCrTotal,dDrTotal,dOpeningAmt,iDrVal,iCrVal
iSno=0
iSubSNo = 0
	'Response.Write iAcchead
	Dim TempNode
	sExp = "//Units/UN"
	Set nodUnit = Root.selectNodes(sExp)
	IF nodUnit.length <> 0 Then

	For iCounter = 0 To nodUnit.length - 1

		iSno=iSno+1
		sUnitId= Trim(nodUnit.Item(iCounter).Attributes.Item(0).value)
'		sUnitName= nodUnit.Item(iCounter).Attributes.Item(1).value
		'Response.Write sUnitId
		sExp = "//Units/UN[@Code="""&sUnitId&"""]/SubLedger"
		Set TempNode = Root.selectNodes(sExp)
		IF TempNode.length <> 0 Then
			bSubLedger = TempNode.item(0).Attributes.Item(0).value
		End IF


'======================= Blocked on 10/02/2005 by Manohar For taking the Total Party Opening amt ===============================
    if trim(iAcchead)<>"" then
		sQuery = "Select OpeningAmount,OpeningCDIndication From  "&_
				 "Acc_T_GLAccOpeningAmt where AccountHead= "&iAcchead&" and OUDefinitionID = '"&Trim(sUnitId)&"' "

		'Response.Write sQuery &"<br>"
		with objRs
			.CursorLocation =3
			.CursorType =3
			.Source =sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection=nothing
		IF Not objRs.EOF Then
			iOpenAmt = objRs(0)
			sCdindi = objRs(1)
		Else
			iOpenAmt = 0
			sCdindi = "C"
		End IF
		objRs.Close
    end if'if trim(iAcchead)<>"" then

		'Response.write iOpenAmt
'======================= Blocked on 10/02/2005 by Manohar For taking the Total Party Opening amt ===============================

'================Taking Opening Amount =====================================================
'sQuery = "Select T.OpeningAmount,T.OpeningCDIndication From Acc_T_PartyOpeningAmt T, "&_
'		 "Acc_R_OrgPartyType R Where R.OUDefinitionID = T.OUDefinitionID and "&_
'		 "R.PartyType = T.PartyType and R.PartySubType = T.PartySubType and "&_
'		 "R.AccountHead = "&iAcchead&" and R.OUDefinitionID = '"&sUnitId&"' and  "&_
'		 "T.OpeningMonthYear = '"&Trim(getFromFinYear)&"' "

if trim(iAcchead)<>"" then
    sQuery = "Select SUM(T.OpeningAmount),T.OpeningCDIndication From Acc_T_PartyOpeningAmt T, "&_
		     "Acc_R_OrgPartyType R,VWOrgParty M Where R.OUDefinitionID = T.OUDefinitionID and "&_
		     "R.PartyType = T.PartyType and R.PartySubType = T.PartySubType and "&_
		     "R.AccountHead = "&iAcchead&" and R.OUDefinitionID = '"&sUnitId&"' and  "&_
		     "T.OpeningMonthYear = '"&Trim(getFromFinYear)&"' and "&_
		     "M.PartyCode = T.PartyCode and M.OUDEFINITIONID = '"&sUnitId&"' and "&_
		     "R.PartyType = M.PartyType and R.PartySubType = M.PartySubType "&_
		     "Group By T.OpeningCDIndication "


    'response.write squery

    With objRs
	    .CursorLocation = 3
	    .CursorType = 3
	    .ActiveConnection = Con
	    .Source = sQuery
	    .Open
    End With

    dCrTotal = 0
    dDrTotal = 0
    Set objRs.ActiveConnection = Nothing
    'Response.write Objrs.RecordCount &" " & sUnitID &"<br>"
    Do While Not objRs.EOF
	    'iDrVal = Trim(objRs(0))
	    'iDrVal = CDbl(iDrVal)

	    IF CStr(objRs(1)) = "C" Then
		    'dCrTotal = CDbl(dCrTotal + iDrVal)
		    dCrTotal = CDbl(Objrs(0))
	    Else
		    'dDrTotal = CDbl(dDrTotal + iDrVal)
		    dDrTotal = CDbl(Objrs(0))
	    End IF
	    objRs.MoveNext
    Loop
    objRs.Close

    iDrVal = 0

    'Response.write dCrTotal &" " & dDrTotal

    IF CDbl(dCrTotal) >= CDbl(dDrTotal) Then
	    dOpeningAmt = CDbl(dCrTotal - dDrTotal)
	    sCdindi = "C"
    Else
	    dOpeningAmt = CDbl(dDrTotal - dCrTotal)
	    sCdindi = "D"
    End IF
else
    dOpeningAmt = 0
end if 'if trim(iAcchead)<>"" then

'================Taking Opening Amount =====================================================

		iParTypeCount=0
			'sQuery="select PartyType, PartySubType, SubTypeShortName FROM APP_M_PartyTypes where "&_
			'			"PartyType+ltrim(str(PartySubType)) NOT in (select PartyType+ltrim(str(PartySubType))"&_
			'			" from Acc_R_OrgPartyType where OUDefinitionID='"&Trim(sUnitId)&"')"

%>
									<tr>
                                        <td class="ExcelDisplayCell" valign="top">

	<%
	    if Trim(iAcchead)<>"" then
	        sQuery=	"select PartyType, PartySubType, SubTypeName from vwOrgPartyType where "&_
					" OUDefinitionID='"&Trim(sUnitId)&"' and AccountHead="&iAccHead&" "

			'Response.Write sQuery &"<br>"
			with objRs
				.CursorLocation =3
				.CursorType =3
				.Source =sQuery
				.ActiveConnection = con
				.Open
			end with
			set objRs.ActiveConnection=nothing
			iParTypeCount=objRs.RecordCount

			IF CStr(bSubLedger) = "1" Then
				do while not objRs.EOF
				    iSubSNo = iSubSNo + 1
	%>
									<input type="checkbox" name="selPartyTypeZ<%=iSubSNo%>"
									value="<%=sUnitId&"?"& objrs(0) &"?" & objrs(1)%>" checked> <%=objrs(0) &"-"& objrs(2)%><br>
	<%
					objRs.MoveNext
				loop
				objRs.Close
		    end if 'IF CStr(bSubLedger) = "1" Then
		 end if 'if Trim(iAcchead)<>"" then

			IF CStr(bSubLedger) = "1" Then
				sQuery="select PartyType, PartySubType, SubTypeName FROM APP_M_PartyTypes where "&_
							"PartyType+ltrim(str(PartySubType)) NOT in (select PartyType+ltrim(str(PartySubType))"&_
							" from Acc_R_OrgPartyType where OUDefinitionID='"&Trim(sUnitId)&"')"



				with objRs
					.CursorLocation =3
					.CursorType =3
					.Source =sQuery
					.ActiveConnection = con
					.Open
				end with
				set objRs.ActiveConnection=nothing
				do while not objRs.EOF
				    iSubSNo = iSubSNo + 1
	%>
									<input type="checkbox" name="selPartyTypeZ<%=iSubSNo%>"
									value="<%=sUnitId&"?"& objrs(0) &"?" & objrs(1)%>"> <%=objrs(0) &"-"& objrs(2)%><br>
	<%
				objRs.MoveNext
				loop
				objRs.Close

			End IF
			

	%>
	</td>

                                       </tr>

<%
	next
	end if

%>



                                                </table>
<%else 'IF Subledger = '0' i.e Not a Party Control Account Head %>
                                                <table border="0" cellspacing="1" class="ExcelTable" width="569">
<%

iSno=0
	For Each nodTemp In nodUnit.childNodes
		iSno=iSno+1
		sUnitId=trim(nodTemp.Attributes.Item(0).nodeValue)
'		sUnitName=nodTemp.Attributes.Item(1).nodeValue

		sQuery = "Select OpeningAmount,OpeningCDIndication From  "&_
				 "Acc_T_GLAccOpeningAmt where AccountHead= "&iAcchead&" and OUDefinitionID = '"&sUnitId&"' "&_
				 "and OpeningMonthYear = '"&getFromFinYear&"' "

		'Response.Write sQuery
		with objRs
			.CursorLocation =3
			.CursorType =3
			.Source =sQuery
			.ActiveConnection = con
			.Open
		end with
		set objRs.ActiveConnection=nothing
		IF Not objRs.EOF Then
			iOpenAmt = objRs(0)
			sCdindi = objRs(1)
		Else
			iOpenAmt = 0
			sCdindi = "C"
		End IF
		objRs.Close
%>
									<tr>
                                        <td class="ExcelSerial"  width="25" align="center"  valign="top"><%=iSno%></td>
                                        <td class="ExcelDisplayCell"  valign="top"><%=sUnitName%></td>
                                        <td class="ExcelFieldCell" width="170" valign="top">
                                        <table width="100%"><tr>
                                        <td class="ExcelFieldCell"><input type="text" name="txtOpenBal<%=sUnitId%>" maxlength="12" size="15" class="FormElem" style="text-align: Right" value="<%=trim(iOpenAmt)%>"></td>
                                        <td class="ExcelFieldCell">
                                        <% IF Trim(CStr(sCdindi)) = "D" Then %>

                                        <input type="radio" value="D" name="optOpenCD<%=sUnitId%>" class="FormElem" checked> Dr
                                        <%else%>
                                        <input type="radio" value="D" name="optOpenCD<%=sUnitId%>" class="FormElem"> Dr
                                        <%end if %>
                                        </td>
                                        <td class="ExcelFieldCell">
                                        <% IF Trim(CStr(sCdindi)) = "C" Then %>

                                        <input type="radio" value="C" name="optOpenCD<%=sUnitId%>" checked class="FormElem"> Cr
                                        <%else%>
                                        <td class="ExcelFieldCell">
                                        <input type="radio" value="C" name="optOpenCD<%=sUnitId%>" class="FormElem"> Cr
                                        <%end if %>
                                        </td>
                                        </tr></table>

                                       </tr>

<%
	next
%>                                              </table>

<%end if%>
</div>
    <input type=hidden name="hSubTypeCount" value="<%=iSubSNo%>"
								</td>
							</tr>
							<tr>
								<td valign="top" class="MiddlePack">
								</td>
							</tr>
							<tr>
								<td valign="top" class="BottomPack">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
															    <input type="button" value="Add Sub Type" name="B3" class="ActionButtonX" tabindex=5 onclick="NewSubType()">
                                                                <input type="button" value="Done" name="B2" class="ActionButton" tabindex="3" onClick="PageSubmit()">
																<input type="reset" value="Reset" name="B1" class="ActionButton" tabindex="4" >
														</td>
													</tr>
												</table>
								</td>
							</tr>
							<tr>
								<td valign="top" class="BottomPack">
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
