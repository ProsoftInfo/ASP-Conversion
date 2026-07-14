<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ParUnitPopup.asp
	'Module Name				:	ACCOUNTS (Master Creation)
	'Author Name				:	Ragavendran R
	'Created On					:	Nov 17,2010
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
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/accpopulate.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<%
dim objRs
'XML DOM Variables
Dim oDOM,nodHeader,Root,nodBook,nodUnit,nodTemp,sParName,sParCode,iPartyCode,Objrs2,Objrs3
Dim MainNode,UnitNode,sCDIndicate,sAction
' Create our DOM Document Objects

iPartyCode = Request.QueryString("PartyCode")
sAction = Request.QueryString("Action")

Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
Set objRs = Server.CreateObject("ADODB.RecordSet")
Set objRs2 = Server.CreateObject("ADODB.RecordSet")
Set objRs3 = Server.CreateObject("ADODB.RecordSet")

'Response.Write Session.SessionID

oDOM.Load server.MapPath("../temp/master/Party_Master_"&Session.SessionID&".xml")
Set Root = oDOM.documentElement
For Each nodHeader In Root.childNodes
	if StrComp(nodHeader.nodeName,"ParName") = 0 then
		sParName=nodHeader.text
	end if
	if StrComp(nodHeader.nodeName,"ShortName") = 0 then
		sParCode=nodHeader.text
	end if
	if StrComp(nodHeader.nodeName,"Units") = 0 then
		set nodUnit=nodHeader
	end if
next

if trim(sAction)="EDIT" then
	Set oDOM = Nothing
	Set oDom = server.CreateObject("Microsoft.xmlDom")
	Set Root = oDom.createElement("Root")
	oDom.appendChild Root

	For Each nodTemp In nodUnit.childNodes

			sUnitId=trim(nodTemp.Attributes.Item(0).nodeValue)
			sUnitName=nodTemp.Attributes.Item(1).nodeValue
			'Response.Write sUnitId &"<br>"
			'sQuery="Select  P.PartyType,P.PartySubType,V.SubTypeName,P.OpeningAmount From "&_
			'		"App_M_PartyTypes V, Acc_T_PartyOpeningAmt P Where P.PartyCode = "&iPartyCode&"  "&_
			'		"and P.OUDefinitionID = '"&sUnitId&"' "&_
			'		"and V.PartyType = P.PartyType and V.PartySubType = P.PartySubType "

		'	sQuery="select PartyType,PartySubType,SubTypeName  from vwOrgPartyType "&_
		'			"where OUDefinitionID='"&sUnitId&"'"

		    sQuery = "Select PartyType,PartySubType,SubTypeName from APP_M_PartyTypes "

			with objRs
					.CursorLocation =3
					.CursorType =3
					.Source = sQuery
					.ActiveConnection = con
					.Open
			end with
			set objRs.ActiveConnection=nothing

			set iParType=objRs(0)
			set iParSubType=objRs(1)
			set sParSubTypeName=objRs(2)
			'set iOpenAmt = objRs(3)
			iOpenAmt = 0

			iParTypeCount=objRs.RecordCount
			bFirst=true
			if iParTypeCount >0 then
				do while not objRs.EOF

					sQuery = "Select H.PartyType From Acc_T_VoucherHeader V,Acc_T_CreatedVoucherHeader H "&_
						     "Where H.OUDefinitionID = '"&sUnitId&"' and V.OUDefinitionID = '"&sUnitId&"' "&_
						     "and H.PartyType = '"&iParType&"' and V.PartyType = '"&iParType&"' and H.PartySubType = "&iParSubType&" "&_
						     "and V.PartySubType = "&iParSubType&" and V.PartyCode = "&iPartyCode&" and H.PartyCode = "&iPartyCode&" "

					' Response.Write sQuery
					With Objrs2
						.CursorLocation = 3
						.CursorType = 3
						.ActiveConnection = Con
						.Source = sQuery
						.Open
					End With
					iRecCount = Objrs2.RecordCount
					Set Objrs2.ActiveConnection = Nothing
					Objrs2.Close

					sQuery="Select  P.PartyType,P.PartySubType,V.SubTypeName,P.OpeningAmount From "&_
							"App_M_PartyTypes V, Acc_T_PartyOpeningAmt P Where P.PartyCode = "&iPartyCode&"  "&_
							"and P.OUDefinitionID = '"&sUnitId&"' "&_
							"and V.PartyType = P.PartyType and V.PartySubType = P.PartySubType "

					'Response.Write sQuery
					With Objrs2
						.CursorLocation = 3
						.CursorType = 3
						.ActiveConnection = Con
						.Source = sQuery
						.Open
					End With
					Set Objrs2.ActiveConnection = Nothing
					IF Not Objrs2.EOF Then
						iOpenAmt = Objrs2(3)


					Else
						iOpenAmt = 0

					End IF
					Objrs2.Close

					sQuery = "Select OpeningBalance,OpeningCDIndication From APP_R_OrgParty Where PartyCode = "&iPartyCode&" and PartyType = '"&iParType&"' "&_
							 "and PartySubType = "&iParSubType&" and OUDefinitionID = '"&sUnitId&"' "


					With Objrs3
						.CursorLocation = 3
						.CursorType = 3
						.ActiveConnection = Con
						.Source = sQuery
						.Open
					End With
					Set Objrs3.ActiveConnection = Nothing
					sPartyCheck = Objrs3.RecordCount
					IF Not Objrs3.EOF Then
						iOpenAmt = Objrs3(0)
						sCDIndicate = Objrs3(1)
					Else
						iOpenAmt = 0
						sCDIndicate = "C"
					End IF
					Objrs3.Close

					IF CStr(sPartyCheck) = "0" Then
						sType = ""
					Else
						sType = "Checked"
					End IF


					IF CStr(iRecCount) = "0" Then
						sAppType = "0"
						sStat = 1
						'sType = ""
					Else
						sAppType = "1"
						sStat = 0
						'sType = "checked"
					End IF

					sAppType = "0"

					Set Mainnode = oDom.createElement("Party")
					Mainnode.setAttribute "Enqno", objrs(0)
					Mainnode.setAttribute "Units", sUnitId
					Mainnode.setAttribute "PartyType", iParType
					Mainnode.setAttribute "PartySubType", iParSubType
					Mainnode.setAttribute "OpenAmt", iOpenAmt
					Mainnode.setAttribute "CRDRIndicate", sCDIndicate
					Mainnode.setAttribute "DelOption", sAppType
					Root.appendChild MainNode

					objRs.MoveNext
				loop
			end if
			objRs.Close
			sType = ""
		next

		oDOM.Save server.MapPath("../Temp/Master/"&Session.SessionID&"-PartyCheck.xml")
end if 'if trim(sAction)="EDIT" then



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<base target="_self"></base>
<TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<script type="application/xml" data-itms-xml-island="1" id="PartyData" data-src="<%="../temp/master/Party_Master_"&Session.SessionID&".xml"%>" ></script>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT SRC="../../scripts/trim.js"></SCRIPT>
<SCRIPT SRC="../../scripts/cancel.js"></SCRIPT>
<script src="/Scripts/itms-modern-compat.js"></script>
<script>
function CheckSubmit()
{
var bFlag;
bFlag=false;
	for (i=1;i<document.formname.elements.length;i++)
	{
		if(document.formname.elements[i].type=="text")
			{
				if (trim(document.formname.elements[i].value)!="")
				{
					if (isNaN(document.formname.elements[i].value))
					{
						alert("Enter Numeric Value");
						document.formname.elements[i].select();
						return false;
					}
				}
			}
	}
	for (i=1;i<document.formname.elements.length;i++)
	{
		if(document.formname.elements[i].type=="checkbox")
			{
				if (document.formname.elements[i].checked==true)
				{
					bFlag= true;
					break;
				}
			}
	}
if (bFlag==false)
{
	alert("Select atleast one party type");
	return false;
}
return true;
}
</script>
<script src="../../scripts/ModalReturnCompat.js"></script>
<script>
window.ITMSModalReturnCompat.install(function () {
	return window.ITMSModalReturnCompat.xmlIsland("PartyData");
});
</script>
<script>
window.__itmsPopupCompat = { type: "partyUnitPopup" };
</script>
<script src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">

<form method="POST" name="formname">
<input type=hidden name=hFinFromYear value="<%=getFromFinYear%>" >
<input type=hidden name=hFinToYear value="<%=getToFinYear%>" >
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">
		<%	if trim(sAction)="CREATE" then
				Response.Write "Party Creation"
			else
				Response.Write "Party Amendment"
			end if 'if trim(sAction)="CREATE" then%>
		</p>
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
								<td align="center">
								</td>
								<td valign="top" width="100%">
                                    <table cellpadding="0" cellspacing="0" width="100%">
                                    <tr>
										<td class="FieldCell" width="115" valign="top">Party Name</td>
										<td class="FieldsubCell" ><span class="DataOnly"><%=sParName%></td>
								  </tr>
								  <tr>
										<td class="FieldCell" width="115" valign="top">Party Code</td>
										<td class="FieldsubCell"><span class="DataOnly"><%=sParCode%></td>
								  </tr>
                                 <tr>
                            <td class="FieldCell" width="115" valign="top">Opening MonthYear</td>
                            <td>
                            <input type="text" value="<%=getFromFinYear%>" name="txtOpenYear" readonly maxlength="6" size="7" class="FormElemRead">
                            </td>
                                </tr>
                                    </table>
								</td>
								<td align="center">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
												<DIV class=frmBody id=frm1 style="width: 585; height:240;">
                                                <table border="0" cellspacing="1" class="ExcelTable" width="569">
                                            <tr>
                                        <td class="ExcelHeaderCell" align="center" valign="Top" rowspan="2" width="25"><p align="center">S.No.</td>
                                        <td class="ExcelHeaderCell" align="left" colspan="2" valign="Top">&nbsp;&nbsp;&nbsp;Unit Name</td>
                                        </tr>
                                        <tr>
                                        <td class="ExcelHeaderCell" width="150" align="center">Opening Balance</td>
                                        <td class="ExcelHeaderCell" align="center"> Party Type</td>

                                            </tr>
<%
dim sUnitId,sUnitName,iSno,sQuery,iParTypeCount,bFirst,iOpenAmt,sPartyCheck
dim iParType,iParSubType,sParSubTypeName,iCounter,sAppType,iRecCount,sType
Dim sStat

iSno=1
sStat = 1

	For Each nodTemp In nodUnit.childNodes

		sUnitId=trim(nodTemp.Attributes.Item(0).nodeValue)
		sUnitName=nodTemp.Attributes.Item(1).nodeValue
		'Response.Write sUnitId &"<br>"
		'sQuery="Select  P.PartyType,P.PartySubType,V.SubTypeName,P.OpeningAmount From "&_
		'		"App_M_PartyTypes V, Acc_T_PartyOpeningAmt P Where P.PartyCode = "&iPartyCode&"  "&_
		'		"and P.OUDefinitionID = '"&sUnitId&"' "&_
		'		"and V.PartyType = P.PartyType and V.PartySubType = P.PartySubType "

'		sQuery="select PartyType,PartySubType,SubTypeName  from vwOrgPartyType "&_
'				"where OUDefinitionID='"&sUnitId&"'"

        sQuery = "Select PartyType,PartySubType,SubTypeName from APP_M_PartyTypes "
		'Response.Write sQuery

		with objRs
				.CursorLocation =3
				.CursorType =3
				.Source = sQuery
				.ActiveConnection = con
				.Open
		end with
		set objRs.ActiveConnection=nothing
		'set iOpenAmt = objRs(3)
		iOpenAmt = 0

		iParTypeCount=objRs.RecordCount
		bFirst=true
		if iParTypeCount >0 then
			do while not objRs.EOF
				iParType=objRs(0)
				iParSubType=objRs(1)
				sParSubTypeName=objRs(2)

			'	Response.Write "iParType = "& iParType
			'	Response.Write "iParSubType = "& iParSubType
			'	Response.Write "sParSubTypeName = "& sParSubTypeName

			if trim(sAction) = "EDIT" then

				sQuery = "Select H.PartyType From Acc_T_CreatedVoucherHeader H "&_
					     "Where H.OUDefinitionID = '"&sUnitId&"' "&_
					     "and H.PartyType = '"&iParType&"' and H.PartySubType = "&iParSubType&" "&_
					     "and H.PartyCode = "&iPartyCode&" "

				'Response.Write sQuery
				With Objrs2
					.CursorLocation = 3
					.CursorType = 3
					.ActiveConnection = Con
					.Source = sQuery
					.Open
				End With
				iRecCount = Objrs2.RecordCount
				Set Objrs2.ActiveConnection = Nothing
				Objrs2.Close

				IF Cstr(iRecCount) = "0" Then
					sQuery = "Select AccUnitPartyType From Acc_T_CreatedVoucherDetails H Where "&_
							 "AccountingUnit = '"&sUnitId&"' and AccUnitPartyType = '"&iParType&"' and  "&_
							 "AccUnitPartySubType = "&iParSubType&" and AccUnitPartyCode = "&iPartyCode&" "

					With Objrs2
						.CursorLocation = 3
						.CursorType = 3
						.ActiveConnection = Con
						.Source = sQuery
						.Open
					End With
					iRecCount = Objrs2.RecordCount
					Set Objrs2.ActiveConnection = Nothing
					Objrs2.Close

				End IF

				sQuery="Select  P.PartyType,P.PartySubType,V.SubTypeName,P.OpeningAmount From "&_
						"App_M_PartyTypes V, Acc_T_PartyOpeningAmt P Where P.PartyCode = "&iPartyCode&"  "&_
						"and P.OUDefinitionID = '"&sUnitId&"' "&_
						"and V.PartyType = P.PartyType and V.PartySubType = P.PartySubType "


				With Objrs2
					.CursorLocation = 3
					.CursorType = 3
					.ActiveConnection = Con
					.Source = sQuery
					.Open
				End With
				Set Objrs2.ActiveConnection = Nothing
				IF Not Objrs2.EOF Then
					iOpenAmt = Objrs2(3)

				Else
					iOpenAmt = 0

				End IF
				Objrs2.Close

				sQuery = "Select isNull(OpeningBalance,0),isNull(Useable,'0'),OpeningCDIndication From APP_R_OrgParty Where PartyCode = "&iPartyCode&" and PartyType = '"&iParType&"' "&_
						 "and PartySubType = "&iParSubType&" and OUDefinitionID = '"&sUnitId&"' "



				With Objrs3
					.CursorLocation = 3
					.CursorType = 3
					.ActiveConnection = Con
					.Source = sQuery
					.Open
				End With
				Set Objrs3.ActiveConnection = Nothing
				sPartyCheck = Objrs3.RecordCount
				IF Not Objrs3.EOF Then
					iOpenAmt = Objrs3(0)
					sType = Objrs3(1)
					sPartyCheck = Trim(Objrs3(2))
				Else
					iOpenAmt = 0
				End IF
				Objrs3.Close

				IF CStr(sType) = "0" Then
					sType = "Checked"
				Else
					sType = ""
				End IF

				'IF CStr(sPartyCheck) = "0" Then
					'sType = ""
				'Else
					'sType = "Checked"
				'End IF


				IF CStr(iRecCount) = "0" Then
					sAppType = "R" 'Reference Can be Removed
					sStat = 1
					'sType = ""
				Else
					sAppType = "A" 'Reference Cannot be Removed
					sStat = 0
					'sType = "checked"
				End IF

			end if ' if trim(sAction) = "EDIT" then


				iCounter=iCounter+1
				'Response.Write sPartyCheck
%>
				<tr>
				<%if bFirst then%>
                    <td class="ExcelSerial" width="25" valign="Top" rowspan="<%=iParTypeCount+1%>" align="center"><%=iSno%></td>
                    <td class="ExcelDisplayCell" valign="Top" colspan="2" ><b><%=sUnitName%></b></td>
                    </tr>
                    <tr>
               <%
					bFirst=false
					iSno=iSno+1
					iCounter=1
				end if
				%>
					<td class="ExcelFieldCell" valign="Top">
					<table><tr>
					<td class="ExcelFieldCell">
					<input type="text" value="<%=iOpenAmt%>" style="text-align:right" name="txtBalance<%=sUnitId%>Z<%=iParType%>Z<%=iParSubType%>Z<%=iCounter%>" size="10" class="FormElem"></td>
                    <td class="ExcelFieldCell">
                    <%IF CStr(sPartyCheck) = "D" Then %>
                    <input type="radio" value="D" name="radCRDR<%=sUnitId%>Z<%=iParType%>Z<%=iParSubType%>Z<%=iCounter%>" class="FormElem" Checked>   Dr</td>
                    <td class="ExcelFieldCell"><input type="radio" value="C" name="radCRDR<%=sUnitId%>Z<%=iParType%>Z<%=iParSubType%>Z<%=iCounter%>" class="FormElem">  Cr</td>
                    <%Else%>
                    <input type="radio" value="D" name="radCRDR<%=sUnitId%>Z<%=iParType%>Z<%=iParSubType%>Z<%=iCounter%>" class="FormElem" >   Dr</td>
                    <td class="ExcelFieldCell"><input type="radio" value="C" name="radCRDR<%=sUnitId%>Z<%=iParType%>Z<%=iParSubType%>Z<%=iCounter%>" class="FormElem" Checked>  Cr</td>
                    <%End if %>
                    </tr></table>
                    </td>

                    <td class="ExcelFieldCell" valign="Top">
                    <input type="checkbox"  name="chkParType<%=sUnitId%>Z<%=iCounter%>" value="<%=iParType%>?<%=iParSubType%>?<%=iCounter%>" class="FormElem" <%=sType%> onClick="CheckPartyChk('<%=sAppType%>',this)" ><%=sParSubTypeName%></td>
                    <input type="hidden" name="hParCheck<%=sUnitId%>?<%=iParType%>?<%=iParSubType%>?<%=iCounter%>" value="<%=sStat%>">


                        </tr>

<%
						objRs.MoveNext
					loop
				end if
				%>
					<input type="hidden" name="hRowCntUnitZ<%=sUnitId%>" value="<%=iCounter%>">

				<%

				objRs.Close
				sType = ""
			next
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
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
												<table border="0" cellpadding="0" cellspacing="0" width="100%">
													<tr>
														<td valign="middle" class="ActionCell">
															<p align="center">
															    <input type="button" value="Add Sub Type" name="B5" class="ActionButtonX" onClick="AddNew()">
                                                                <input type="button" value="Save" name="B2" class="ActionButton" onClick="PageSubmit()">
                                                                <input type="button" value="Close" name="B3" class="ActionButton" onClick=window.close()   >
																<input type="reset" value="Reset" name="B1" class="ActionButton" >
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
