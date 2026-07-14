<%@ Language=VBScript %>
<%	option explicit	%>
<%
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/sessionVerify.asp"-->
<!--#include virtual="/include/PurchaseTermsConditions.asp"-->
<!--#include virtual="/include/populate.asp"-->
<!--#include virtual="/include/purpopulate.asp"-->
<!--#include virtual="/include/IncludeDatePicker.asp"-->
<!--#include virtual="/include/PurChkItemSpecPack.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS Rate UOM Selection </title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../scripts/invPurInvEntryRateUomPop.js"></SCRIPT>
<script type="application/xml" id="TempData" data-itms-xml-island="1"><root/></script>
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
Function Done_Clk()

	if document.formname.mCmbRateUOM.value = "0" then
		alert("Select UOM")
		document.formname.mCmbRateUOM.focus
		exit function
	end if

	set obj = document.formname.mCmbRateUOM
	RATEUOM		= obj.value
	CLASSCODE	= document.formname.hClassCode.value
	QTYUOM		= document.formname.hUOM.value

	if Trim(RATEUOM) <> "" then
		Arr1 = split(RATEUOM,":")
		RATEUOM = Arr1(0)
	end if


	if (CLASSCODE = "0" or CLASSCODE="TEMP") and (trim(RATEUOM) <> trim(QTYUOM)) then
		msgbox "Quantity UoM and Rate UoM should be same for Temporary Items",0,"Purchase Invoice"
		obj.focus()
		exit function
	end if

	if trim(document.formname.mTxtNewRate.value) = "" then
		alert("Enter Rate")
		document.formname.mTxtNewRate.focus
		exit function
	end if

	if not IsNumeric(trim(document.formname.mTxtNewRate.value) ) then
		alert("Enter Number")
		document.formname.mTxtNewRate.focus
		exit function
	end if

	if CDbl(document.formname.mTxtNewRate.value) <= 0 then
		alert("Rate should be > 0 ")
		document.formname.mTxtNewRate.focus
		exit function
	end if

	set Root = TempData.DocumentElement
	Root.SetAttribute "RateUOM" , RATEUOM
	Root.SetAttribute "RatePerQtyUoM" , document.formname.mTxtNewRate.value
	window.close()
End Function
'-------------------------------------------------------------------------------------------
Function window_onunload()
	'alert(TempData.xml)
	Set window.returnvalue= TempData.DocumentElement
End Function
'-------------------------------------------------------------------------------------------
</script>
<%
'Declaring Variables
Dim sUOM,sOrgID,sRateUOM

Dim iClassCode,iItemCode,nRatePerQtyUOM,nRet

sOrgID		= Request.QueryString("hOrgID")
iClassCode  = Request.QueryString("hClassCode")
iItemCode   = Request.QueryString("hItemCode")
nRatePerQtyUOM= Request.QueryString("hRatePerQtyUOM")
sUOM		= Request.QueryString("hUOM")
sRateUOM	= Request.QueryString("hRateUOM")

'iClassCode = "TEMP"
'Response.Write "<p><p> " & Request.QueryString

%>
</head>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<form method="POST" name="formname" >
<input type="hidden" name="hClassCode" value="<%=iClassCode %>" >
<input type="hidden" name="hUOM" value="<%=sUOM%>" >

	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopUpTable">
		<tr><td height="1px"></td></tr>
		<tr>
			<td class="PageTitle">
				Purchase Invoice Entry - Item Rate
			</td>
		</tr>

		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" bordercolor="#000000">
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width=100%>
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
										<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
									</td>
								</tr>
								<%  If cstr(iClassCode) <> "TEMP" Then
									nRet = FindPurOptionalUOM(iClassCode,iItemCode,sRateUOM)
								else
									nRet = FindUoMAll(sRateUOM)
								end if

								if trim(nRet) = "0" then
								%>
								<tr>
									<td colspan="3" class="FieldCell" align="center">
										<br>
										<br>
										Optional UOM not exit
										<br>
										<br>
									</td>
								</tr>
								<%
								else
								%>

								<tr>
									<td>
									</td>
									<td valign="top" width="100%">
										<table cellpadding="0" cellspacing="0">
											<tr>
												<td class="FieldCell">UOM</td>
												<td class="FieldCellSub"><%=sUOM%></td>
											</tr>
											<tr>
												<td class="FieldCell">Select Rate UOM</td>
												<td class="FieldCellSub">
													<select size="5" name="mCmbRateUOM" class="FormElem">
														<option value = "0" selected>Select</option>
														<%  If cstr(iClassCode) <> "TEMP" Then
																''To Populate Purchase & OptionalUOM
																popPurOptionalUOM iClassCode,iItemCode,sRateUOM
															Else
																''To Populate All UOM
																populateUoMAll sRateUOM
															End If
														%>

													</select>
												</td>
											</tr>
											<tr>
												<td class="FieldCell">UOM Rate</td>
												<td class="FieldCellSub"><input type="text" name="mTxtNewRate" value="<%=nRatePerQtyUOM%>" class="FormElem" ></td>
											</tr>

										</table>
									</td>
									<td>
									</td>
								</tr>
								<%end if 'if trim(nRet) = "0" then %>

								<tr>
									<td colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td></td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="middle" class="ActionCell">
													<%if trim(nRet) <> "0" then%>
														<input type="button" value="Done" name="B4" class="ActionButton" onclick="Done_Clk()">
													<%end if %>
													<input type="button" value="Close" name="B4" class="ActionButton" onclick="window.close()">
												</td>
											</tr>

										</table>
									</td>
									<td></td>
								</tr>

								<tr>
									<td colspan="3" class="BottomPack">
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
<%
'----------To Get Purchase & Optional UOMs-------------------------------
Function popPurOptionalUOM(iClassCode,iItemCode,sUoM)
Dim iUOMCode,sUOMDesc,sDecimal,sSql

Dim rsTemp

set rsTemp  =server.CreateObject("ADODB.RecordSet")

	''Purchase UOM
	sSql ="Select PurchaseUoM,UoMShortDescription,DECIMALALLOWED from INV_M_ITEMMASTER,MS_UnitOfMeasurement " &_
	 " Where itemcode="&iItemCode&" and classificationcode="&iClassCode&" and organisationcode='"&sOrgID &"'and PurchaseUOM = UOMCode"

	'Response.Write "<p> sSql = " & sSql
	With rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsTemp.ActiveConnection = nothing

	'Response.Write "<p> rsTemp.EOF = "& rsTemp.EOF
	If Not rsTemp.EOF then
		iUOMCode = rsTemp(0)
		sUOMDesc = rsTemp(1)
		sDecimal = rsTemp(2)
		if trim(sUoM) = trim(iUOMCode) then
			Response.Write("<OPTION VALUE="""&trim(iUOMCode)&":"&trim(sDecimal)&""" selected >"&sUOMDesc&"</OPTION>" &vbcrlf)
		else
			Response.Write("<OPTION VALUE="""&trim(iUOMCode)&":"&trim(sDecimal)&""">"&sUOMDesc&"</OPTION>" &vbcrlf)
		end if
	End if
	rsTemp.Close

	''Optional UOM
	sSql ="Select IM.UoMCode,UM.UoMShortDescription,DECIMALALLOWED from Inv_M_ItemOptionalUoM IM,MS_UnitOfMeasurement UM" &_
	 " Where IM.itemcode="&iItemCode&" and IM.classificationcode="&iClassCode&" and organisationcode='"&sOrgID &"' " &_
	 " and IM.UOMCode = UM.UOMCode And IM.OptionalUoMFor='P'"
	With rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsTemp.ActiveConnection = nothing

	'Response.Write "<p> rsTemp.EOF = "& rsTemp.EOF
	If Not rsTemp.EOF then
	do while not rsTemp.EOF
		iUOMCode = rsTemp(0)
		sUOMDesc = rsTemp(1)
		sDecimal = rsTemp(2)
		if trim(sUoM) = trim(iUOMCode) then
			Response.Write("<OPTION VALUE="""&trim(iUOMCode)&":"&trim(sDecimal)&""" selected>"&sUOMDesc&"</OPTION>" &vbcrlf)
		else
			Response.Write("<OPTION VALUE="""&trim(iUOMCode)&":"&trim(sDecimal)&""">"&sUOMDesc&"</OPTION>" &vbcrlf)
		end if
	rsTemp.MoveNext
	loop
	End if
	rsTemp.Close
End Function
'------------------------------
Function populateUoMAll(sUoM)
		' Declaration of variables
		Dim oDom,fs,Root,PGNode
		dim sUoMID,sUoMName,sUoMShName,sDecimal

		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		Set fs = CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(Server.MapPath("../../Inventory/xmldata/UoM.xml")) then

			oDOM.Load server.MapPath("../../Inventory/xmldata/UoM.xml")
			Set Root = oDOM.documentElement
			if Root.HaschildNodes() then
				For Each PGNode In Root.childNodes

					sUoMID = trim(PGNode.Attributes.Item(0).nodeValue)
					sUoMName = trim(PGNode.Attributes.Item(1).nodeValue)
					sUoMShName = trim(PGNode.Attributes.Item(2).nodeValue)
					sDecimal = trim(PGNode.Attributes.Item(3).nodeValue)

					if trim(sUoM) = trim(sUoMID) then
						Response.Write("<OPTION VALUE="""&trim(sUoMID)&":"&trim(sDecimal)&""" selected>"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
					else
						Response.Write("<OPTION VALUE="""&trim(sUoMID)&":"&trim(sDecimal)&""">"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
					end if
				next
			end if
		end if
End Function

'----------To Get Purchase & Optional UOMs-------------------------------
Function FindPurOptionalUOM(iClassCode,iItemCode,sUoM)
Dim ssql

Dim nRet

Dim rsTemp

nRet = 0

set rsTemp  =server.CreateObject("ADODB.RecordSet")


	''Optional UOM
	sSql ="Select IM.UoMCode,UM.UoMShortDescription,DECIMALALLOWED from Inv_M_ItemOptionalUoM IM,MS_UnitOfMeasurement UM" &_
	 " Where IM.itemcode="&iItemCode&" and IM.classificationcode="&iClassCode&" and organisationcode='"&sOrgID &"' " &_
	 " and IM.UOMCode = UM.UOMCode And IM.OptionalUoMFor='P'"
	With rsTemp
		.CursorLocation = 3
		.CursorType = 3
		.Source = sSql
		.ActiveConnection = con
		.Open
	End With
	Set rsTemp.ActiveConnection = nothing

	'Response.Write "<p> rsTemp.EOF = "& rsTemp.EOF
	If Not rsTemp.EOF then
		nRet = rsTemp.recordcount
	End if
	rsTemp.Close
	FindPurOptionalUOM = nRet
End Function
'------------------------------
Function FindUoMAll(sUoM)
		' Declaration of variables
		Dim oDom,fs,Root,PGNode,nRet

		nRet = 0

		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		Set fs = CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(Server.MapPath("../../Inventory/xmldata/UoM.xml")) then

			oDOM.Load server.MapPath("../../Inventory/xmldata/UoM.xml")
			Set Root = oDOM.documentElement
			if Root.HaschildNodes() then
				nRet = Root.childNodes.length
			end if
		end if

		FindUoMAll = nRet
End Function
%>
