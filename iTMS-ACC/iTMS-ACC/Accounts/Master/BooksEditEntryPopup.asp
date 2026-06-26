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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/sessionVerify.asp"-->
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
<XML ID="SeriesNoData"><Root /></XML>
<XML ID="AccData"><Root /></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/trim.js"></SCRIPT>
<SCRIPT language="javascript" SRC="../../scripts/ExcelFunctions.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/cancel.js"></SCRIPT>
<SCRIPT language="vbscript">
'*******************************************
Function DivClick(sValue)
'	alert(sValue)
	if sValue = "DivContra" then
		if DivContra.style.display="block" then
			DivContra.style.display="none"
			document.all.imgContra.src = "../../assets/images/plus.gif"
		else
			DivContra.style.display="block"
			document.all.imgContra.src = "../../assets/images/minus.gif"
			DivBasic.style.display="none"
			document.all.imgBasic.src = "../../assets/images/plus.gif"
		end if
	elseif sValue="DivBasic" then
		if DivBasic.style.display="block" then
			DivBasic.style.display="none"
			document.all.imgBasic.src = "../../assets/images/plus.gif"
		else
			DivBasic.style.display="block"
			document.all.imgBasic.src = "../../assets/images/minus.gif"
			if trim(document.formname.hFromHead.value)<>"0" then
				DivContra.style.display="none"
				document.all.imgContra.src = "../../assets/images/plus.gif"
			end if 'if trim(document.formname.hFromHead.value)<>"0" then
		end if
	end if
End Function
'*******************************************************
Function FunUseable()
	Dim objhttp,sOrgCode,sBookCode,sBookNo,sUseable

	sOrgCode = document.formname.hOrgCode.value
	sBookCode = document.formname.hBookCode.value
	sBookNo = document.formname.hBookNo.value


	set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.open "GET","GetBookUsedDetails.asp?OrgCode="&sOrgCode&"&BookCode="&sBookCode&"&BookNo="&sBookNo,False
		objhttp.send
		if cstr(objhttp.responseText)="Y" then
			alert("The Current Book is Already having transactions could not be change the Useable Status")
			document.formname.submit
			exit function
		end if


	if document.formname.optUseable(0).checked =true then
		sUseable = 0
	else
		sUseable = 1
	end if

	if confirm("Do you want to change the Book Useable?") then
		set objhttp = CreateObject("Microsoft.XMLHTTP")
		objhttp.open "GET","BookUseableUpdate.asp?OrgCode="&sOrgCode&"&BookCode="&sBookCode&"&BookNo="&sBookNo&"&Useable="&sUseable,False
		objhttp.send
		if trim(objhttp.responseText)<>"" then
			if len(objhttp.responseText) > 1 then
				alert(objhttp.responseText)
				exit function
			else
				sUseable = cint(objhttp.responseText)
				if sUseable=0 then
					document.formname.optUseable(0).checked = true
				else
					document.formname.optUseable(1).checked = true
				end if
			end if
		end if
	end if
End Function
'***********************************************************
Function FormClose()
	document.formname.hAction.value = "Close"
	window.close
End Function
'*************************************************************
Function window_onunload()
	window.returnvalue = document.formname.hAction.value
End Function
'**************************************************************
Function DisplayBook()
	dim iSeriesNo,iStNo
	dim Root,sExp,iEntLen
	Set Root = SeriesNoData.documentElement
	ClearTable1 document.formname.selPayRecNo.value
	j=1
	iStNo = 1
	iEntLen = 0
	if document.formname.selNoSeries.selectedIndex<> "0" then
		iSeriesNo= document.formname.selNoSeries.value

		sExp = "//Series[@No="&iSeriesNo&"]/Entry"
		Set TempNode = Root.selectNodes(sExp)
		iEntLen = Tempnode.length

		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = iSeriesNo then
				IF CInt(iEntLen) = 12 Then
					document.formname.hSeriesType.value="M"
				Elseif Cint(iEntLen) = 4 Then
					document.formname.hSeriesType.value="Q"
				Elseif CInt(iEntLen) = 1 Then
					document.formname.hSeriesType.value="Y"
				Else
					document.formname.hSeriesType.value=HeaderNode.Attributes.Item(2).nodeValue
				End IF
				document.formname.hSeriesLen.value=HeaderNode.Attributes.Item(4).nodeValue

				For Each EntryNode In HeaderNode.childNodes

					iEntryNo=EntryNode.Attributes.Item(0).nodeValue
					sPeriod=EntryNode.Attributes.Item(1).nodeValue
					iNumber=EntryNode.Attributes.Item(2).nodeValue
					sPrefix=EntryNode.Attributes.Item(3).nodeValue
					sSufix=EntryNode.Attributes.Item(4).nodeValue

					Select Case HeaderNode.Attributes.Item(3).nodeValue
					   Case "M" sPeriod="Month-"&sPeriod
					   Case "Q" sPeriod="Quater-"&sPeriod
					   Case "Y" sPeriod="Yearly"
					End Select
					if document.formname.selPayRecNo.value="Y" then
						set oRow = document.all.tblBook.insertRow(j)

						InsertCell oRow,1,"",j,"ExcelSerial","Center","",0,0,0,0,""
						InsertCell oRow,1,"",sPeriod,"ExcelDisplayCell","left","",0,0,0,0,""
						InsertCell oRow,2,"txtCrStartNo"&iEntryNo,iStNo,"ExcelInputCell","","",5,7,0,0,""
						InsertCell oRow,2,"txtCrPrefix"&iEntryNo,sPrefix,"ExcelInputCell","","",11,10,0,0,""
						InsertCell oRow,2,"txtCrSuffix"&iEntryNo,sSufix,"ExcelInputCell","","",11,10,0,0,""
						InsertCell oRow,2,"txtDrStartNo"&iEntryNo,iStNo,"ExcelInputCell","","",5,7,0,0,""
						InsertCell oRow,2,"txtDrPrefix"&iEntryNo,sPrefix,"ExcelInputCell","","",11,10,0,0,""
						InsertCell oRow,2,"txtDrSuffix"&iEntryNo,sSufix,"ExcelInputCell","","",11,10,0,0,""
					else

						set oRow = document.all.tblBook.insertRow(j)
						InsertCell oRow,1,"",j,"ExcelSerial","Center","",0,0,0,0,""
						InsertCell oRow,1,"",sPeriod,"ExcelDisplayCell","left","",0,0,0,0,""
						InsertCell oRow,2,"txtStartNo"&iEntryNo,iStNo,"ExcelInputCell","","",5,7,0,0,""
						InsertCell oRow,2,"txtPrefix"&iEntryNo,sPrefix,"ExcelInputCell","","",11,10,0,0,""
						InsertCell oRow,2,"txtSuffix"&iEntryNo,sSufix,"ExcelInputCell","","",11,10,0,0,""
					end if
					j=j+1
				next
			end if
		next
	end if
end Function

'Function popSeriesNo()
'	Dim sSerName,sExp,TempNode,sSerCode
'	sSerCode = document.formname.hCounterType.Value
'	IF CStr(sSerCode) = "" Then
'		Exit Function
'	End IF
'	Set Root = SeriesNoData.documentElement
'	sExp = "//Series[@No="&sSerCode&"]"
'	Set TempNode = Root.selectNodes(sExp)
'	IF TempNode.length <> 0 Then
'		sSerName = TempNode.Item(0).Attributes.Item(1).value
'	End IF
'	document.formname.txtNumSerName.value = sSerName
'
'end Function

Function popSeriesNo()
	Dim iCtr
	Set Root = SeriesNoData.documentElement

	document.formname.selNoSeries.options.length = 0
	document.formname.selNoSeries.length = document.formname.selNoSeries.length+1
	document.formname.selNoSeries.options(document.formname.selNoSeries.length-1).text = "Select Number Series"
	document.formname.selNoSeries.options(document.formname.selNoSeries.length-1).Value = "0"

	For Each HeaderNode In Root.childNodes
		if HeaderNode.Attributes.Item(3).nodeValue = "M" then
			document.formname.selNoSeries.length = document.formname.selNoSeries.length+1
			document.formname.selNoSeries.options(document.formname.selNoSeries.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
			document.formname.selNoSeries.options(document.formname.selNoSeries.length-1).Value =HeaderNode.Attributes.Item(0).nodeValue
		end if
	next

	For iCtr = 0 To document.formname.selNoSeries.length - 1
		IF CStr(document.formname.selNoSeries.options(iCtr).Value) = CStr(document.formname.hCounterType.Value) Then
			document.formname.selNoSeries.selectedIndex = iCtr
			Exit For
		End IF
	Next
end Function



Function ClearTable1(sFlag)
	dim i
	for i=0 to document.all.tblBook.rows.length - 1
		document.all.tblBook.deleteRow(0)
	next
	if sFlag="Y" then
		set oRow = document.all.tblBook.insertRow(0)
		InsertCell oRow,1,"","S.No","ExcelSerial","Center","",0,0,0,0,""
		InsertCell oRow,1,"","Period","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","CR StartNo","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","CR Prefix","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","CR Suffix","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","DR StartNo","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","DR Prefix","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","DR Suffix","ExcelHeaderCell","left","",0,0,0,0,""
	else
		set oRow = document.all.tblBook.insertRow(0)
		InsertCell oRow,1,"","S.No","ExcelSerial","Center","",0,0,0,0,""
		InsertCell oRow,1,"","Period","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","StartNo","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","Prefix","ExcelHeaderCell","left","",0,0,0,0,""
		InsertCell oRow,1,"","Suffix","ExcelHeaderCell","left","",0,0,0,0,""
	end if
end Function

Function validateForm()
	Dim Root,TempNode,iEntLen,sExp,iSeriesNo

	Set Root = SeriesNoData.documentElement
	if Trim(document.formname.txtName.value)="" then
		MsgBox "Enter Book Name"
		 document.formname.txtName.focus
		Exit Function
	end if
	If document.formname.selNoSeries.selectedIndex = 0 Then
		Msgbox "Select No Series "
		document.formname.selNoSeries.focus()
		Exit Function
	end if

	IF Cstr(document.formname.hSeriesType.value) = "" Then
		iSeriesNo = document.formname.selNoSeries.value
		sExp = "//Series[@No="&iSeriesNo&"]/Entry"
		Set TempNode = Root.selectNodes(sExp)
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = iSeriesNo then
				document.formname.hSeriesType.value=HeaderNode.Attributes.Item(2).nodeValue
				document.formname.hSeriesLen.value=HeaderNode.Attributes.Item(4).nodeValue
			End IF
		Next
	End IF

	document.formname.hAction.value = "Done"
	document.formname.hCallType.value = "E"
	document.formname.b3.disabled = True
	document.formname.B4.disabled = True
	document.formname.action = "BooksEditPopupUpdate.asp"
	document.formname.submit

	'MsgBox document.formname.hSeriesType.value
'	sTempValue = "hOrgCode="&document.formname.hOrgCode.value	&"&hBookCode="& document.formname.hBookCode.value&"&selNoSeries="
'	sTempValue = sTempValue & document.formname.selNoSeries(document.formname.selNoSeries.selectedIndex).value
'	sTempValue = sTempValue & "&hSeriesType="& document.formname.hSeriesType.value
'	sTempValue = sTempValue & "&selPayRecNo="& document.formname.selPayRecNo(document.formname.selPayRecNo.selectedIndex).value
'	sTempValue = sTempValue & "&hSeriesLen="& document.formname.hSeriesLen.value&"&txtName="&document.formname.txtName.value
'	if document.formname.optEligible(0).checked =true then
'		sTempValue = sTempValue & "&optEligible="&document.formname.optEligible(0).value
'	else
'		sTempValue = sTempValue & "&optEligible="&document.formname.optEligible(0).value
'	end if
'	sTempValue = sTempValue & "&hBookNo="& document.formname.hBookNo.value
'	sTempValue = sTempValue & "&hCallType="&document.formname.hCallType.value &"&hEditType="& document.formname.hEditType.value&"&hRowCnt="& document.formname.hRowCnt.value
'
'	set objhttp = CreateObject("Microsoft.XMLHTTP")
'	objhttp.open "POST","BooksEditPopupUpdate.asp?"&sTempValue,false
'	objhttp.send
'	if objhttp.responseText="" then
'		window.close
'		document.formname.hAction.value = "Edit"
'	else
'		alert(objhttp.responseText)
'		exit function
'	end if

end Function


Function DelBook()
	document.formname.hAction.value = "Done"
	document.formname.hCallType.value = "D"
	document.formname.b3.disabled = True
	document.formname.B4.disabled = True
	document.formname.action = "BooksEditPopupUpdate.asp"
	document.formname.submit

'	sTempValue = "hOrgCode ="&document.formname.hOrgCode.value	&"&hBookCode="& document.formname.hBookCode.value&"&selNoSeries="
'	sTempValue = sTempValue & document.formname.selNoSeries(document.formname.selNoSeries.selectedIndex).value
'	sTempValue = sTempValue & "&hSeriesType="& document.formname.hSeriesType.value
'	sTempValue = sTempValue & "&selPayRecNo="& document.formname.selPayRecNo(document.formname.selPayRecNo.selectedIndex).value
'	sTempValue = sTempValue & "&hSeriesLen="& document.formname.hSeriesLen.value&"&txtName="&document.formname.txtName.value
'	if document.formname.optEligible(0).checked =true then
'		sTempValue = sTempValue & "&optEligible="&document.formname.optEligible(0).value
'	else
'		sTempValue = sTempValue & "&optEligible="&document.formname.optEligible(0).value
'	end if
'	sTempValue = sTempValue & "&hBookNo="& document.formname.hBookNo.value
'	sTempValue = sTempValue & "&hCallType="&document.formname.hCallType.value &"&hEditType="& document.formname.hEditType.value&"&hRowCnt="& document.formname.hRowCnt.value
'
'	set objhttp = CreateObject("Microsoft.XMLHTTP")
'	objhttp.open "POST","BooksEditPopupUpdate.asp?"&sTempValue,false
'	objhttp.send
'	if objhttp.responseText="" then
'		window.close
'		document.formname.hAction.value = "Delete"
'	else
'		alert(objhttp.responseText)
'		exit function
'	end if
'document.formname.hAction.value = "Done"
End Function
'************************************************************
'Contra Mapping Script Language
'***********************************************************
Function CheckSubmit()
	Dim objHttp,sToAccHead,sFromHead,sOrgCode
	sOrgCode = document.formname.hOrgCode.value
	sFromHead = document.formname.hFromHead.value

	if document.formname.hToAccHead.value="Y" then
		for iCnt = 0 to cint(document.formname.selToAccHead.length)-1
			if document.formname.selToAccHead(iCnt).selected = true then
				sToAccHead = sToAccHead &","& document.formname.selToAccHead(iCnt).value
			end if ' if document.formname.selToAccHead(iCnt).selected = true then
		next
	else
		exit function
	end if
	'alert(sToAccHead)
	if Trim(sToAccHead)<>"" then
		sToAccHead = mid(sToAccHead,2)
	end if
	'alert(sToAccHead)
	set objHttp = CreateObject("Microsoft.XMLHTTP")
	objHttp.open "POST","ContraEntryPopupUpdate.asp?OrgCode="&sOrgCode&"&FromHead="&sFromHead&"&ToHead="&sToAccHead,false
	objHttp.send
	if trim(objHttp.responseText)<>"" then
		alert(objHttp.responseText)
		exit function
	else
		document.formname.hAction.value = "Done"
		window.close
	end if
End Function
'*******************************************************
Function DelMapBook()
Dim nRow,iCnt,sDelItem,objHttp,sFromHead,sOrgCode,iSNo
	nRow = document.formname.hRowContraCnt.value
	sOrgCode = document.formname.hOrgCode.value
	sFromHead = document.formname.hFromHead.value
	for iCnt = 1 to nRow
		if eval("document.formname.chkBox"&iCnt).checked = true then
			sDelItem = sDelItem &","& eval("document.formname.chkBox"&iCnt).value
		end if ' if eval("document.formname.chkBox"&iCnt).checked = true then
	next
	if trim(sDelItem)<>"" then
		sDelItem=mid(sDelItem,2)
	else
		alert("Select Mapped Book to Delete")
		exit function
	end if

	set objHttp = CreateObject("Microsoft.XMLHTTP")
	objHttp.open "POST","ContraEntryPopupDelete.asp?OrgCode="&sOrgCode&"&FromHead="&sFromHead&"&ToHead="&sDelItem,false
	objHttp.send
	if trim(objHttp.responseText)<>"" then
		alert(objHttp.responseText)
		exit function
	else
		ClearTable()
		set objHttp = CreateObject("Microsoft.XMLHTTP")
		objHttp.open "GET","GetMappedContraDetails.asp?OrgCode="&sOrgCode&"&FromHead="&sFromHead,false
		objHttp.send
		if trim(objHttp.responseXML.xml)<>"" then
			AccData.loadXML(objHttp.responseXML.xml)
		else
			alert(objHttp.responseText)
		end if

		set ndRoot = AccData.documentElement
		if ndRoot.hasChildNodes() then
			iSNo = 1
			for each ndAcc in ndRoot.childNodes
				set oRow = document.all.tblMap.insertRow(document.all.tblMap.Rows.length)
				set iCell = oRow.insertCell
				iCell.innerHtml = iSNo
				iCell.className="ExcelHeaderCell"
				iCell.align="left"
				set iCell = oRow.insertCell
				if ndAcc.getAttribute("Records") = "Y" then
					iCell.innerHtml = "<input type=checkbox name=chkbox"&iSNo&" class=FormElem value="&ndAcc.getAttribute("No") &" disabled >"
				else
					iCell.innerHtml = "<input type=checkbox name=chkbox"&iSNo&" class=FormElem value="&ndAcc.getAttribute("No") &" >"
				end if
				iCell.className = "ExcelDisplayCell"
				iCell.align="center"
				set iCell = oRow.insertCell
				iCell.innerHtml = ndAcc.getAttribute("Name")
				iCell.ClassName = "ExcelDisplayCell"
				iCell.align="left"
				iSNo = iSNo +1
			next
		end if
	end if
	document.formname.hRowCnt.value = iSNo-1
End Function
'*******************************************************
Function ClearTable()
	Dim  nRow
	nRow = cint(document.all.tblMap.rows.length)-1
	for iCnt = 1 to nRow
		document.all.tblMap.DeleteRow(1)
	next
End Function
'*******************************************************
Function CreateXML()
    set objhttp = CreateObject("Microsoft.XMLHTTP")
    objhttp.open "GET","../../Admin/Master/XMLGetNoSeriesPattern.asp",false
    objhttp.send
    if trim(objhttp.responseXML.xml)<>"" then
        SeriesNoData.loadXML(objhttp.responseXML.xml)
    else
        alert(objhttp.responseText)
    end if
End Function
'**************************************************************
</script>
<script language="javascript">
window.__itmsPopupCompat = { type: "booksEditEntryPopup" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
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
												<td valign="center"><a style="width: 1em; height: 1em;" title href onclick="DivClick('DivBasic')" >

													<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" id="imgBasic" border="0" src="../../assets/images/minus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
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
									<td valign="center"><a style="width: 1em; height: 1em;" title href onclick="DivClick('DivContra')" >
										<img style=" HEIGHT: 1.8em; WIDTH: 1.8em; cursor: hand;" id=imgContra  border="0" src="../../assets/images/plus.gif" width="10" height="10" alt="Expands this section for more search criteria.">
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
																				<img style="cursor: hand;" border="0" src="../../assets/images/iTMS%20Icons/DeleteIcon.gif" width="15" height="15" onClick="DelMapBook()">
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
