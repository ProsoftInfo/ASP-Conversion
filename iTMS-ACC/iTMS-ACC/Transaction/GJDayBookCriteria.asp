<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	GJDayBookCriteria.asp
	'Module Name				:	Accounts (Reports)
	'Author Name				:	N.Rajkumar
	'Created On					:	15th May 2003
	'Modified By				:	UmaMaheswari s
	'Modified On				:	April 19,2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	DBGJView.asp
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
<!--#include file="../../include/DatabaseConnection.asp."-->
<!--#include file="../../include/populate.asp"-->
<!-- #include File="../../include/IncludeDatePicker.asp" -->
<!--#include file="../../include/sessionVerify.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>GJ - Day Book</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<!-- XML Data Island -->
<XML ID="UnitBookData"><Book/></XML>
<XML id="AccHeadData"><account/></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=vbscript>
dim sOptSelect,sOrgId,sBook,sFlag,sOrgName,sVouType,sHead,sHeadCode,saTemp,sHeadDesc,sValue
sBookCode="08"
sFlag="VouNo"
sVouType="B"

'---------------- Function to track the Voucher Type ------------------

Function GetVouType()
If	document.formname.optVoutype(0).checked Then
		sVouType=document.formname.optVoutype(0).value
Elseif document.formname.optVoutype(1).checked then
		sVouType=document.formname.optVoutype(1).value
	Else
		sVouType=document.formname.optVoutype(2).value
End if


End Function

'---------- Selection of Account Head From Pop up Screen --------------

Function SelectAccHead()
dim iGlHead,sOrgId,sAccHead,arrTemp,sRetVal
SelFlag="A"
Set nodAccHead = AccHeadData.documentElement
sHeadDesc = ""

'If document.formname.selUnitId.selectedIndex=0 then
'	Msgbox "Select Organaisation Id",0,"Day Book General Journal"
'	document.formname.selUnitId.focus
if document.formname.selBook.value="S" Then
		Msgbox "Select Book",0,"Day Book Cash Report"
		document.formname.selBook.focus
		Exit Function
Else
	sOrgId=document.formname.hUnitId.value	'document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).value
	sAccHead=document.formname.SelAccHead.options(document.formname.SelAccHead.selectedIndex).Value
	iBookNo=document.formname.selBook.options(document.formname.selBook.selectedIndex).Value
	sValue=sOrgId&"|"&sBookCode&"|"&iBookNo

	Set OutValue = showModalDialog("../Reports/GLHeadSelectionMultiple.asp?orgid="&sOrgId&"&BookId="&sBookCode&"&BookNo="&iBookNo&"&hSelectMode=M",AccHeadData,"dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")

	sAct = UCase(trim(OutValue.getAttribute("Action")))
	sQuery = trim(OutValue.getAttribute("PassQuery"))
	if ucase(trim(sAct)) <> "CLOSE" then

		do while sAct <> "DONE"

			Set OutValue = showModalDialog("../Reports/GLHeadSelectionMultiple.asp?"&sQuery,AccHeadData,"dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
			sAct = UCase(trim(OutValue.getAttribute("Action")))
			if ucase(Trim(sAct)) = "CLOSE" then exit do
			sQuery = trim(OutValue.getAttribute("PassQuery"))
		loop
	end if

	If  OutValue.hasChildNodes Then
		sExp ="//Entry"
		Set AccHeadNode = nodAccHead.Selectnodes(sExp)
		if AccHeadNode.Length > 0 then
			for itr = 0 to AccHeadNode.Length - 1
				sHead=sHead & "," & AccHeadNode.Item(itr).Attributes.getNamedItem("RetField1").Value
				sHeadCode=sHeadCode & "," & AccHeadNode.Item(itr).Attributes.getNamedItem("RetField2").value
				sHeadDesc=sHeadDesc & "," & AccHeadNode.Item(itr).Attributes.getNamedItem("RetField0").value
			next
		end if
		sHead = mid(sHead,2)
		sHeadCode = mid(sHeadCode,2)
		sHeadDesc = Mid(sHeadDesc,2)
		window.spAccHead.innerHTML=sHeadDesc
		document.formname.SelAccHead.selectedIndex =1
	else
		document.formname.SelAccHead.selectedIndex=0
	End if


	If 1 = 2 Then

	OutValue = showModalDialog("GLHeadSelection.asp?Value="&sValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	arrTemp = split(OutValue,":")
	while UBound(arrTemp) = 0
		OutValue = showModalDialog("GLHeadSelection.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend
	sRetVal = OutValue
	if UBound(arrTemp) <= 1 then exit function
	If trim(sRetVal)<>"0" then
		saTemp=split(sRetVal,":")
		SelFlag="S"
		sHead=saTemp(0)
		sHeadCode=saTemp(1)
		sHeadDesc=saTemp(2)
		Window.spAccHead.innerHTML =sHeadDesc
		document.formname.SelAccHead.selectedIndex=1
	Else
		document.formname.SelAccHead.selectedIndex=0
	End if
	End IF	'If 1 = 2 Then
End if
End Function

'----------- Function to Get the option Selected for Display ----------
'------------- Options - VouNo,VouDate,VouAmount,AccHead --------------

Function OptSelection_old()

if document.formname.optCriteria(0).checked then

	sFlag=document.formname.optCriteria(0).value
	document.formname.txtNoFrom.readOnly =false
	document.formname.txtNoTo.readOnly =false
	document.formname.txtGAmount.value =""
	document.formname.txtLAmount.value =""
	document.formname.txtGAmount.readOnly=True
	document.formname.txtLAmount.readOnly=True
	document.formname.CtlVouFromDate.disabled=true
	document.formname.CtlVouToDate.disabled=true
	document.formname.SelAccHead.disabled=true
	window.spAccHead.innerHTML =""

Elseif document.formname.optCriteria(1).checked then
	sFlag=document.formname.optCriteria(1).value
	document.formname.txtNoFrom.value =""
	document.formname.txtNoTo.value =""
	document.formname.txtGAmount.value =""
	document.formname.txtLAmount.value =""
	document.formname.txtNoFrom.readOnly =true
	document.formname.txtNoTo.readOnly =true
	document.formname.txtGAmount.readOnly=true
	document.formname.txtLAmount.readOnly =true
	document.formname.SelAccHead.disabled=true
	window.spAccHead.innerHTML =""


Elseif document.formname.optCriteria(2).checked then
	sFlag=document.formname.optCriteria(2).value
	document.formname.txtNoFrom.value =""
	document.formname.txtNoTo.value =""
	document.formname.txtNoFrom.readOnly =true
	document.formname.txtNoTo.readOnly =true
	document.formname.SelAccHead.disabled=true
	document.formname.txtGAmount.readOnly=false
	document.formname.txtLAmount.readOnly =false
	window.spAccHead.innerHTML =""

Elseif document.formname.optCriteria(3).checked then
	sFlag=document.formname.optCriteria(3).value
	document.formname.txtNoFrom.value =""
	document.formname.txtNoTo.value =""
	document.formname.txtGAmount.value =""
	document.formname.txtLAmount.value =""
	document.formname.txtNoFrom.readOnly =true
	document.formname.txtNoTo.readOnly =true
	document.formname.txtGAmount.readOnly=true
	document.formname.txtLAmount.readOnly =true
	document.formname.SelAccHead.disabled=false
End if
End Function

'----------------- Function Numeric Data in a Control -----------------

Function checkNumbers(val)
	dim valid,temp,i
	valid = "0123456789"
	for i=1 to len(val)
		temp = mid(val,i,1)
		if Instr(1,valid,temp) > 0 then
			checkNumbers = true
		else
			checkNumbers = false
			exit for
		end if
	next
end Function

Function SelNew()
document.formname.SelAccHead.selectedIndex=0
document.formname.txtGAmount.value  =""
document.formname.txtLAmount.value =""
document.formname.txtNoFrom.value =""
document.formname.txtNoTo.value =""
window.spAccHead.innerHTML =""
End Function

'------------- Function to Validate and Submit the Values -------------

Function CheckSubmit_old(sCallTy)
dim iVocNoFrom,iVocNoTo,sFromDate,sToDate,dGAmount,dLAmount,sAccHead,iBookNo


	'If document.formname.selUnitId.selectedIndex=0 then
	'	Msgbox "Select Organaisation Id",0,"Day Book General Journal"
	'	document.formname.selUnitId.focus
	'	Exit Function
	if document.formname.selBook.selectedIndex =0 then
		Msgbox "Select Book ",0,"Day Book General Journal"
		document.formname.selBook.focus
		Exit Function
	Else
		sOrgId=document.formname.hUnitId.value	'document.formname.selUnitId.options(document.formname.selUnitId.options.SelectedIndex).Value
		sOrgName=document.formname.hUnitName.value	'document.formname.selUnitId.options(document.formname.selUnitId.selectedIndex).Text
		sBook=document.formname.selBook.options(document.formname.selBook.options.SelectedIndex).Text
		iBookNo=document.formname.selBook.options(document.formname.selBook.options.selectedIndex).value

	End if

'------------- Coding For the case VouNo is Selected ------------------

	If sFlag="VouNo" Then

		If  document.formname.txtNoFrom.value="" Then
			Msgbox "Enter Voucher No. From ",0,"Day Book General Journal"
			document.formname.txtNoFrom.select
			Exit Function
		ElseIf document.formname.txtNoTo.value ="" Then
			Msgbox "Enter Voucher No. To ",0,"Day Book General Journal"
			document.formname.txtNoTo.select
			Exit Function

		'Elseif not(checkNumbers(document.formname.txtNoFrom.value)) Then
		'		Msgbox "Enter Numbers Only",0,"Day Book General Journal"
		'		document.formname.txtNoFrom.select
		'		Exit Function
		'Elseif not(checkNumbers(document.formname.txtNoTo.value)) Then
		'		Msgbox "Enter Numbers Only",0,"Day Book General Journal"
		'		document.formname.txtNoTo.select
		'		Exit Function
		Else
			iVocNoFrom=document.formname.txtNoFrom.value
			iVocNoTo=document.formname.txtNoTo.value

			If iVocNoFrom > iVocNoTo Then
				Msgbox "Voucher Number Should be Greater Than Previous",0,"Day Book General Journal"
				document.formname.txtNoTo.select
				Exit Function
			Else
				IF CStr(sCallTy) = "S" Then
					showModalDialog "DBGJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&iVocNoFrom&"|"&iVocNoTo&"|"&sVouType,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
				Else
					showModalDialog "PrnDJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&iVocNoFrom&"|"&iVocNoTo&"|"&sVouType,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
				End IF

			End if
		End if

'------------- Coding For the case VouDate is Selected ----------------

	Elseif sFlag="VouDate" Then
			sFromDate=document.formname.CtlVouFromDate.GetDate
			sToDate=document.formname.CtlVouToDate.GetDate
			If dateDiff("d",sFromDate,sToDate)<0 Then
				Msgbox "To Date Should be Greater than From Date",0,"Day Book General Journal"
			Exit Function
			Else
				IF CStr(sCallTy) = "S" Then
					showModalDialog "DBGJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&sFromDate&"|"&sToDate&"|"&sVouType,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
				Else
					showModalDialog "PrnDJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&sFromDate&"|"&sToDate&"|"&sVouType,"A","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No"
				End IF
			End if

'----------- Coding For the case Voucher Amount is Selected -----------

	Elseif sFlag="Amount" Then
		If	document.formname.txtGAmount.value="" Then
				Msgbox "Enter From Amount",0,"Day Book General Journal"
				document.formname.txtGAmount.select
				Exit Function
		Elseif document.formname.txtLAmount.value="" Then
				Msgbox "Enter To Amount",0,"Day Book General Journal"
				document.formname.txtLAmount.select
				Exit Function

		Elseif not(checkNumbers(document.formname.txtGAmount.value)) Then
				Msgbox "Enter Numbers Only",0,"Day Book General Journal"
				document.formname.txtGAmount.select
				Exit Function
		Elseif not(checkNumbers(document.formname.txtLAmount.value)) Then
				Msgbox "Enter Numbers Only",0,"Day Book General Journal"
				document.formname.txtLAmount.select
				Exit Function
		Else
			dGAmount=cdbl(document.formname.txtGAmount.value)
			dLAmount=cdbl(document.formname.txtLAmount.value)
			If cdbl(dGAmount)>cdbl(dLAmount) Then
				Msgbox "To Amount Should be Greater From Amount",0,"Day Book General Journal"
				document.formname.txtLAmount.value =""
				document.formname.txtLAmount.select
				Exit Function
			Else
				IF CStr(sCallTy) = "S" Then
					showModalDialog "DBGJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&dGAmount&"|"&dLAmount&"|"&sVouType,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
				Else
					showModalDialog "PrnDJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&dGAmount&"|"&dLAmount&"|"&sVouType,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
				End IF
			End if
		End if

'------------- Coding For the case VouNo is Selected ------------------

	Elseif sFlag ="AccHead" Then

		if document.formname.SelAccHead.value="0" then
			Msgbox "Select Account Head ",0,"Day Book General Journal"
			document.formname.SelAccHead.focus
			Exit Function
		Else
			sAccHead=document.formname.SelAccHead.options(document.formname.SelAccHead.selectedIndex).Value
			IF CStr(sCallTy) = "S" Then
				showModalDialog "DBGJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&sHead&"|"&sHeadCode&"|"&sVouType&"|"&sHeadDesc,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
			Else
				showModalDialog "PrnDJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&sHead&"|"&sHeadCode&"|"&sVouType&"|"&sHeadDesc,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
			End IF
		End if

	End if
End Function

Function CheckSubmit(sCallTy)
	Dim iVocNoFrom,iVocNoTo,sFromDate,sToDate,dGAmount,dLAmount,sAccHead,iBookNo,sCheck

 	sFromDate = document.formname.ctlVouFromDate.GetDate
	sToDate = document.formname.ctlVouToDate.GetDate
	sMinDate = document.formname.hFDate.value
	sMaxDate = document.formname.hTDate.value
	sCheck =False
	sFlag = ""

 	If document.formname.selBook.selectedIndex =0 then
		Msgbox "Select Book ",0,"Day Book General Journal"
		document.formname.selBook.focus
		Exit Function
	Else
		sOrgId=document.formname.hUnitId.value
		sOrgName=document.formname.hUnitName.value
		sBook=document.formname.selBook.options(document.formname.selBook.options.SelectedIndex).Text
		iBookNo=document.formname.selBook.options(document.formname.selBook.options.selectedIndex).value
	End if

'------------- Coding For the case VouNo is Selected ------------------

	'If sFlag="VouNo" Then
	If document.formname.chkBox1.checked Then

		sCheck = True
		sFlag = sFlag & "," & document.formname.chkBox1.value

		If  document.formname.txtNoFrom.value="" Then
			Msgbox "Enter Voucher No. From ",0,"Day Book General Journal"
			document.formname.txtNoFrom.select
			Exit Function
		ElseIf document.formname.txtNoTo.value ="" Then
			Msgbox "Enter Voucher No. To ",0,"Day Book General Journal"
			document.formname.txtNoTo.select
			Exit Function
		Else
			iVocNoFrom=document.formname.txtNoFrom.value
			iVocNoTo=document.formname.txtNoTo.value

			If iVocNoFrom > iVocNoTo Then
				Msgbox "Voucher Number Should be Greater Than Previous",0,"Day Book General Journal"
				document.formname.txtNoTo.select
				Exit Function
			Else
				'IF CStr(sCallTy) = "S" Then
				'	showModalDialog "DBGJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&iVocNoFrom&"|"&iVocNoTo&"|"&sVouType,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
				'Else
				'	showModalDialog "PrnDJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&iVocNoFrom&"|"&iVocNoTo&"|"&sVouType,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
				'End IF

			End if
		End if
	End IF
'------------- Coding For the case VouDate is Selected ----------------

	'Elseif sFlag="VouDate" Then
			sFromDate=document.formname.CtlVouFromDate.GetDate
			sToDate=document.formname.CtlVouToDate.GetDate
			If dateDiff("d",sFromDate,sToDate)<0 Then
				Msgbox "To Date Should be Greater than From Date",0,"Day Book General Journal"
			Exit Function
			Else
				'IF CStr(sCallTy) = "S" Then
				'	showModalDialog "DBGJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&sFromDate&"|"&sToDate&"|"&sVouType,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
				'Else
				'	showModalDialog "PrnDJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&sFromDate&"|"&sToDate&"|"&sVouType,"A","dialogHeight:200px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No"
				'End IF
			End if

'----------- Coding For the case Voucher Amount is Selected -----------

	'Elseif sFlag="Amount" Then
	If document.formname.chkBox2.checked Then

		sCheck = True
		sFlag = sFlag & "," & document.formname.chkBox2.value

		If	document.formname.txtGAmount.value="" Then
				Msgbox "Enter From Amount",0,"Day Book General Journal"
				document.formname.txtGAmount.select
				Exit Function
		Elseif document.formname.txtLAmount.value="" Then
				Msgbox "Enter To Amount",0,"Day Book General Journal"
				document.formname.txtLAmount.select
				Exit Function

		Elseif not(checkNumbers(document.formname.txtGAmount.value)) Then
				Msgbox "Enter Numbers Only",0,"Day Book General Journal"
				document.formname.txtGAmount.select
				Exit Function
		Elseif not(checkNumbers(document.formname.txtLAmount.value)) Then
				Msgbox "Enter Numbers Only",0,"Day Book General Journal"
				document.formname.txtLAmount.select
				Exit Function
		Else
			dGAmount=cdbl(document.formname.txtGAmount.value)
			dLAmount=cdbl(document.formname.txtLAmount.value)
			If cdbl(dGAmount)>cdbl(dLAmount) Then
				Msgbox "To Amount Should be Greater From Amount",0,"Day Book General Journal"
				document.formname.txtLAmount.value =""
				document.formname.txtLAmount.select
				Exit Function
			Else
			'	IF CStr(sCallTy) = "S" Then
			'		showModalDialog "DBGJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&dGAmount&"|"&dLAmount&"|"&sVouType,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
			'	Else
			'		showModalDialog "PrnDJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&dGAmount&"|"&dLAmount&"|"&sVouType,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
			'	End IF
			End if
		End if
	End IF
'------------- Coding For the case VouNo is Selected ------------------

	'Elseif sFlag ="AccHead" Then

		'if document.formname.SelAccHead.value="0" then
		'	Msgbox "Select Account Head ",0,"Day Book General Journal"
		'	document.formname.SelAccHead.focus
		'	Exit Function
		'Else
		if document.formname.SelAccHead.value="S" then
			sAccHead=document.formname.SelAccHead.options(document.formname.SelAccHead.selectedIndex).Value
			sCheck = True
			'IF CStr(sCallTy) = "S" Then
			'	showModalDialog "DBGJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&sHead&"|"&sHeadCode&"|"&sVouType&"|"&sHeadDesc,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
			'Else
			'	showModalDialog "PrnDJView.asp?Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&sHead&"|"&sHeadCode&"|"&sVouType&"|"&sHeadDesc,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
			'End IF
		End if

	'End if

	If Not sCheck Then
		alert("Select any one option to View Data")
		Exit Function
	End IF

	If sFlag <> "" Then sFlag = mid(sFlag,2)

	sPassStr="Value="&sOrgId&"|"&sOrgName&"|"&sBook&"|"&iBookNo&"|"&sFlag&"|"&iVocNoFrom&"|"&iVocNoTo&"|"&sVouType&"|"&sFromDate&"|"&sToDate&"|"&dGAmount&"|"&dLAmount&"|"&sHead&"|"&sHeadCode&"|"&sHeadDesc
	IF CStr(sCallTy) = "S" Then
		showModalDialog "DBGJView.asp?"&sPassStr,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
	Else
		'showModalDialog "../Reports/PrnDJView.asp?"&sPassStr,"A","dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No"
		showModalDialog "../Reports/PrnDJView.asp?"&sPassStr,"A","dialogHeight:150px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No"
	End IF

End Function
'---------------- Function to Display The BookHead --------------------

Function DisplayBook()
dim iUnitNo,arrTemp
dim Root
	document.formname.selBook.options.length = 1
	SelNew()
	'if objUnit.selectedIndex <> "0" then
	'	iUnitNo= objUnit(objUnit.selectedIndex).value
		iUnitNo = document.formname.hUnitId.value

		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","XMLGetOrgBook.asp?BkCode=08&orgID=" & iUnitNo , false
		objhttp.send

		if objhttp.responseXML.xml <> "" then
			UnitBookData.loadXML objhttp.responseXML.xml
			Set Root = UnitBookData.documentElement

			For Each HeaderNode In Root.childNodes
				document.formname.selBook.length = document.formname.selBook.length+1
				document.formname.selBook.options(document.formname.selBook.length-1).text = HeaderNode.Attributes.Item(1).nodeValue
				document.formname.selBook.options(document.formname.selBook.length-1).Value = HeaderNode.Attributes.Item(0).nodeValue
			next
		end if
	'end if
end Function
Function MinDate()
	sFromDate = document.formname.ctlVouFromDate.GetDate
	sToDate = document.formname.ctlVouToDate.GetDate
	sMinDate = document.formname.hFDate.value
	sMaxDate = document.formname.hTDate.value
'alert(sMinDate & sMaxDate)
	If dateDiff("d",sFromDate,document.formname.hFDate.value) > 0 or  dateDiff("d",sFromDate,document.formname.hTDate.value) < 0 then
		alert("Date Should be within the Financial Year  "& sMinDate&" to " & sMaxDate )
		document.formname.ctlVouFromDate.SetDate =	document.formname.hFDate.value
		exit Function
	end if
	If dateDiff("d",sToDate,document.formname.hFDate.value) > 0 or datediff("d",sToDate,document.formname.hTDate.value) < 0  then
		alert("Date Should be within the Financial Year  "& sMinDate &" to " & sMaxDate )
		document.formname.ctlVouToDate.SetDate = document.formname.hTDate.value
		exit Function
	end if
End Function
Function OptSelection()

	If document.formname.chkBox1.checked Then
		sFlag=document.formname.chkBox1.value
		document.formname.txtNoFrom.readOnly =false
		document.formname.txtNoTo.readOnly =false
		document.formname.txtGAmount.value =""
		document.formname.txtLAmount.value =""
	Else
		document.formname.txtNoFrom.readOnly =True
		document.formname.txtNoTo.readOnly =True
		document.formname.txtNoFrom.value =""
		document.formname.txtNoTo.value =""
	End IF

	if document.formname.chkBox2.checked Then
		sFlag=document.formname.chkBox2.value
		document.formname.txtGAmount.value =""
		document.formname.txtLAmount.value =""
		document.formname.txtGAmount.readOnly=False
		document.formname.txtLAmount.readOnly =False
	Else
		document.formname.txtGAmount.value =""
		document.formname.txtLAmount.value =""
		document.formname.txtGAmount.readOnly=True
		document.formname.txtLAmount.readOnly =True
	End If

End Function
'**********************************
Function setdate()
sFromDate = document.formname.hFDate.value
sToDate = document.formname.hTDate.value
    if DateDiff("d",sToDate,date)>0 then
        document.formname.ctlVouFromDate.setdate = sFromDate
        document.formname.ctlVouToDate.setdate = sToDate
    else
        document.formname.ctlVouFromDate.setdate = sFromDate
        document.formname.ctlVouToDate.setdate = date
    end if

End Function
</SCRIPT>
<%
dim sFinPeriod,sFinTemp,sMaxDate,sMinDate,Da,Mo,Yr,sSelDayBook
Dim sFinFromDate,sFinToDate
sFinPeriod = Session("FinPeriod")
sFinTemp = Split(sFinPeriod,":")
sFinFromDate = "01/04/"& sFinTemp(0)
sFinToDate = "31/03/"&sFinTemp(1)

sSelDayBook  = Request("RadDayBook")
If sSelDayBook = "" Then sSelDayBook = "01"
'Response.Write sMinDate & " *** "& sMaxDate
%>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onload="setdate();DisplayBook()">

<form method="POST" name="formname" action="">
	<Input type="hidden" name="hUnitId" value="<%=Session("organizationcode")%>">
	<Input type="hidden" name="hUnitName" value="<%=Session("orgshortName")%>">
	<Input type="hidden" name="hFDate" value="<%=sFinFromDate%>">
	<Input type="hidden" name="hTDate" value="<%=sFinToDate%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
		<td class=PageTitle>
			GJ Day Book
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
					<TD class="TabBodyWithTopLine">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
								</td>
								<td valign="top" align=left width="100%">
<TABLE BORDER="0" CELLSPACING=0 CELLPADDING=0>
<!--<TR><TD class="FieldCell"> Organization</TD>
<TD class="FieldCellSub">
                                                           <select size="1" name="selUnitId" class="FormElem" onChange="DisplayBook(this)">
									<OPTION value="0">Select a Unit</option>
									<%populateOrganizationList%>
                              </select>
                              </TD>
</TR>-->
<tr>
<TD class="FieldCell">
GJ Day Book</TD>
<TD class="FieldCellSub">
                                                            <select size="1" name="selBook" class="FormElem" OnChange="SelNew()">
                        <option value="S">Select Book</option>
                            </select></TD>
</tr>
</TABLE>

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
								</td>
								<td valign="top" align="center">
																<table class="BodyTable" cellSpacing="0" cellPadding="2" border="0">
																<tbody>
																<tr>
																	<td class="ExcelHeaderCell">Filter By</td>
																	<td class="ExcelHeaderCell">From&nbsp;&nbsp;</td>
																    <td class="ExcelHeaderCell">To&nbsp;&nbsp;&nbsp;&nbsp;</td>
																    <td class="ExcelHeaderCell"></td>
																</tr>
																<tr>
																    <td class="FieldCellSub">
																		<!--<input onclick="OptSelection()" type="radio" value="VouDate" name="optCriteria"  CHECKED> Voucher Date-->
																		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Voucher Date
																	</td>
																    <td align="left" class="FieldCellSub">
																	   <object id="ctlVouFromDate"  onblur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD" codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89px" height="20px" class="FormElem" viewastext>
																	   	<param name="_ExtentX" value="2355">
																	   	<param name="_ExtentY" value="529">
																	   </object>
																	</td>
																    <td align="left" class="FieldCellSub">
																	    <object id="ctlVouToDate"  onblur="MinDate()" classid="CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD" codebase="../../components/DatePicker.CAB#version=1,0,0,0" width="89px" height="20px" class="FormElem" viewastext>
																	   	<param name="_ExtentX" value="2355">
																	   	<param name="_ExtentY" value="529">
																	   </object>
																	</td>
																    <td align="left" class="FieldCellSub">&nbsp;</td>
																</tr>
																<tr>
																	<td class="FieldCellSub">
																		<!--<input type="radio" value="VouNo" name="optCriteria" onclick="OptSelection()">Voucher	Number&nbsp;-->
																		<input type="checkbox" value="VouNo" name="chkBox1" onclick="OptSelection()">Voucher	Number&nbsp;
																	</td>
																	<td align="left" class="FieldCellSub"><input class="FormElem"  size="11" name="txtNoFrom" Readonly></td>
																	<td align="left" class="FieldCellSub"><input class="FormElem"  size="11" name="txtNoTo" Readonly></td>
																	<td align="left" class="FieldCellSub">&nbsp;</td>
																</tr>
																<tr>
																   <td class="FieldCellSub">
																		<!--<input type="radio" onclick="OptSelection()"  value="Amount" name="optCriteria">-->
																		<input type="Checkbox" onclick="OptSelection()"  value="Amount" name="chkBox2">
																		Amount&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
																	<td align="left" class="FieldCellSub"><input class="FormElem" size="11" Readonly name="txtGAmount"></td>
																	<td align="left" class="FieldCellSub"><input class="FormElem" size="11" Readonly name="txtLAmount"></td>
																	<td align="left" class="FieldCellSub"></td>
																</tr>

																<tr>
																	<td class="FieldCellSub">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																		Account Head</td>
																	<td align="left" class="FieldCellSub">
																		<select class="FormElem" OnChange="SelectAccHead()" size="1" name="SelAccHead">
																		  <option value="0">Select Option</option>
																		  <option value="S" >Select Account Head</option>
																		</select>
																   </td>
																   <!--<td colSpan="2" class="FieldCellSub">
																   <span id="spAccHead" class="DataOnly"></span>&nbsp;
																   </td>-->
																</tr>
																 <tR>
																	<td colSpan="3" class="FieldCellSub">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																		<span id="spAccHead" class="DataOnly"></span>&nbsp;
																   </td>
																</tr>
																<!--<tr>
																	<td vAlign="center" class="ExcelHeaderCell" align="center">Viewed
                                                                      By</td>
																<td vAlign="center" align="center" class="ExcelHeaderCell">From&nbsp;&nbsp;</td>
                                                                  <td vAlign="center" align="center" class="ExcelHeaderCell">To&nbsp;&nbsp;&nbsp;&nbsp;</td>
                                                                  <td vAlign="center" align="center" class="ExcelHeaderCell"></td>
                                                                </tr>
																<tr>
																	<td vAlign="center" class="FieldCellSub"><input type="radio" value="VouNo" CHECKED name="optCriteria" onclick="OptSelection()">
																	Voucher	Number&nbsp;</td>
																<td vAlign="center" align="left" class="FieldCellSub" width="10"><input class="formelem"  size="11" name="txtNoFrom"></td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub" width="10"><input class="formelem"  size="11" name="txtNoTo"></td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub"></td>
                                                                </tr>
                                                                <tr>
                                                                  <td vAlign="center" class="FieldCellSub"><input onclick="OptSelection()" type="radio" value="VouDate" name="optCriteria">
                                                                    Voucher Date</td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub">
<% ' Function Call to Insert Date Picker
	Response.Write InsertDatePicker("ctlVouFromDate")
 %>
 </td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub">
<% ' Function Call to Insert Date Picker
	Response.Write InsertDatePicker("ctlVouToDate")
 %>

</td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub">

</td>
                                                                </tr>
                                                                <tr>
                                                                  <td vAlign="center" class="FieldCellSub"><input type="radio" onclick="OptSelection()"  value="Amount" name="optCriteria">
                                                                    Amount&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub"><input class="formelem" size="11" Readonly name="txtGAmount"></td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub"><input class="formelem" size="11" Readonly name="txtLAmount"></td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub"></td>
                                                                </tr>
                                                                <tr>
                                                                  <td vAlign="center" class="FieldCellSub"><input type="radio" onclick="OptSelection()" value="AccHead" name="optCriteria">
                                                                    Account Head</td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub"><select class="formelem" onchange="SelectAccHead()" disabled size="1" name="SelAccHead">
                                                                      <option value="0">Select Option</option>
                                                                      <option value="S">Selected Account Head</option>
                                                                    </select></td>
                                                                  <td vAlign="center" align="left" class="FieldCellSub" colspan="2"><span id="spAccHead" class="DataOnly"></span>&nbsp;</td>
                                                                </tr>-->
                                                              </tbody>
                                                            </table>
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
									<img border="0" src="../../assets/images/clearpixel.gif" width="5px" height="5px">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td class="ActionCell">
                                                <input type="button" value="View" class="ActionButton" onClick="CheckSubmit('S')" >
                                                <input type="button" value="Print" class="ActionButton" onClick="CheckSubmit('P')"  id=button1 name=button1>
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



