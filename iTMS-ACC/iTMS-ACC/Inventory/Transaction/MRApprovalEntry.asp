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
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
dim j,iCounter1,ssupcode,iser,iTempCode,Partyname
j = 0
iser = 0
iCounter1 = 0
iTempCode = 0
Function Init(dDate)

	IF DateValue(dDate) <  DAteValue(document.formname.hToDate.value) or   DateValue(Date()) > dateValue(document.formname.hFrmDate.value) then
		document.formname.ctlCDDate.SetDate = dDate
	Else
	  document.formname.ctlCDDate.SetDate = document.formname.hToDate.value
	End IF
	
	sIssToType = document.formname.hIssueToType.value
	sIssToCode = document.formname.hIssueToCode.value
	sIssToSubCode = document.formname.hIssueToSubCode.value
	
	for iCnt = 0 to cint(document.formname.selIssueTo.length-1)
	    if lcase(sIssToType)="party" then
	        if lcase(document.formname.selIssueTo(iCnt).value)=lcase(sIssToType) then
	            document.formname.selIssueTo.selectedIndex = iCnt
	        end if
	    else
	        if lcase(document.formname.selIssueTo(iCnt).value)=lcase(sIssToType)&":"&lcase(sIssToCode) then
	            document.formname.selIssueTo.selectedIndex = iCnt
	        end if
	    end if
	next
	
	
end Function
'********************************************************************************************
Function MinDate()

  	Dim sMinDate,sFinPeriod,sSelDate,sMaxDate
  	'alert("date check")
  	'sFinPeriod = document.formname.hFinPeriod.value
  	sMinDate = document.formname.hFrmDate.value
  	sMaxDate = document.formname.hToDate.value
  	dDate = document.formname.ctlCDDate.getdate
  	dMrsDate = document.formname.hMRDate.value
  	'alert(RngFrom &"="& sMinDate)
  	If DateValue(dDate) < DateValue(sMinDate) or  DateValue(dDate) > DateValue(sMaxDate) then
  		Alert("Date Should be With in the Range "& sMinDate & " to " & sMaxDate)
  		document.formname.ctlCDDate.Setdate = dMrsDate
  		Exit function
  	End If

End Function
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

Function ClearTable()
	dim i
	for i=1 to document.all.tblLot.rows.length - 1
		document.all.tblLot.deleteRow(1)
	next
	j = 0
end Function

Function GetItems(todaysDate)
'	if document.formname.selUnit.selectedIndex = "0" then
'		alert("Select Unit")
'		document.formname.selUnit.focus
'		exit function
'	else
		sorgID = document.formname.hUnit.value

		OutValue = showModalDialog("ItemSelectPop.asp?orgID=" & sorgID,"","dialogHeight:600px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
		while UBound(arrTemp) = 0
			OutValue = showModalDialog("ItemSelectPop.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No")
			arrTemp = split(OutValue,":")
		wend

		if UBound(arrTemp) = 1 then exit function

		Set objhttp = CreateObject("Microsoft.XMLHTTP")
		OutValue = OutValue & "``" & sorgID
		objhttp.Open "GET","XMLGetItemDetails.asp?sDet="&OutValue,false
		objhttp.send
		'alert(objhttp.responseText)
		saTemp = split(OutValue,",")
		If objhttp.responseXML.xml <> "" then
			'ClearTable()
			ItemData.loadXML objhttp.responseXML.xml
			set rootData = ItemData.DocumentElement
			set root = OutData.DocumentElement
			if rootData.hasChildNodes then
				For each ndTemp in rootData.childNodes
					if root.hasChildNodes then
						sExp ="//ITEMDETAILS [ @ITEMCODE = "&ndTemp.attributes.getNamedItem("ITEMCODE").value&" and @CLASSCODE = "&ndTemp.attributes.getNamedItem("CLASSCODE").value&"]"
						Set CheckNode = Root.Selectnodes(sExp)
						if CheckNode.Length = 0 then
							Set newElem = OutData.createElement("ITEMDETAILS")
							newElem.setAttribute "ITEMCODE", ndTemp.attributes.getNamedItem("ITEMCODE").value
							newElem.setAttribute "CLASSCODE", ndTemp.attributes.getNamedItem("CLASSCODE").value
							newElem.setAttribute "UNIT", ndTemp.attributes.getNamedItem("UNIT").value
							newElem.setAttribute "ITEMNAME", ndTemp.attributes.getNamedItem("ITEMNAME").value
							newElem.setAttribute "UOM", ndTemp.attributes.getNamedItem("UOM").value
							newElem.setAttribute "DECIMAL", ndTemp.attributes.getNamedItem("DECIMAL").value
							newElem.setAttribute "DISPLAYED", "N"
							newElem.setAttribute "QTY", ""
							newElem.setAttribute "REQUIREDBY", ""
							newElem.setAttribute "REQUIREDVALUE", ""

							Root.appendChild newElem
						end if
					else
						Set newElem = OutData.createElement("ITEMDETAILS")
						newElem.setAttribute "ITEMCODE", ndTemp.attributes.getNamedItem("ITEMCODE").value
						newElem.setAttribute "CLASSCODE", ndTemp.attributes.getNamedItem("CLASSCODE").value
						newElem.setAttribute "UNIT", ndTemp.attributes.getNamedItem("UNIT").value
						newElem.setAttribute "ITEMNAME", ndTemp.attributes.getNamedItem("ITEMNAME").value
						newElem.setAttribute "UOM", ndTemp.attributes.getNamedItem("UOM").value
						newElem.setAttribute "DECIMAL", ndTemp.attributes.getNamedItem("DECIMAL").value
						newElem.setAttribute "DISPLAYED", "N"
						newElem.setAttribute "QTY", ""
						newElem.setAttribute "REQUIREDBY", ""
						newElem.setAttribute "REQUIREDVALUE", ""

						Root.appendChild newElem
					end if
				next
			end if
		end if
		'alert(OutData.xml)
		DisplayTable todaysDate
	'end if
End Function

Function DeleteItems()
	set root = OutData.DocumentElement

	sExp ="//ITEMDETAILS"
	Set ItemNode = Root.Selectnodes(sExp)
	if ItemNode.Length > 0 then
		for itr = 0 to ItemNode.Length - 1
			iItemDel = ItemNode.Item(itr).Attributes.getNamedItem("ITEMCODE").value
			iClassDel = ItemNode.Item(itr).Attributes.getNamedItem("CLASSCODE").value
			iEntNo = ItemNode.Item(itr).Attributes.getNamedItem("ENTRYNO").value
			set objSel = eval("document.formname.chkDeleteA"&CStr(iItemDel)&"A"&iClassDel&"A"&iEntNo)
			if objSel.checked then
				ItemNode.Item(itr).Attributes.getNamedItem("DISPLAYED").value = "Y"
			else
				ItemNode.Item(itr).Attributes.getNamedItem("DISPLAYED").value = "N"
			end if
		next
	end if
End Function

Function GetAddDetails(sItem,sClass,sOrg,iEntNo,iAttrList)
	sUsage = document.formname.selUsage.value
	set Q = eval("document.formname.txtQtyZ"&sItem&"Z"&sClass&"Z"&iEntNo)

	if trim(Q.value) = "" then
		alert("Enter Quantity")
		exit function
	end if

	if cdbl(Q.value) = 0 then exit function

	sTempValues = sClass&"|"&sItem&"|"&sOrg&"|"&sUsage&"|"&trim(Q.value)&"|"&iEntNo&"|"&iAttrList
	'alert(sTempValues)

	if sUsage = "PRD" then
		set OutValue = showModalDialog("DirectIssueMixEntry.asp?sTemp="&sTempValues,OutData,"dialogHeight:300px;dialogWidth:300px;center:Yes;help:No;resizable:No;status:No")
	elseif sUsage = "PAC" then
		set OutValue = showModalDialog("DirectIssuePackingEntry.asp?sTemp="&sTempValues,OutData,"dialogHeight:400px;dialogWidth:325px;center:Yes;help:No;resizable:No;status:No")
	elseif sUsage = "WIP" or sUsage = "MAT" then
		'set OutValue = showModalDialog("DirectIssueAddEntry.asp?sTemp="&sTempValues,OutData,"dialogHeight:370px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
		set OutValue = showModalDialog("AddEntryDetails.asp?sTemp="&sTempValues,OutData,"dialogHeight:370px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No")
	end if
end Function

Function DisplayStock(sItem,sClass,sOrg,iEntNo,sItemName,sAttList)
	showModalDialog "../master/itmStockDetailsPop.asp?sItem="&sItem&"&sClass="&sClass&"&EntNo="&iEntNo&"&ItemName="&sItemName&"&sOrg="&sOrg,"Stock","dialogHeight:400px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No"
end Function

Function CheckSch(obj,todaysdate,sUoM,sAttribList)
	Set Root = OutData.documentElement
	'alert(OutData.xml)
	dim sItem,sClass,a
	arrTemp = split(obj.name,"Z")

	sItem = arrTemp(1)
	sClass = arrTemp(2)
	iEntNo = arrTemp(3)
	iAttribList = sAttribList

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	if trim(iAttribList) <> "" then
	 	'sOptName = FunAttribName(sAttributeList)
	 	objhttp.Open "GET","XMLGetAttributeName.asp?Para="&iAttribList,false
		objhttp.send
		sOptName = objhttp.responsetext
	else
		sOptName = ""
	end if
	For Each HeaderNode In Root.childNodes

		if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem  and HeaderNode.Attributes.Item(2).nodeValue = sClass then
			if HeaderNode.HaschildNodes() then
				For Each HNode In HeaderNode.childNodes
					if StrComp(Trim(HNode.NodeName),"Schedule") = 0 or StrComp(Trim(HNode.NodeName),"ScheduleDetails") = 0 then
						set a = HeaderNode.removeChild(HNode)
					end if
				next
			end if
		end if
	Next

	if (obj.selectedIndex = "1") then
		Set Root = OutData.documentElement
		iSchEntNo = 1
		For Each HeaderNode In Root.childNodes
			if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem  and HeaderNode.Attributes.Item(2).nodeValue = sClass then
				Set newElem = OutData.createElement("Schedule")
				newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
				newElem.setAttribute "SVALUE", todaysdate
				newElem.setAttribute "ITEMCODE", sItem
				newElem.setAttribute "CLASSCODE", sClass
				newElem.setAttribute "SCHENTRYNO",iSchEntNo
				HeaderNode.appendChild newElem
			end if
			iSchEntNo = iSchEntNo + 1
		Next
	end if
	if (obj.selectedIndex = "2") then
		value = prompt("Enter No of Days","0")
		if (isNull(value)) then
			obj.selectedIndex=0
			exit function
		elseif (trim(value)="") then
			obj.selectedIndex=0
			exit function
		else
			if(trim(value)="") then
				msgbox "Enter Number of Days",0,"Number of Days"
				obj.selectedIndex=0
				exit function
			else
				if(not checkNumbers(value)) then
					msgbox "Enter Numerals Only",0,"Numerals"
					obj.selectedIndex=0
					exit function
				else
				iSchEntNo = 1
					Set Root = OutData.documentElement
					For Each HeaderNode In Root.childNodes
						if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem  and HeaderNode.Attributes.Item(2).nodeValue = sClass then
							Set newElem = OutData.createElement("Schedule")
							newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
							newElem.setAttribute "SVALUE", trim(value)
							newElem.setAttribute "ITEMCODE", sItem
							newElem.setAttribute "CLASSCODE", sClass
							newElem.setAttribute "SCHENTRYNO",iSchEntNo
							HeaderNode.appendChild newElem
						end if
						iSchEntNo = iSchEntNo + 1
					Next
				end if
			end if
		end if
	end if
	if (obj.selectedIndex = "3") then
		value=prompt("Enter the Date","")
		if (isNull(value)) then
			obj.selectedIndex=0
			exit function
		elseif (trim(value)="") then
			objType.selectedIndex=0
			objValue.value=	""
			exit function
		else
			if (not vd(value,todaysdate)) then
				MsgBox "Invalid Date",0,"Invalid Date"
				obj.selectedIndex=0
				Exit Function
			end if
			if (DateDiff("d",todaysdate,value) < 0) then
				MsgBox "Date should be greater or equal to Today's Date",0,"Invalid Date"
				obj.selectedIndex=0
				Exit Function
			else
				iSchEntNo = 1
				Set Root = OutData.documentElement
				For Each HeaderNode In Root.childNodes
					if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem  and HeaderNode.Attributes.Item(2).nodeValue = sClass then
						Set newElem = OutData.createElement("Schedule")
						newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
						newElem.setAttribute "SVALUE", trim(value)
						newElem.setAttribute "ITEMCODE", sItem
						newElem.setAttribute "CLASSCODE", sClass
						newElem.setAttribute "SCHENTRYNO",	iSchEntNo
						HeaderNode.appendChild newElem
					end if
					iSchEntNo =iSchEntNo + 1
				Next
			end if
		end if
	end if
	if (obj.selectedIndex = "4") then
		dim qty
		set qty = eval("document.formname.txtQtyZ"+cstr(sItem)+"Z"+cstr(sClass)+"Z"+iEntNo)
		if (trim(qty.value)="") then
			MsgBox "Enter Quantity",0,"Quantity"
			qty.focus()
			obj.selectedIndex=0
			exit function
		elseif(not checkNumbers(qty.value)) then
			msgbox "Enter Numerals Only",0,"Numerals"
			qty.focus()
			obj.selectedIndex=0
			exit function
		else
			iSchEntNo = 1
			Set Root = OutData.documentElement
			For Each HeaderNode In Root.childNodes
				if HeaderNode.Attributes.Item(0).nodeValue = iEntNo and HeaderNode.Attributes.Item(1).nodeValue = sItem  and HeaderNode.Attributes.Item(2).nodeValue = sClass then
					Set newElem = OutData.createElement("Schedule")
					newElem.setAttribute "STYPE", trim(obj(obj.selectedIndex).value)
					newElem.setAttribute "SVALUE", ""
					newElem.setAttribute "ITEMCODE", sItem
					newElem.setAttribute "CLASSCODE", sClass
					newElem.setAttribute "SCHENTRYNO",	iSchEntNo
					HeaderNode.appendChild newElem
				end if
				iSchEntNo = iSchEntNo + 1
			Next
			'
			sTempValues = qty.value&":"&sItem&":"&sClass&":"&iEntNo&":"&document.formname.hUnit.value&":"&sUoM&":"&sOptName

			Set OutDataValue = showModalDialog("MRGenSchedulePoP.asp?sTemp="&sTempValues,OutData,"dialogHeight:510px;dialogWidth:375px;center:Yes;help:No;resizable:No;status:No")
			'alert("Chk="&OutData.xml)
		end if
	end if
end Function

Function CheckSubmit(todaysdate)

sRequestTo = document.formname.selIssueTo(document.formname.selIssueTo.selectedIndex).value
sRequestFor =document.formname.cmbIssType(document.formname.cmbIssType.selectedIndex).value

    if document.formname.selIssueTo.selectedIndex<=0 then
        alert("Select Requested By")
        document.formname.selIssueTo.focus
        exit function
    elseif trim(sRequestFor)="SEL" then
        alert("Select Requested For")
        document.formname.cmbIssType.focus
        exit function
    elseif(datediff("d",todaysdate,document.formname.ctlCDDate.GetDate)) > 0 then
		alert("Created On should be less than or equal to Today's Date")
		exit function
	elseif len(trim(document.formname.txtRemarks.value)) > 200 then
		alert("Remarks should be less than 200 characters")
		document.formname.txtRemarks.select
		exit function
	else
		itr = 0
		sAddSpcsCount = 0
		set root = OutData.DocumentElement
		if root.haschildnodes then
			for each node in root.childnodes
				if trim(node.NodeName) = trim("HEADER") then
					root.removechild node
				end if
			next

		end if

		if root.hasChildNodes then
			sExp ="//ITEMDETAILS"
			Set ItemNode = Root.Selectnodes(sExp)
			for itr = 0 to ItemNode.Length - 1
				iEntNo = ItemNode.Item(itr).Attributes.getNamedItem("ENTRYNO").value
				iItem = ItemNode.Item(itr).Attributes.getNamedItem("ITEMCODE").value
				iClass = ItemNode.Item(itr).Attributes.getNamedItem("CLASSCODE").value

				Set objQty = eval("document.formname.txtQtyZ"&iItem&"Z"&iClass&"Z"&iEntNo)
				Set objReq = eval("document.formname.selSchZ"&iItem&"Z"&iClass&"Z"&iEntNo)

				if objQty.value = "" or objQty.value = "0" then
					alert("Enter Quantity")
					objQty.select
					exit function
				elseif objReq.selectedIndex = "0" then
					alert("Select Required By")
					objReq.focus
					exit function
				end if
				ItemNode.Item(itr).Attributes.getNamedItem("QTY").value = objQty.value
				ItemNode.Item(itr).Attributes.getNamedItem("REQUIREDBY").value = objReq.value




				sExp1 ="//ITEMDETAILS [ @CLASSCODE = "&iClass&" and @ITEMCODE = "&iItem&" and @ENTRYNO = "&iEntNo&"]/AddDet"
				Set ADNode = Root.Selectnodes(sExp1)
				if ADNode.Length <= 0 then
				    sAddSpcsCount = cdbl(sAddSpcsCount) + 1
				end if
			next
		end if
	'	if sAddSpcsCount > 0 then
	'	    if Not confirm("Add.Spec Details are not Specified for all Items. Do you want to Save?") then
	'		    exit function
	'	    end if
	'	end if ' if sAddSpcsCount > 0 then

		Set newElem = OutData.createElement("HEADER")
		newElem.setAttribute "FORUNIT", document.formname.hUnit.value
		newElem.setAttribute "CREATEDON", document.formname.ctlCDDate.GetDate
		newElem.setAttribute "TYPE",""' document.formname.selReqType.value
		newElem.setAttribute "USAGE", ""'document.formname.selUsage.value
		newElem.setAttribute "REMARKS", trim(document.formname.txtRemarks.value)
		newElem.setAttribute "CREATEDBY", document.formname.hCreatedBy.value
		newElem.setAttribute "MRNO", document.formname.hMRNo.value
		newElem.setAttribute "LOTCARDNO", ""
		newElem.setAttribute "MACHINENO",""
		newElem.setAttribute "COSTCENTER", trim(document.formname.selCC.value)
		newElem.setAttribute "REFTYPE",""
		newElem.setAttribute "ISSTOTYPE",document.formname.hIssueToType.value
        newElem.setAttribute "ISSTOCODE",document.formname.hIssueToCode.value
        newElem.setAttribute "ISSTOSUBCODE",document.formname.hIssueToSubCode.value
        newElem.setAttribute "ISSUETYPECODE",sRequestFor
		Root.appendChild newElem
	end if
	DeleteItems

	'alert(OutData.xml)
	'exit function
	sAction = document.formname.hAction.value
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "POST","MRApprovalInsert.asp?Action="& sAction  , false
	objhttp.send OutData.XMLDocument
	'alert(objhttp.responseText)
	'exit function

	if objhttp.responseText = "" then
	    if trim(sAction)="Amend" then
	        sMessage =  "Material Requisition has been Updated. Do you want to approve another one?"
	    elseif trim(sAction)="Approve" then
	        sMessage =  "Material Requisition has been Approved / Rejected. Do you want to approve another one?"
	    elseif trim(sAction)="Cancel" then
	        sMessage =  "Material Requisition has been Cancelled. Do you want to approve another one?"
	    end if
		if confirm(sMessage) then
			window.location.href "MRSMGMTLIST.ASP"
		else
			window.location.href "../welcome_Inventory.asp"
		end if
	else
		alert(objhttp.responseText)
	end if

end Function



</SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="/Scripts/itms-modern-compat.js"></SCRIPT>
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
