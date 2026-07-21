<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ArrangeBinDetailsEntry.asp
	'Module Name				:	Inventory (Storage Creation)
	'Author Name				:	UmaMaheswari S
	'Created On					:	May 31, 2011
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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Storage Bin Details</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
 <script LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<script LANGUAGE=javascript SRC="../../scripts/Cancel.js"></script>
<!--SCRIPT LANGUAGE=javascript SRC="../scripts/stoLocBinDetails.js"></SCRIPT-->
<SCRIPT type="text/plain" data-itms-legacy-client-script="1">
set objTemp = window.dialogarguments
'*****************************************************************************
Function checkSubmit()
	Set Root = objTemp.documentElement
	
	iLocNo = document.formname.hLocNo.value
	iItemcode = Trim(document.formname.hItemCode.value)
	iClassCode = Trim(document.formname.hClassCode.value)
	
	If cdbl(document.formname.txtTotQty.value) <> cdbl(document.formname.hTotStkQty.value) Then
		alert("Item Qty and Bin Qty is Not Equal")
		Exit Function
	End IF
	
	'sExp = "//Item[@ICode="&iItemcode&" and @CCode="& iClassCode &"]/LOCDET[@LOC="&iLocNo&"]/BINDET"
	sExp = "//Item[@ICode="&iItemcode&" and @CCode="& iClassCode &"]/LOCDET[@LOC="&iLocNo&"]/STOREBINDET/BIN"
	Set BinNode = Root.SelectNodes(sExp)
	IF BinNode.length <> 0 then
		For i = 1 to  BinNode.length
			If Eval("document.formname.Chk"&i).checked Then
				'alert(Eval("document.formname.Chk"&i).value)
			Else
				'alert(Eval("document.formname.Chk"&i).value)
				nBinNo = Eval("document.formname.Chk"&i).value
				
				sExp1 = "//Item[@ICode="&iItemcode&" and @CCode="& iClassCode &"]/LOCDET[@LOC="&iLocNo&"]/STOREBINDET"
				Set TestNode = Root.SelectNodes(sExp1)
				
				sExp1 = "//Item[@ICode="&iItemcode&" and @CCode="& iClassCode &"]/LOCDET[@LOC="&iLocNo&"]/STOREBINDET/BIN[@NO="& nBinNo &"]"
				Set DelNode = Root.SelectNodes(sExp1)
				IF DelNode.length <> 0 then
					set oNode = TestNode.Item(0).RemoveChild(DelNode.item(0))
				End IF
			End If
			BinNode.item(i-1).Attributes.getNamedItem("QTY").value = eval("document.formname.txtQty"&i).value
		Next
	End IF

	'alert "Fin="& Root.xml
	'exit function

	 'set objhttp = CreateObject("Microsoft.XMLHTTP")
	 'objhttp.Open "POST","XMLSave.asp?Name=StorageNew", false
	 'objhttp.send ObjTemp.XMLDocument
	 
	 set objHttp = CreateObject("Microsoft.XMLHTTP")
	 objHttp.open "POST","ArrBinDetInsert.asp",False
	 objHttp.send ObjTemp.XMLDocument
	 
	 If objHttp.responseText <> "" Then
		alert(objHttp.responseText)
	 Else
		'window_Unload()
		alert("Bin Details Arranged")
		window.returnvalue="Done"
	 End IF
	 window.close
End Function
'*****************************************************************************
'Function window_Unload()
	 'set window.returnValue = ObjTemp.documentElement
'End Function
'*****************************************************************************
Function DisplaytableBin()
	Dim iBinNo,arrTemp,nQty,nTotQty

	iBinNo = document.formname.hBinNo.value
	Set Root = objtemp.documentElement
	 
		ClearTable
		set oRow = document.all.tblBin.insertRow(0)

		set headerCell=oRow.insertCell()
		headerCell.innerHTML="S.No."
		headerCell.className="ExcelHeaderCell"
		headerCell.align="center"

		set headerCell=oRow.insertCell()
		headerCell.innerHTML=""
		headerCell.className="ExcelHeaderCell"
		headerCell.align="center"

		set headerCell=oRow.insertCell()
		headerCell.innerHTML="Bin Code"
		headerCell.className="ExcelHeaderCell"
		headerCell.align="center"

		set headerCell=oRow.insertCell()
		headerCell.innerHTML="Quantity"
		headerCell.className="ExcelHeaderCell"
		headerCell.align="center"

	
	
	iLocNo = document.formname.hLocNo.value
	iItemcode = Trim(document.formname.hItemCode.value)
	iClassCode = Trim(document.formname.hClassCode.value)
	
	'sExp = "//Item[@ICode="&iItemcode&" and @CCode="& iClassCode &"]/LOCDET[@LOC="&iLocNo&"]/BINDET"
	'Set Rt = Root.Selectnodes(sExp)
	'If Rt.length <> 0 then
	'	For j = 1 to Rt.length
	'		iBinNo = (Rt.item(j-1).Attributes.getNamedItem("BINNO").value)
	'		nQty   = (Rt.item(j-1).Attributes.getNamedItem("QTY").value)
			
			'for j=1 to iBinNo
	sExp = "//Item[@ICode="&iItemcode&" and @CCode="& iClassCode &"]/LOCDET[@LOC="&iLocNo&"]/STOREBINDET/BIN"
	Set binNode = Root.SelectNodes(sExp)
	If binNode.length <> 0 Then
		For j= 1 to binNode.length
			iBinNo  = binNode.item(j-1).Attributes.getNamedItem("NO").value
			nQty   = (binNode.item(j-1).Attributes.getNamedItem("QTY").value)
			sCheck = "N"
			
			sExp = "//Item[@ICode="&iItemcode&" and @CCode="& iClassCode &"]/LOCDET[@LOC="&iLocNo&"]/BINDET[@BINNO="&Trim(iBinNo)&"]"
			Set Rt = Root.Selectnodes(sExp)
			If Rt.length <> 0 then
				nQty   = Rt.item(0).Attributes.getNamedItem("QTY").value
				sCheck = "Y"
			End IF
			
			set oRow = document.all.tblBin.insertRow(j)

			set headerCell=oRow.insertCell()
			headerCell.innerHTML=j
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			'set oText = document.createElement("<input type=""Checkbox"" name=""Chk"&CStr(j)&""" Value="""&CStr(j)&"""  class=""Formelem"">" )
			If sCheck = "Y" Then
				set oText = document.createElement("<input type=""Checkbox"" name=""Chk"&CStr(j)&""" Value="""&iBinNo&"""  CHECKED class=""Formelem"">" )
			Else
				set oText = document.createElement("<input type=""Checkbox"" name=""Chk"&CStr(j)&""" Value="""&iBinNo&"""  class=""Formelem"">" )
			End IF
			headerCell.appendChild(oText)
			headerCell.className="ExcelDisplayCell"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtBinCode"&CStr(j)&""" size=""30"" Value ="""&iBinNo&""" READONLY maxlength=10 class=""Formelem"">" )
			headerCell.appendChild(oText)
			headerCell.className="ExcelInputCell"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtQty"&CStr(j)&""" size=""30"" Value ="""&nQty&""" maxlength=10 class=""Formelem"" onchange=""Calculate()"">")
			headerCell.appendChild(oText)
			headerCell.className="ExcelInputCell"
			
			nTotQty = cdbl(nTotQty) + cdbl(nQty)
			document.formname.hCnt.value = j
			
		Next
	End If
	
	nKK = document.formname.hCnt.value+1
	set oRow = document.all.tblBin.insertRow(nKK)
	set headerCell=oRow.insertCell()
	headerCell.innerHTML="Total"
	headerCell.colspan="3"
	headerCell.className="ExcelHeaderCell"
	headerCell.align="center"
	
	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text"" name=""txtTotQty"" size=""30""  value="""& nTotQty &""" maxlength=10 READONLY class=""Formelem"">")
	headerCell.appendChild(oText)
	headerCell.className="ExcelInputCell"
			
	
end Function
'*****************************************************************************
Function Calculate()
	Dim nNoOfRows,iCtr
	nNoOfRows = document.formname.hCnt.value
	For iCtr = 1 To nNoOfRows
		If Eval("document.formname.txtQty"&iCtr).value < 0 Then
			alert("Enter Valid value")
			Exit Function
		Elseif Not checkNumbers(Eval("document.formname.txtQty"&iCtr).value) Then
			alert("Enter Numerals only")
			Exit Function
		End IF
		nTotQty = cdbl(nTotQty) + cdbl(Eval("document.formname.txtQty"&iCtr).value)
	Next
	document.formname.txtTotQty.value = nTotQty
	
End Function
'*****************************************************************************
Function checkNumbers(val)
	dim valid,temp,i
	valid = "0123456789."
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
	for i=0 to document.all.tblBin.rows.length - 1
		document.all.tblBin.deleteRow(0)
	next
end Function
'*****************************************************************************
Function ClearAll()
	document.formname.selLocName.options.length = 0
	document.formname.selLocName.length = document.formname.selLocName.length+1
	document.formname.selLocName.options(document.formname.selLocName.length-1).text = "Select"
	document.formname.selLocName.options(document.formname.selLocName.length-1).Value = "select"
	ClearTable
end Function
'*****************************************************************************
Function FnInit()
	Set Root = objTemp.documentElement
	'alert("Init="&Root.xml)
	DisplaytableBin()
End Function
'*****************************************************************************
</SCRIPT>
<%
Dim sData,sTempArr,nItemCode,nClassCode,sItemDesc,sStoreName,nTotQty,sLocNo,sBinNo
sData = Request.QueryString("Data")
sTempArr = Split(sData,":")
'Response.Write "<p><Font color=red>data="&sData

nItemCode = sTempArr(0)
nClassCode = sTempArr(1)
sItemDesc = sTempArr(2)
sStoreName = sTempArr(3)
nTotQty = sTempArr(4)
sLocNo = sTempArr(5)


%>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" onLoad="FnInit()">

<form method="POST" name="formname" action="">
<input type="hidden" name="hBinNo" value="<%=sBinNo%>">
<input type="hidden" name="hLocNo" value="<%=sLocNo%>">
<input type="hidden" name="hItemCode" value="<%=nItemCode%>">
<input type="hidden" name="hClassCode" value="<%=nClassCode%>">
<input type="hidden" name="hCnt" value="0">
<input type="hidden" name="hTotStkQty" value="<%=nTotQty%>">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Storage Location Creation
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
					<td height="20" valign="bottom">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" width="100%">

									</table>
								</td>
								<td >
									<table border="0" cellpadding="0" cellspacing="0" width="100%" >

									</table>
								</td>
								<td class="TabCellEnd" valign="bottom" align="left">
                                    <p align="center"><font face="Verdana" size="1" color="#FFFFFF">&nbsp;</font>
								</td>
							</tr>
						</table>
					</td>
				</tr >
				<TR>
					<TD class=TabBody>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" width="100%">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td class=FieldCell> Item Description</td>
											<td class='FieldCellSub'><span id="ItemDesc" class="Dataonly"><%=sItemDesc%></span>
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell> Store Name</td>
											<td class='FieldCellSub'>
											<span id="StoreName" class="Dataonly"><%=sStoreName%></span>
                                            </td>
										</tr>
										<tr>
											<td class=FieldCell> Total Qty</td>
											<td class='FieldCellSub'>
											<span id="TotQty" class="Dataonly"><%=nTotQty%></span>
                                            </td>
										</tr>
									</table>
                                    </div>
								</td>
								<td align="center" width="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5">
								</td>
								<td valign="top" class="MiddlePack">
                                    <table border="0" cellspacing="1" Id ="tblBin" name="tblBin" class="ExcelTable" width="400"></table>
								</td>
								<td align="center" width="5">
								</td>
							</tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" width="5">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
												<p align="center">
													<!--<input type="button" value="Add New" name="B4" class="ActionButtonX" onClick="AddNew()">-->
                                                    <input type="button" value="Done" name="B2" class="ActionButton" onClick="CheckSubmit()">
													<input type="button" value="Close" name="B3" class="ActionButton" onClick="window.close()">
											</td>
										</tr>
									</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center" width="5" colspan="3" class="BottomPack">
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
Function DisplaytableBinOLD()
	Dim iBinNo,arrTemp,nQty,nTotQty

	iBinNo = document.formname.hBinNo.value
	Set Root = objtemp.documentElement
	 
		ClearTable
		set oRow = document.all.tblBin.insertRow(0)

		set headerCell=oRow.insertCell()
		headerCell.innerHTML="S.No."
		headerCell.className="ExcelHeaderCell"
		headerCell.align="center"

		set headerCell=oRow.insertCell()
		headerCell.innerHTML=""
		headerCell.className="ExcelHeaderCell"
		headerCell.align="center"

		set headerCell=oRow.insertCell()
		headerCell.innerHTML="Bin Code"
		headerCell.className="ExcelHeaderCell"
		headerCell.align="center"

		set headerCell=oRow.insertCell()
		headerCell.innerHTML="Quantity"
		headerCell.className="ExcelHeaderCell"
		headerCell.align="center"

	
	
	iLocNo = document.formname.hLocNo.value
	iItemcode = Trim(document.formname.hItemCode.value)
	iClassCode = Trim(document.formname.hClassCode.value)
	
	sExp = "//Item[@ICode="&iItemcode&" and @CCode="& iClassCode &"]/LOCDET[@LOC="&iLocNo&"]/BINDET"
	Set Rt = Root.Selectnodes(sExp)
	If Rt.length <> 0 then
		For j = 1 to Rt.length
			iBinNo = (Rt.item(j-1).Attributes.getNamedItem("BINNO").value)
			nQty   = (Rt.item(j-1).Attributes.getNamedItem("QTY").value)
			
			'for j=1 to iBinNo
			sExp = "//Item[@ICode="&iItemcode&" and @CCode="& iClassCode &"]/LOCDET[@LOC="&iLocNo&"]/STOREBINDET/BIN"
			Set binNode = Root.SelectNodes(sExp)
			If binNode.length <> 0 Then
				For k = 1 to binNode.length
					nStoreBinNo  = binNode.item(k-1).Attributes.getNamedItem("NO").value
					
				Next
			End IF
			
			set oRow = document.all.tblBin.insertRow(j)

			set headerCell=oRow.insertCell()
			headerCell.innerHTML=j
			headerCell.className="ExcelDisplayCell"
			headerCell.align="center"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""Checkbox"" name=""Chk"&CStr(j)&""" Value="""&CStr(j)&"""  class=""Formelem"">" )
			headerCell.appendChild(oText)
			headerCell.className="ExcelDisplayCell"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtBinCode"&CStr(j)&""" size=""30"" Value ="""&iBinNo&""" READONLY maxlength=10 class=""Formelem"">" )
			headerCell.appendChild(oText)
			headerCell.className="ExcelInputCell"

			set headerCell=oRow.insertCell()
			set oText = document.createElement("<input type=""text"" name=""txtQty"&CStr(j)&""" size=""30"" Value ="""&nQty&""" maxlength=10 class=""Formelem"" onchange=""Calculate()"">")
			headerCell.appendChild(oText)
			headerCell.className="ExcelInputCell"
			
			nTotQty = cdbl(nTotQty) + cdbl(nQty)
			document.formname.hCnt.value = j
			
		Next
	End If
	
	nKK = document.formname.hCnt.value+1
	set oRow = document.all.tblBin.insertRow(nKK)
	set headerCell=oRow.insertCell()
	headerCell.innerHTML="Total"
	headerCell.colspan="3"
	headerCell.className="ExcelHeaderCell"
	headerCell.align="center"
	
	set headerCell=oRow.insertCell()
	set oText = document.createElement("<input type=""text"" name=""txtTotQty"" size=""30""  value="""& nTotQty &""" maxlength=10 class=""Formelem"">")
	headerCell.appendChild(oText)
	headerCell.className="ExcelInputCell"
			
	
end Function
'*****************************************************************************
%>