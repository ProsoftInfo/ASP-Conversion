<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/ReportsBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<XML id="OutData"><Root/></xml>
<XML id="TempData"><Root/></XML>

</head>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" Onload="Loadvalues()">
<%
	Dim GroupCode,iHeadID,OutValue,splt
	GroupCode = Request("GroupCode")
	'splt = Split(OutValue,":")
	'GroupCode = CInt(splt(0))
	'HeadID = CInt(splt(1))  
	iHeadID = Request("HeadID")
	'Response.Write iHeadID
	'HeadID = 3	
	
Response.Write OutValue
%>
<Script Language = VBScript >
Function CheckBoxClick()
	'MsgBox "Ok" 
End Function 

Function Loadvalues()
	Dim objhttp,GroupCode,id,iHeadID,Merged,iTDSSubHeads,recno,sExp,Root,i
	Dim Formula,splt,splt1,splt2,Voucher,iGroupHead,iCount,iCheck,iSum
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	GroupCode= document.formname.GroupCode.value  
	iHeadID = document.formname.HeadID.value  
	'MsgBox iHeadID
	iSum = 0
	iCheck = 0
	id = "2"
	objhttp.Open "GET","TDSXMLGenerate.asp?GroupCode="&GroupCode&"&HeadID="&iHeadID&"&id="&id, false
	objhttp.send
			'MsgBox objhttp.responseText	
			OutData.loadXML objhttp.responseXML.xml	
			Set Root = OutData.documentELement
			sExp = "//iXML"
			Set HeadNode = Root.selectNodes(sExp)
			recno = HeadNode.length-1
		'	MsgBox OutData.XML
			For Each HeadNode in Root.ChildNodes
				Formula = HeadNode.Attributes.Item(5).nodeValue	
			Next
			'MsgBox recno 
			If recno >= 0 Then	
				If Formula <>"" Then
					splt= Split(Formula,",")
					splt1 = Split(splt(0),"-")  
					Voucher = splt1(0)
					If Voucher="0#0" Then
						Formula = Replace(Formula,splt(0)&",","")  
						splt= Split(Formula,",")
						splt1 = Split(splt(0),"-") 
					End If
					If UBound(splt1) > 0 Then
					document.formname.txtpercentage.value = splt1(1)  
					End If
				End If
				
				'Add Voucher
				set oRow = document.all.tbltds.insertRow(document.all.tbltds.rows.length)
				iCount = document.all.tbltds.rows.length
				set headerCell=oRow.insertCell()				
				headerCell.innerHTML=  Cint(document.formname.iRowCount.value)   
				headerCell.className="ExcelHeaderCell"
				headerCell.align="center"
				set headerCell=oRow.insertCell()
				'msgbox Voucher				
				If Voucher = "0#0" Then 
					set oText = document.createElement("<input type=""checkbox""  name=""chkDel"" class=""ExcelDisplayCell"" Value=""0#0"" Checked >"  )
				Else
					set oText = document.createElement("<input type=""checkbox""  name=""chkDel"" class=""ExcelDisplayCell"" Value=""0#0"" >"  )
				End If
				headerCell.appendChild(oText)
				headerCell.className="ExcelDisplayCell"
				headerCell.align="center"
'					
				set headerCell=oRow.insertCell()				
				headerCell.innerHTML = "Voucher Details"
				headerCell.className="ExcelDisplayCell"
				headerCell.align="left"
				document.formname.iRowCount.value = Cint(document.formname.iRowCount.value) + 1    
			
				For Each HeadNode in Root.ChildNodes
					
					If Formula <> "" Then	
						iCheck = 0 
						splt= Split(Formula,",")
						If ubound(splt)>=iSum Then
						splt1 = Split(splt(iSum),"-") 
						splt2 = Split(splt1(0),"#")
						iCheck = splt2(1)
						End If
					End If
					
					document.formname.txtGname.value = HeadNode.Attributes.Item(6).nodeValue
					document.formname.txtcomputationfor.value = HeadNode.Attributes.Item(7).nodeValue  
					set oRow = document.all.tbltds.insertRow(document.all.tbltds.rows.length)
					iCount = document.all.tbltds.rows.length
					set headerCell=oRow.insertCell()				
					headerCell.innerHTML=  Cint(document.formname.iRowCount.value)   
					headerCell.className="ExcelHeaderCell"
					headerCell.align="center"
					set headerCell=oRow.insertCell()
					iTDSSubHeads = HeadNode.Attributes.Item(0).nodeValue
					Merged = iHeadID&"#"&iTDSSubHeads 
					'MsgBox Merged
					If iCheck <> 0 Then
						set oText = document.createElement("<input type=""checkbox""  name=""chkDel"" class=""ExcelDisplayCell"" Value="&Merged&" Onclick=""CheckBoxClick()"" Checked >" )
					Else
						set oText = document.createElement("<input type=""checkbox""  name=""chkDel"" class=""ExcelDisplayCell"" Value="&Merged&" Onclick=""CheckBoxClick()"" >" )
					End If
					headerCell.appendChild(oText)
					headerCell.className="ExcelDisplayCell"
					headerCell.align="center"
					
					set headerCell=oRow.insertCell()				
					headerCell.innerHTML = HeadNode.Attributes.Item(1).nodeValue  
					headerCell.className="ExcelDisplayCell"
					headerCell.align="left"
					document.formname.iRowCount.value = Cint(document.formname.iRowCount.value) + 1    
					iSum = iSum + 1
			Next
			
			Else
				If Formula <>"" Then
					splt= Split(Formula,",")
					splt1 = Split(splt(0),"-")  
					Voucher = splt1(0)
					If Voucher="0#0" Then
						Formula = Replace(Formula,splt(0)&",","")  
						splt= Split(Formula,",")
						splt1 = Split(splt(0),"-") 
					End If
					If UBound(splt1) > 0 Then
					document.formname.txtpercentage.value = splt1(1)  
					End If
				End If
					For Each HeadNode in Root.ChildNodes			
						document.formname.txtGname.value = HeadNode.Attributes.Item(6).nodeValue
						document.formname.txtcomputationfor.value = HeadNode.Attributes.Item(7).nodeValue  
						set oRow = document.all.tbltds.insertRow(document.all.tbltds.rows.length)
						iCount = document.all.tbltds.rows.length
						set headerCell=oRow.insertCell()				
						headerCell.innerHTML=  Cint(document.formname.iRowCount.value)   
						headerCell.className="ExcelHeaderCell"
						headerCell.align="center"
						set headerCell=oRow.insertCell()	
						If Voucher = "0#0" Then			
							set oText = document.createElement("<input type=""checkbox""  name=""chkDel"" class=""ExcelDisplayCell"" Value=""0#0"" Checked >"  )
						Else
							set oText = document.createElement("<input type=""checkbox""  name=""chkDel"" class=""ExcelDisplayCell"" Value=""0#0"" >"  )
						End If
						headerCell.appendChild(oText)
						headerCell.className="ExcelDisplayCell"
						headerCell.align="center"
	'					
						set headerCell=oRow.insertCell()				
						headerCell.innerHTML = "Voucher Details"
						headerCell.className="ExcelDisplayCell"
						headerCell.align="left"
						document.formname.iRowCount.value = Cint(document.formname.iRowCount.value) + 1    
					Next
	End IF
End Function

Function Updatexml()
	Dim GroupCode,iHeadID
	Dim iCtr,ComputeFormula
	Dim objhttp,Node1
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	Set Root = TempData.documentElement
	GroupCode= document.formname.GroupCode.value  
	iHeadID = document.formname.HeadID.value  
	If document.all.tbltds.rows.length - 2 = 0 Then
		For iCtr = 0 To document.all.tbltds.rows.length - 2
			IF document.formname.chkDel.Checked = True Then
					If ComputeFormula <>"" Then
						If document.formname.txtpercentage.value <>"" then  
							ComputeFormula = ComputeFormula&","&document.formname.chkDel.Value&"-"&document.formname.txtpercentage.value 
						Else
							ComputeFormula = ComputeFormula&","&document.formname.chkDel.Value
						End If
					Else
						If document.formname.txtpercentage.value <>"" then  
							ComputeFormula = document.formname.chkDel.Value&"-"&document.formname.txtpercentage.value 
						Else
							ComputeFormula = document.formname.chkDel.Value
						End If
					End If
			End IF
		Next
	Else
		For iCtr = 0 To document.all.tbltds.rows.length - 2
			IF document.formname.chkDel(iCtr).Checked = True Then
					If ComputeFormula <>"" Then
						If document.formname.txtpercentage.value <>"" then  
						ComputeFormula = ComputeFormula&","&document.formname.chkDel(iCtr).Value&"-"&document.formname.txtpercentage.value 
						Else
						ComputeFormula = ComputeFormula&","&document.formname.chkDel(iCtr).Value
						End If
					Else
						If document.formname.txtpercentage.value <>"" then  
						ComputeFormula = document.formname.chkDel(iCtr).Value&"-"&document.formname.txtpercentage.value 
						Else
						ComputeFormula = document.formname.chkDel(iCtr).Value
						End If
					End If
			End IF
		Next
	End If
			'MsgBox GroupCode&iHeadID&ComputeFormula 
			Set Root = TempData.documentElement 
			set Node1 = TempData.CreateElement("Schedule")
			Node1.setAttribute "GroupID",GroupCode
			Node1.setAttribute "HeadID",iHeadID
			Node1.setAttribute "ComputeFormula",ComputeFormula
			Root.Appendchild Node1
				
	
	'Msgbox TempData.XMl
	'Exit Function
	objhttp.Open "Post","XMLcomputeSave.asp?Name=TDSComputeSave&Mod=Acc", false
	objhttp.send TempData.XMLDocument
	If objhttp.responseText <> "" Then
		alert(objhttp.responseText)
	Else
		window.returnvalue = "Y"
		window.close()
	End If	
	
	
	
End Function
</Script>
<script language="javascript">
window.__itmsPopupCompat = { type: "tdsComputation" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>



	<form method="POST" name="formname" action onload="Loadvalues()">
	<table border="0" width="100%" cellspacing="0" cellpadding="0" class="PopupTable">
		<tr>
			<td align="center" class="PageTitle" height="20">Computation Detail
			</td>
		</tr>
		<input type="hidden" name="GroupCode" value="<%=GroupCode%>">
		<input type="hidden" name="HeadID" value="<%=iHeadID%>">
		<input type="hidden" name="iRowCount" value="1">
		
		<tr>
			<td align="center" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="center" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="center">
									</td>
									<td width="100%">
										<table border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td class="FieldCell">TDS Group Name
												</td>
												<td class="FieldCellSub">
													<input type="text" name="txtGname" size="30" maxlength="13" class="FormElem" ReadOnly>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Computation for
												</td>
												<td class="FieldCellSub">
													<input type="text" name="txtcomputationfor" size="30" maxlength="13" class="FormElem" ReadOnly>
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Percentage
												</td>
												<td class="FieldCellSub">
													<input type="text" name="txtpercentage" size="4" maxlength="3" class="FormElem">
												</td>
											</tr>

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
									</td>
									<td>
										<div class="frmBody" id="td" style="width: 355; height:120;">
											<table border="0" cellspacing="1" class="ExcelTable" width="345" id="tbltds">
												<tr>
													<td class="ExcelHeaderCell" align="center" width="10">S.No.
													</td>
													<td class="ExcelHeaderCell" align="center" width="10">
													</td>
													<td class="ExcelHeaderCell" align="center">TDS Head Name
													</td>
												</tr>

												

												<center>
												</table>
											</div>
										</center>
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
												<td valign="middle" class="ActionCell" align="center">
													<input type="button" value="Save" name="B1" class="ActionButton" onClick="UpdateXML()">
 													<input type="button" value="Close" name="B2" class="ActionButton" onclick="window.close();">
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
</body>
