<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	TDSGroupingSetup.asp
	'Module Name				:	Accounts-TDS (Master Amedment)
	'Author Name				:	Kumar K.A.
	'Created On					:	January 17 2007
	'Modified By				:	UmaMaheswari S
	'Modified On				:	May 05, 2011
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'							:
	'Connects To				:	
	'Procedures/Functions Used	:
	'Internal Variables			:

	'Database					:	ITMS_Test
	'Queries Used				:
	'Counters					:
	'String						:
	'Boolean					:
	'Object Holders				:
	'Description				:
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html><head>
<title>iTMS</title>
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/Accpopulate.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="STYLESHEET" href="../../assets/styles/StandardBody.css" type="text/css">
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
</head>
<XML ID="TempData"><Root/></XML>
<XML id="OutData"><Root/></xml>
<xml id="GLHeadData"><Root></Root></xml>
<%
Dim Objrs,OrgUnit,GroupName,GNameText,sType,sTemp,nGroupHeadID

Set Objrs = Server.CreateObject("ADODB.RecordSet")

OrgUnit = Session("organizationcode")
sTemp  = Trim(Request.QueryString("CallType"))
sType = Split(sTemp,":")(0)

'Response.Write "<p><Font color=red>Data="&sTemp

If sType = "E" Then
	'GroupName = Request.Form("SelGPName")
	GroupName =  Split(sTemp,":")(1)
Else
	GroupName = "0"
	nGroupHeadID = "0"
End IF
GNameText = Request.Form("TxtGroupName")

sql = "Select GroupID,GroupName From ACC_M_TDSGroup Where isNull(Useable,'Y') = 'Y' and OUDefinitionid = '"&OrgUnit&"' "
With Objrs
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = sql 
	.Open 
End With
Do While Not Objrs.EOF 
	If (Cstr(GroupName) = Cstr(Objrs(0))) Or (Cstr(GNameText)= Cstr(Objrs(1))) Then 
		GroupName = Objrs(0)
		GNameText = objrs(1)
	End IF
	Objrs.MoveNext 
Loop
Objrs.Close 
'Response.Write "<p>Data="&GroupName & " = "& GNameText
'If GroupName="A" Then GroupName = 0
%>
<Script Language= VBScript>

Function AddGroup()
	Dim oRow,hRowvalue,oText1,oText,iCount,OrgID
	Dim objhttp,Node1,sCallType
	
	OrgID = document.formname.OrgID.value
	sCallType = document.formname.hType.value

	If sCallType = "E"  Then
		If document.formname.TxtGroupName.value="" or document.formname.GroupName.value = "" Then
		'Or document.formname.selGPName.value = "0" Then
			Msgbox "Group Name Required..!"
			document.formname.TxtGroupName.focus()
			Exit Function
		End IF
	Elseif sCallType = "C" Then
		If document.formname.TxtGroupName.value="" Then
			Msgbox "Group Name Required..!"
			document.formname.TxtGroupName.focus()
			Exit Function
		End IF
	End IF
		
	If document.formname.TxtHead.value ="" Or document.formname.selHead.value ="0" Then
		Msgbox("Entet Head Name..!")
		document.formname.TxtHead.focus()
		Exit Function
	ElseIf document.formname.TxtAccHeadName.value="" Then
		Msgbox("Select Account Head..!")
		Exit Function
	End If  
	
	If Not IsNumeric(document.formname.TxtHierachy.value) Then
	    MsgBox "Enter Numeric Value For Hierachy"
	    document.formname.TxtHierachy.value = ""
	    Exit Function
	End If  
	
	If document.formname.hHeadName.value <> document.formname.TxtGroupName.value Then
	'	set oRow = document.all.tblemp.insertRow(document.all.tblemp.rows.length)
	'	set headerCell=oRow.insertCell()				
	'	headerCell.innerHTML=  document.formname.TxtGroupName.value   '"<b>TDS Group Head</b>"
	'	headerCell.className="ExcelDisplayCell"
	'	headerCell.align="Left"
	'	headerCell.colspan=5	
		document.formname.hHeadName.value = document.formname.TxtGroupName.value    
	End If    

	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	Set Root = TempData.documentElement 
	set Node1 = TempData.CreateElement("Schedule")
		If sCallType = "E" Then
			Node1.setAttribute "GroupID",document.formname.GroupName.value	'document.formname.selGPName.value
		Else
			Node1.setAttribute "GroupID","A"
		End IF
	 	Node1.setAttribute "GroupName",document.formname.TxtGroupName.value
	 	Node1.setAttribute "HeadID",document.formname.selHead.value
	 	Node1.setAttribute "HeadName",document.formname.TxtHead.value
	 '	MsgBox document.formname.hAccHead.value    
	 	Node1.setAttribute "HeadCode",document.formname.hAccHead.value
	 	Node1.setAttribute "AccountHead",document.formname.TxtAccHeadName.value  
	 '	Node1.setAttribute "GroupID",document.formname.selGPName.value 
	 	Node1.setAttribute "HeadDetails",document.formname.selHead.value     
	 	If document.formname.R1(0).checked = True Then  
		 	Node1.setAttribute "Mode", "F"
		Else
			Node1.setAttribute "Mode", "P"
		End If
	 	Node1.setAttribute "Hierarchy",document.formname.TxtHierachy.value   
	 	Node1.setAttribute "CreatedOn", "0"'
		Node1.setAttribute "sOrgID",document.formname.OrgID.value
		Root.Appendchild Node1
	
	objhttp.Open "Post","XMLTDSSave.asp?Name=TDSDetails&Mod=Acc", false
	objhttp.send TempData.XMLDocument
	   
	If objhttp.responseText <> "" Then
		alert(objhttp.responseText)
	Else
		document.formname.action = "TDSGroupingSetup.asp?CallType="&document.formname.hRequest.value
		document.formname.submit()
	End If
End Function
	

Function Submit()

Dim objhttp,sCallType
Dim tottext,HeadID,sFormula,i,GroupID,id
GroupID = document.formname.GroupName.Value 
sCallType = document.formname.hType.value

Set objhttp = CreateObject("Microsoft.XMLHTTP")

Set Root = TempData.documentElement 
sFormula = ""
id = "3"

If document.formname.TCount.value = "0" Then 
	'IF document.formname.selGPName.selectedIndex > 0 Then
		MsgBox "Saved.. Sucessfully! "
	'End IF
	Exit Function
End IF
	
	For i = 1 to document.formname.iTxtCount.Length-1
		set Node1 = TempData.CreateElement("TDS")
		Node1.setAttribute "id",id
		Node1.setAttribute "GroupID",GroupID
		Node1.setAttribute "TDSHeadID",document.formname.TxtVal(i).Value 
		Node1.setAttribute "Formula",document.formname.txtFormula(i).value   
		Root.Appendchild Node1
	Next
	'MsgBox TempData.XML
	objhttp.Open "Post","XMLTDSFormulaSave.asp?Name=TDSFormulaUpdate&Mod=Acc", false
	objhttp.send TempData.XMLDocument
	
	If objhttp.responseText <> "" Then
		alert(objhttp.responseText)
	Else
		MsgBox "Saved...!"	
		'document.formname.action = "TDSGroupingSetup.asp?CallType="&document.formname.hRequest.value
		document.formname.action = "TdsGroups.asp" 
		document.formname.submit()
	End If
	
	
End Function

	
	
Function SuppName()
	
	Dim iUnitNo,saTemp,iGlHead,sRetVal,OutValue,sAccDesc,nAccHeadHead,sOrgId
	
	sOrgId =document.formname.OrgID.value
	
	set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?orgId="+sOrgId,GLHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
	    sQuery = OutValue.getAttribute("PassQuery")
	    if OutValue.getAttribute("Action")="CLOSE" then exit function

	while OutValue.getAttribute("Action")<>"Done"
		set OutValue = showModalDialog("../../Common/GLHeadSelection.asp?"&sQuery,GLHeadData,"dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		    sQuery = OutValue.getAttribute("PassQuery")
		    if OutValue.getAttribute("Action")="CLOSE" then exit function
	wend
	alert(OutValue.xml)
	'AccountDescription,AccountHead
	if OutValue.hasChildNodes() then
	    For each ndChild in OutValue.childNodes
	        nAccHeadHead = ndChild.getAttribute("RetField0")
	        sAccDesc = ndChild.getAttribute("RetField5")
	    Next
	end if
	
	document.formname.hAccHead.value = nAccHeadHead
	document.formname.TxtAccHeadName.value = sAccDesc
End Function

Function Calc(GroupCode,HeadID)
Dim OutValue,sVal,iHeadVal
	iHeadVal = HeadID
	showModalDialog "TDSComputationDetailPopup.asp?GroupCode="&GroupCode&"&HeadID="&iHeadVal,"A","dialogHeight:320px;dialogWidth:710px;center:Yes;help:No;resizable:No;status:No"
End Function

Function GNameChange()
Dim oRow,hRowvalue,oText1,oText,iCount
Dim GPValue,GroupID
Dim objhttp
Dim Val,Val1,Val2,Val3,Val4,Val5	
	'MsgBox document.formname.selGPName.value   
	document.formname.TxtGroupName.Value = ""  
	sCallType = document.formname.hType.value
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	'If document.formname.selGPName.value <>"A" then
	If document.formname.GroupName.value <> "A" Then
	 
		document.formname.action = "TDSGroupingSetup.asp?CallType="&document.formname.hRequest.value
		document.formname.submit()
	Else
		document.formname.TxtGroupName.Value = ""  
	End If
	
'	document.formname.TxtGroupName.value = document.formname.selGPName.options(document.formname.selGPName.selectedIndex).Text      
	'If document.formname.selGPName.value <>"A" Then
	'	GPValue = Split(document.formname.selGPName.value,",")
	'	GroupID = GPValue(0) 
	'	id = "1"
	'	document.formname.selHead.length = 0
	'	document.formname.selHead.length = document.formname.selHead.length + 1
	'	document.formname.selHead.options(document.formname.selHead.length-1).text = "Add New"
	'	document.formname.selHead.options(document.formname.selHead.length-1).Value = "A" 
	'	objhttp.Open "GET","TDSXMLGenerate.asp?GroupID="&GroupID&"&id="&id, false
	'	objhttp.send 
	'	IF objhttp.responsexml.XML <> "" Then
	'		OutData.loadXML objhttp.responseXML.xml
	'		Set Root = OutData.documentELement
	'		sExp = "//Root"
	'		Set HeadNode = Root.selectNodes(sExp)
	'		recno = HeadNode.length-1
	'	'	document.all.tblemp.rows = 0
	'		For Each HeadNode in Root.ChildNodes
	'			Val1 = HeadNode.Attributes.Item(0).nodeValue
	'			Val2 = HeadNode.Attributes.Item(2).nodeValue
	'			Val3 = HeadNode.Attributes.Item(3).nodeValue
	'			Val4 = HeadNode.Attributes.Item(4).nodeValue
	'			Val5 = HeadNode.Attributes.Item(5).nodeValue
	'			Val = Val1&","&Val2&","&val3&","&Val4&","&Val5    
	'			document.formname.selHead.length = document.formname.selHead.length + 1
	'			document.formname.selHead.options(document.formname.selHead.length-1).text = HeadNode.Attributes.Item(1).nodeValue
	'			document.formname.selHead.options(document.formname.selHead.length-1).Value = Val 
	'			 				
	'			set oRow = document.all.tblemp.insertRow(document.all.tblemp.rows.length)
	'			iCount = document.all.tblemp.rows.length
	'			iCount = document.formname.SelSeh.value   
	'			set headerCell=oRow.insertCell()				
	'			headerCell.innerHTML=Cint(document.formname.iRowCount.value)   
	'			headerCell.className="ExcelHeaderCell"
	'			headerCell.align="center"
	'			Set headerCell=oRow.insertCell()
	'			headerCell.innerHTML= HeadNode.Attributes.Item(1).nodeValue
	'			headerCell.className = "ExcelInputCell"'"ExcelInputCell"
	'			headerCell.align = "center"
	'
	'			Set headerCell=oRow.insertCell()
	'			headerCell.innerHTML= Val3 
	'			headerCell.className = "ExcelInputCell"'"ExcelInputCell"
	'			headerCell.align = "center"
	'
	'			If Val2  = "F" Then  
	'				Set headerCell=oRow.insertCell()
	'				set oText = document.createElement("<input type=""Text""  name=""txtdocument"&iRowCount&""" class=""FormElem"" >" )
	'				'headerCell.innerHTML= "New"
	'				headerCell.appendChild(oText)
	'				headerCell.className = "ExcelInputCell"'"ExcelInputCell"
	''				headerCell.align = "center"
	'				document.formname.iTxtCount.value = CInt(document.formname.iTxtCount.value)+1      
	'			Else	
	'				set headerCell=oRow.insertCell()									
	'				set oText = document.createElement("<input type=""Button""  name=""btnselect"&document.formname.iButCount.value&""" class=""ActionButton"" Value=""Select"" Onclick=""PopUpCall()"">"  )
	'				headerCell.appendChild(oText)
	'				headerCell.className="ExcelInputCell"
	'				headerCell.align="center"
	'				document.formname.iButCount.value = CInt(document.formname.iButCount.value)+1    
	'			End If
	'			
	'			document.formname.iRowCount.value = document.formname.iRowCount.value+1     
	'		Next
	'	End IF
	'End If
	'MsgBox document.formname.selGPName.value   
End Function



Function UnitChange(sType)
	Dim	sOrgID,id
	Dim objhttp
	set objhttp = CreateObject("MSXML2.XMLHTTP")
	
	sOrgID = document.formname.OrgID.value 
	If sType = "E" Then
		document.formname.selGPName.length = 0  
		'document.formname.selGPName.length = document.formname.selGPName.length + 1
		'document.formname.selGPName.options(document.formname.selGPName.length-1).text = "Add New"
		'document.formname.selGPName.options(document.formname.selGPName.length-1).Value = "A"
		'  msgbox sOrgID 
		id = "0"
		objhttp.Open "GET","TDSXMLGenerate.asp?sOrgID="&sOrgID&"&id="&id, false
		objhttp.send
		'MsgBox objhttp.responseText
		IF objhttp.responsexml.XML <> "" Then
					OutData.loadXML objhttp.responseXML.xml
					Set Root = OutData.documentELement
					sExp = "//Root"
					Set HeadNode = Root.selectNodes(sExp)
					recno = HeadNode.length-1	
					For Each HeadNode in Root.ChildNodes
						document.formname.selGPName.length = document.formname.selGPName.length + 1
						document.formname.selGPName.options(document.formname.selGPName.length-1).text = HeadNode.Attributes.Item(0).nodeValue
						document.formname.selGPName.options(document.formname.selGPName.length-1).Value = HeadNode.Attributes.Item(1).nodeValue
						'MsgBox HeadNode.Attributes.Item(0).nodeValue
					Next
		End IF
	End IF
End Function

Function Del()
Dim sOrgID,CmbGroupValue,CmbHeadValue,SpltGroup,SpltHead
	sOrgID = document.formname.OrgID.value   
	CmbGroupValue = document.formname.GroupName.value'document.formname.selGPName.value 
	CmbHeadValue = document.formname.selHead.value     
End Function

Function GChange()
	'MsgBox "Ok"
End Function 

Function ShowVouch(iTdsHead)
Dim GroupHeadID,GroupName,GroupHeadName,ComputeMode,AcHeadCode
Dim Herarchy,Formula,AccHeadName,recno,HeadNode,sExp
Dim GroupID,id,iHeadID,Root,sCallType
iHeadID = iTdsHead 
'MsgBox iHeadID 
GroupID = document.formname.GroupName.Value 
sCallType = document.formname.hType.value
If sCallType = "C" Then  Exit Function

Dim objhttp,Node1
	id = "3"
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	objhttp.Open "GET","TDSXMLGenerate.asp?HeadID="&iHeadID&"&GroupID="&GroupID&"&id="&id, false
	objhttp.send
	OutData.loadXML objhttp.responseXML.xml
	'MsgBox OutData.XML
	Set Root = OutData.documentELement
	sExp = "//Root"
	Set HeadNode = Root.selectNodes(sExp)	
	recno = HeadNode.length-1
	'MsgBox recno
	For Each HeadNode in Root.ChildNodes
		GroupHeadID = HeadNode.Attributes.Item(0).nodeValue
		GroupName = HeadNode.Attributes.Item(6).nodeValue
		GroupHeadName =  HeadNode.Attributes.Item(7).nodeValue
		ComputeMode =  HeadNode.Attributes.Item(2).nodeValue
		AcHeadCode =  HeadNode.Attributes.Item(3).nodeValue
		Herarchy =  HeadNode.Attributes.Item(4).nodeValue
		Formula =  HeadNode.Attributes.Item(5).nodeValue
		AccHeadName =  HeadNode.Attributes.Item(8).nodeValue
'	MsgBox GroupHeadID 
	Next
	document.formname.selHead.selectedIndex = 0  
	document.formname.selHead.disabled = True
	document.formname.TxtGroupName.value = GroupName   
	document.formname.TxtHead.value = GroupHeadName   
	document.formname.TxtAccHeadName.value = AccHeadName   
	document.formname.hAccHead.value = AcHeadCode 
	document.formname.TxtHierachy.value = Herarchy
	If ComputeMode = "F" Then
		document.formname.R1(0).checked = True
	Else
		document.formname.R1(1).checked = True
	End If  
	document.formname.ButUpdate.disabled = False
	document.formname.ButDelete.disabled = False
	document.formname.ButAdd.disabled = True
End Function

Function UpdateGroup()
Dim GroupHeadID,GroupName,GroupHeadName,ComputeMode,AcHeadCode,GroupID
Dim Herarchy,Formula,AccHeadName,recno,HeadNode,sExp,id
Dim objhttp
Dim objhttp1,Node1
GroupID = document.formname.GroupName.Value 
sCallType = document.formname.hType.value
Set objhttp = CreateObject("Microsoft.XMLHTTP")
	Set Root = OutData.documentELement
	sExp = "//Root"
	Set HeadNode = Root.selectNodes(sExp)	
	recno = HeadNode.length-1
	'MsgBox recno
	For Each HeadNode in Root.ChildNodes
		GroupHeadID = HeadNode.Attributes.Item(0).nodeValue
		GroupName = document.formname.TxtGroupName.value 
		GroupHeadName =  document.formname.TxtHead.value
		If document.formname.R1(0).checked = True Then
		ComputeMode = "F"
		Else
		ComputeMode = "P"
		End If
		AcHeadCode =  document.formname.hAccHead.value 
		Herarchy =  document.formname.TxtHierachy.value
		AccHeadName =  document.formname.TxtAccHeadName.value
	Next 
	id ="1"
	Set Root = TempData.documentElement 
	set Node1 = TempData.CreateElement("TDS")
		Node1.setAttribute "id",id
		Node1.setAttribute "GroupID",GroupID
		Node1.setAttribute "GroupHeadID",GroupHeadID 
		Node1.setAttribute "GroupName",GroupName
		Node1.setAttribute "GroupHeadName",GroupHeadName
		Node1.setAttribute "ComputeMode",ComputeMode
		Node1.setAttribute "AcHeadCode",AcHeadCode
		Node1.setAttribute "Herarchy",Herarchy
		Node1.setAttribute "AccHeadName",AccHeadName
	Root.Appendchild Node1
	objhttp.Open "Post","XMLTDSUpdateDelete.asp?Name=TDSUpdateDelete&Mod=Acc", false
	objhttp.send TempData.XMLDocument
	If objhttp.responseText <> "" Then
		alert(objhttp.responseText)
	Else
	document.formname.action = "TDSGroupingSetup.asp?CallType="&document.formname.hRequest.value
	document.formname.submit()
	End If
	
End Function


Function DeleteGroup()
Dim GroupHeadID,GroupName,GroupHeadName,ComputeMode,AcHeadCode,GroupID
Dim Herarchy,Formula,AccHeadName,recno,HeadNode,sExp,id
Dim objhttp
Dim objhttp1,Node1

GroupID = document.formname.GroupName.Value 
sCallType = document.formname.hType.value

	Set objhttp = CreateObject("Microsoft.XMLHTTP")

	Set Root = OutData.documentELement
	sExp = "//Root"
	Set HeadNode = Root.selectNodes(sExp)	
	recno = HeadNode.length-1
	
	'MsgBox recno
	For Each HeadNode in Root.ChildNodes
		GroupHeadID = HeadNode.Attributes.Item(0).nodeValue
		GroupName = document.formname.TxtGroupName.value 
		GroupHeadName =  document.formname.TxtHead.value
		If document.formname.R1(0).checked = True Then
		ComputeMode = "F"
		Else
		ComputeMode = "P"
		End If
		AcHeadCode =  document.formname.hAccHead.value 
		Herarchy =  document.formname.TxtHierachy.value
		AccHeadName =  document.formname.TxtAccHeadName.value
	Next 
	id ="2"
	Set Root = TempData.documentElement 
	set Node1 = TempData.CreateElement("TDS")
		Node1.setAttribute "id",id
		Node1.setAttribute "GroupID",GroupID
		Node1.setAttribute "GroupHeadID",GroupHeadID 
		Node1.setAttribute "GroupName",GroupName
		Node1.setAttribute "GroupHeadName",GroupHeadName
		Node1.setAttribute "ComputeMode",ComputeMode
		Node1.setAttribute "AcHeadCode",AcHeadCode
		Node1.setAttribute "Herarchy",Herarchy
		Node1.setAttribute "AccHeadName",AccHeadName
	Root.Appendchild Node1
	objhttp.Open "Post","XMLTDSUpdateDelete.asp?Name=TDSUpdateDelete&Mod=Acc", false
	objhttp.send TempData.XMLDocument
	If objhttp.responseText <> "" Then
		alert(objhttp.responseText)
	Else
	document.formname.TxtGroupName.Value=""
	document.formname.TxtHead.value = ""  
	document.formname.action = "TDSGroupingSetup.asp?CallType="&document.formname.hRequest.value
	document.formname.submit()
	End If

End Function

Function TDSDel()
	document.formname.action = "TDSGroupingDelete.asp"
	document.formname.submit
End Function



</Script>
<script language="javascript">
window.__itmsPopupCompat = { type: "tdsGroupingSetup" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

	<form method="post" name="formname" action="">
	<input type="hidden" name="SelSeh" value="3">
	<input type="hidden" name="OrgID" value="<%=OrgUnit%>">
	<input type="hidden" name="GroupName" value="<%=GroupName%>">
	<input type="hidden" name="hAccHead" value="">
	<input type="hidden" name="hType" value="<%=sType%>">
	
	<input type="hidden" name="iButCount" value="1">
	<input type="hidden" name="iRowCount" value="1">
	<input type="hidden" name="hHeadName" value="">
	
	<input type="hidden" name="TxtFormula" size="40" maxlength="30" class="FormElem"></td>
	<input type="hidden" name="TxtVal" value="">
	<input type="hidden" name="iTxtCount" value="">
	
	<Input type="hidden" name="hRequest" value="<%=Trim(sTemp)%>">
	<input type="hidden" name="hDelFrom" value="S">
	
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td align="middle" class="PageTitle" height="20">TDS Grouping Setup
			</td>
		</tr>

		<tr>
			<td align="middle" class="TopPack">
			</td>
		</tr>

		<tr>
			<td valign="top">
				<table id="Table16" cellspacing="0" cellpadding="0" border="0" width="100%" >
					<tr>
						<td class="TabBodyWithTopLine">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td align="middle" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="middle">
									</td>
									<td width="100%">
										<table border="0" cellspacing="0" cellpadding="0">
											<!--<tr>
												<td class="FieldCell">Unit
												</td>
												<td class="FieldCellSub"><select size="1" name="selUnit" class="FormElem" Onchange="UnitChange()" >
														<option selected>Select</option>
                        <% 
							Dim sql
							sql = "select Orgunitdescription,OUDefinitionID from DCS_OrganizationunitDefinitions WHere Len(OUDefinitionID) > 4 Order By OUDefinitionID "
							With Objrs
								.CursorLocation = 3
								.CursorType = 3
								.Source = sql 
								.ActiveConnection = con
								.Open 
							End With
								While Not Objrs.EOF 
                        %>
                        <% If OrgUnit = Objrs(1) Then %>
                        <Option  Value="<%=Objrs(1)%>" Selected="<%=Objrs(1)%>"><%=Objrs(0)%></Option>
                        <% Else %>
                        <Option Value="<%=Objrs(1)%>"><%=Objrs(0)%></Option>
                        <%End If%>
						<%
							Objrs.MoveNext
							Wend
							Objrs.Close
						%>                        
							</select>
							</td>
							<td class="FieldCellSub">
							</td>
							<td class="FieldCellSub">Created On
							</td>
							<td class="FieldCellSub">
                        <OBJECT class=formelem id=ctlDate1 
                        style="WIDTH: 89px; HEIGHT: 21px" 
                        codeBase=file://ntserver/Websites/iTMS_Garment_Dyeing/components/DatePicker.CAB#version=1,0,0,0 
                        height=21 width=89 
                        classid=CLSID:01E5BF20-F919-44E6-A698-CF7FD7C7D6CD 
                        viewastext><PARAM NAME="_ExtentX" VALUE="2355"><PARAM NAME="_ExtentY" VALUE="556"></OBJECT>
												</td>
											</tr>-->

											<tr>
												<td class="FieldCell">TDS Group Name
												</td>
												<td class="FieldCellSub" colspan="4">
													<%If sType = "E"and 1=2Then%>
													<select size="1" name="selGPName" class="FormElem" OnChange="GNameChange()">
														<option selected value="0">Select</option>
														<%
														IF CStr(OrgUnit) <> "" Then
															sql = "Select GroupID,GroupName From ACC_M_TDSGroup Where isNull(Useable,'Y') = 'Y' and OUDefinitionid = '"&OrgUnit&"' "
															With Objrs
																.CursorLocation = 3
																.CursorType = 3
																.ActiveConnection = con
																.Source = sql 
																.Open 
															End With
															While Not Objrs.EOF 
															%>
															<% If (Cstr(GroupName) = Cstr(Objrs(0))) Or (Cstr(GNameText)= Cstr(Objrs(1))) Then 
																	GroupName = Objrs(0)
																	GNameText = objrs(1)
															%>
																<Option Selected Value="<%=Objrs(0)%>"><%=Objrs(1)%></Option>
															<% Else %>
																<Option Value="<%=Objrs(0)%>"><%=Objrs(1)%></Option>
															<%End If%>
															<%
															Objrs.MoveNext 
															Wend
															Objrs.Close 
														End IF
														%>
														
													<!--<option value="A">Add New</Option>-->
													</select> 
													<%End IF%>
													<input name="TxtGroupName" size="30" maxlength="100" class="FormElem" Value="<%=GNameText%>" >
												</td>
											</tr>
											<tr>
												<td class="FieldCell">TDS Head
												</td>
												<td class="FieldCellSub" colspan="4"><select size="1" name="selHead" class="FormElem" Onchange="">
													<option selected value="0">Select</option>
													<%
													sql = "select GroupHeadID,GroupHeadName,ComputeMode,AcheadCode,ComputeFormula,Herarchy from ACC_M_TDSHeadComputation Where GroupID="&Cint(GroupName)
													With Objrs
														.CursorLocation = 3
														.CursorType = 3
														.ActiveConnection = con
														.Source = sql
														.Open 
													End With
													While Not Objrs.EOF 
													%>	
													<option Value="<%=Objrs(0)&","&Objrs(2)&","&Objrs(3)&","&Objrs(4)&","&Objrs(5)%>"><%=Objrs(1)%></option>
													<%
													Objrs.MoveNext 
													Wend
													Objrs.Close 
													%>
													
													<option Value="A">Add New</option>
													</select> 
													<input name="TxtHead" maxlength="30" class="FormElem" >
													&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Hierarchy 
													<input name="TxtHierachy" size="4" maxlength="3" class="FormElem" >
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Applicable A/c Head
												</td>
												<td class="FieldCellSub" colspan="4">
													<input name="TxtAccHeadName" size="40" maxlength="30" class="FormElem" >
 													<IMG style="CURSOR: hand" onclick="SuppName()" height=11 alt="Select Account Head" src  ="../../assets/images/iTMS%20Icons/EntryIcon.gif" width=10 align=center border=0 >
												</td>
											</tr>

											<tr>
												<td class="FieldCell">Computation Mode
												</td>
												<td class="FieldCell" colspan="4">
													<input type="radio" value="V2" checked name="R1">
 													Flat&nbsp;&nbsp; 
													<input type="radio" value="V1" name="R1">
 													Percentage
												</td>
											</tr>

											<tr>
												<td class="FieldCell">
											
												</td>
												<td class="FieldCellSub">
													<input type="button" value=" Add " name="ButAdd" class="AddButtonx" Onclick="AddGroup()" >
													<input type="button" value=" Update " name="ButUpdate" class="AddButtonx" Onclick="UpdateGroup()" disabled>
													<input type="button" value=" Delete " name="ButDelete" class="AddButtonx" Onclick="DeleteGroup()" disabled>
												</td>
												<td class="FieldCellSub">
															
												</td>
												<td class="FieldCellSub">
												</td>
												<td class="FieldCellSub">
												</td>
											</tr>

										</table>
									</td>
									<td align="middle">
									</td>
								</tr>

								<tr>
									<td align="middle" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="middle">
									</td>
									<td>
										<div class="frmBody" id="frm2" style="WIDTH: 555px; HEIGHT: 231px">
											<table border="0" cellspacing="1" class="ExcelTable" width="550" id="tblemp">
												<tr>
													<td class="ExcelHeaderCell" align="middle" width="10" rowspan="2">S.No.
													</td>
													<td class="ExcelHeaderCell" align="middle" colspan="5">TDS Group Name
													</td>
												</tr>

												<tr>
													<td class="ExcelHeaderCell" align="middle" >
													</td>
													<td class="ExcelHeaderCell" align="middle" width="100">TDS Head
													</td>
													<td class="ExcelHeaderCell" align="middle" width="300">Account Head
													</td>
													<td class="ExcelHeaderCell" align="middle" > Detail
													</td>
												</tr>
												<%
												Dim Txtcount,HID,Formula
													txtcount = 0
													sql = "Select T.GroupHeadID,T.GroupHeadName,T.ComputeMode,T.AcHeadCode,M.AccountDescription,T.ComputeFormula From "&_
														   "ACC_M_TDSHeadComputation T,Acc_M_GLAccountHead M Where T.AcHeadCode = M.AccountHead "&_
															" and T.GroupID = '"&Cint(GroupName)&"'"
													'Response.Write sql
													With Objrs
														.CursorLocation = 3
														.CursorType = 3
														.ActiveConnection = con
														.Source = sql 
														.Open 
													End With
													While Not Objrs.EOF 
													HID = Objrs(0)						
													If Objrs(5) <> "" then Formula = Objrs(5) Else Formula = ""
													%>
													<tr>
														<td class="ExcelSerial" align="center"><%=Objrs(0)%> 
														</td>
														<td class="ExcelDisplayCell" align="left" >
														<a href="#" LANGUAGE="VBSCRIPT" onclick="ShowVouch(<%=Objrs(0)%>)" class="ExcelDisplayLink">Edit</a></td>

														<td class="ExcelDisplayCell" ><%=Objrs(1)%>
														</td>
														<td class="ExcelDisplayCell"><%=Objrs(4)%> 
														</td>
														<%If Objrs(2) = "F" Then %>
														<%If Formula <>"" Then%>
																<td class="ExcelInputCell"><input name="TxtFormula" size="14" maxlength="30" Value="" class="FormElem" style="text-align:right"></td>
															<%Else%>
																<td class="ExcelInputCell"><input name="TxtFormula" size="14" maxlength="30" Value="" class="FormElem" style="text-align:right"></td>
															<%End If%>
																<input type="hidden" name="TxtVal" value="<%=HId%>">
																<input type="hidden" name="iTxtCount" value="<%=TxtCount%>">
														<%
														Txtcount = Txtcount + 1
														Else%>			
														<td class="ExcelDisplayCell" align="center"><input type="Button" value=" Select " name="BtnCalc" class="ActionButtonX" onclick="Calc(<%=Cint(GroupName)%>,<%=Cint(HID)%>)"></td>
														<%End If%>
														
														
													</tr>
													<%
														Objrs.MoveNext 
														Wend	
													Objrs.Close 
													If Txtcount>0 Then
													%>
														<input type="hidden" name="TCount" value="1">
														<%
														Else
														%>
															<input type="hidden" name="TCount" value="0">
														<%End If%>
												</table>
											</div>
									</td>
									<td align="middle">
									</td>
								</tr>

								<tr>
									<td align="middle" colspan="3" class="MiddlePack">
									</td>
								</tr>

								<tr>
									<td align="middle">
										<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
									</td>
									<td valign="top">
										<table border="0" cellpadding="0" cellspacing="0" width="100%">
											<tr>
												<td valign="center" class="ActionCell" align="middle">
													<input type="Button" value="Save" name="BtnSubmit" class="ActionButton" onclick="Submit();">
													<input type="Button" value="Delete" name="BtnDel" class="ActionButton" onclick="TDSDel()">
													<!--input type="button" value="Done" name="B1" class="ActionButton" -->
													
												</td>
											</tr>

										</table>
									</td>
									<td align="middle">
										<IMG height=5 src="../../assets/images/clearpixel.gif" width=5 border=0>
									</td>
								</tr>

								<tr>
									<td align="middle" colspan="3" class="BottomPack">
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
