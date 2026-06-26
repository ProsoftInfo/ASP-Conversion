<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=-10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" %>
<%
	'Program Name				:	AddSchedSubHeads.asp
	'Module Name				:	ACCOUNTS (Master BalSheet and P&L)
	'Author Name				:	Maheshwari S.
	'Created On					:	Dec 19 2006
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
	'Connects To				:	SchSetUp.asp
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
<!--#include file="../../include/Accpopulate.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><title>iTMS - Accounts - Add Schedule Sub Heads</title>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<% sInsDate = Request("InsDate") %>
<XML ID="XmlData">
<Root>
	<Details OrgID="" SchID="" LevelID="" Level1ID="" Level2ID="" Level1Name="" Level2Name="" ModeType="" AccHead="" AccHeadName="" FinYear="" ComputeMode="" InsDate="<%=sInsDate%>">
	</Details>
</Root>
</XML>
<XML id="OutData"><Root/></xml>
<XML ID="TempData"><Root/></XML>
<%
	Dim sOrgId,sSchedNo,Objrs1,Root,iCtr,sCatCode,sName,objfs
	Dim sql,sNo,sHead,sHiera,sApp,sFinyr,iSchId,oDOM,iNo,iCnt,iHiera,sInsDate
	Set objrs1 = Server.createObject("ADODB.Recordset")
	Set objfs = CreateObject("Scripting.FileSystemObject")
	sOrgId = Request("sUnit")
	sSchedNo = Request("sSchName")
	sFinYr = Session("FinPeriod")
	sCatCode = Request("sCatCode")
	
%>
	

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Function SetLevelFun()
	'MsgBox "1"
	Dim objhttp,sschedno,sOrgID,sFinyr,id,EntryType
	Dim iSchNo,SubID,SubSubID,iSplit,sExp,HeadNode 
	sOrgID = document.formname.hOrgId.Value
	sFinyr = document.formname.sfinyr.value  
	sschedno = document.formname.sschedno.value
	If document.formname.selLevel1.value = "0" Then
	   MsgBox("Select SubHeading And Proceed...!"):Exit Function
	Else
		document.formname.txtLev1.disabled = False 


		If document.formname.selLevel1.value <>"A" Then 
			document.formname.txtLev1.value = Trim(document.formname.selLevel1.options(document.formname.selLevel1.selectedIndex).text)
	  		document.formname.txtLev1.size = Len(Trim(document.formname.selLevel1.options(document.formname.selLevel1.selectedIndex).text)) + 7       
	  	Else
	  		document.formname.txtLev1.value = ""
	  	End If
	End If   
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		document.formname.txtLev2.disabled = False
		document.formname.selLevel2.length = 0
		document.formname.selLevel2.length = document.formname.selLevel2.length + 1
		document.formname.selLevel2.options(document.formname.selLevel2.length-1).text = "Add New"
		document.formname.selLevel2.options(document.formname.selLevel2.length-1).Value = "A"   
		'If document.formname.optLevel(1).checked = True And document.formname.selLevel1.value <>"A" Then   
		If document.formname.selLevel1.value <>"A" Then   
			iSplit = Split(document.formname.selLevel1.value,"-")
			SubID= iSplit(0)
			SubSubID = iSplit(1)     
			EntryType = iSplit(2) 
		
			If isplit(3)="+" Then
				document.formname.optCompMode(0).checked = True
			Else
				document.formname.optCompMode(1).checked = True    
			End If  
		
			'Msgbox EntryType
		End If
		
		If document.formname.optLevel(1).checked = True Then
		id = "2"
		'	MsgBox sOrgID&EntryType&sFinyr&sschedno&SubID&SubSubID&id   
		objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&EntryType="&EntryType&"&sFinyr="&sFinyr&"&sschedno="&sschedno&"&SubID="&SubID&"&SubSubID="&SubSubID&"&id="&id, false
		objhttp.send
		'MsgBox objhttp.responsexml.XML	
		IF objhttp.responsexml.XML <> "" Then
			OutData.loadXML objhttp.responseXML.xml
			Set Root = OutData.documentELement
		sExp = "//Root"
		Set HeadNode = Root.selectNodes(sExp)
		'recno = HeadNode.length-1
		'msgbox recno 
		For Each HeadNode in Root.ChildNodes
			document.formname.selLevel2.length = document.formname.selLevel2.length + 1
			document.formname.selLevel2.options(document.formname.selLevel2.length-1).text = HeadNode.Attributes.Item(1).nodeValue
			document.formname.selLevel2.options(document.formname.selLevel2.length-1).Value = HeadNode.Attributes.Item(0).nodeValue'+"-"+HeadNode.Attribute.Item(2).nodeValue
		'MsgBox HeadNode.Attributes.Item(0).nodeValue
		Next
		End IF
		End If  
		'Msgbox EntryType
	'	If document.formname.selLevel2.length < 2  Then
		If EntryType = "D" Then
			document.formname.optMode(0).checked = true 
		ElseIf EntryType = "A" Then
			document.formname.optMode(1).checked = true
			document.formname.ButAcHead.disabled = False
				id = "3"
				'MsgBox SubID&SubSubID&EntryType&sschedno  
				objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&EntryType="&EntryType&"&sFinyr="&sFinyr&"&sschedno="&sschedno&"&SubID="&SubID&"&SubSubID="&SubSubID&"&id="&id, false
				objhttp.send
	'			MsgBox objhttp.responsexml.XML	
				IF objhttp.responsexml.XML <> "" Then
				OutData.loadXML objhttp.responseXML.xml
				Set Root = OutData.documentELement
				sExp = "//Root"
				Set HeadNode = Root.selectNodes(sExp)
				'recno = HeadNode.length-1
				'msgbox recno 
				For Each HeadNode in Root.ChildNodes
					document.formname.hAccHead.value = HeadNode.Attributes.Item(0).nodeValue
					document.formname.txtAcHead.value = HeadNode.Attributes.Item(1).nodeValue
				Next
				End IF
		ElseIf EntryType = "S" Then
			 document.formname.optMode(2).checked = true
		Elseif EntryType = "N" then 
			 document.formname.optMode(3).checked = true 
		End if
	'	End If
		
		
End Function


Function LevelFun()
	If document.formname.optLevel(0).checked = True then  
		document.formname.txtLev2.value = ""  
		document.formname.txtLev2.disabled = True  
		'document.formname.txtAcHead.value = ""
		'document.formname.txtAcHead.disabled = True   
		document.formname.selLevel2.disabled = true
		document.formname.FinYear.disabled = true
	'	document.formname.optMode(3).checked = True  
	Else
		document.formname.txtLev2.disabled = False
		'document.formname.txtAcHead.disabled = False   
		document.formname.selLevel2.disabled = False
		document.formname.FinYear.disabled = False
	'	document.formname.optMode(0).checked = True  
	End If
End Function
	
Function setlevelfun1()
	Dim iModeType,iArray,SubSubID,isplit,sOrgID,sFinyr,sschedno
	Dim objhttp
	sOrgID = document.formname.hOrgId.Value
	sFinyr = document.formname.sfinyr.value  
	sschedno = document.formname.sschedno.value
	set objhttp = CreateObject("MSXML2.XMLHTTP")

	If document.formname.selLevel2.value <>"A" Then   
		document.formname.txtLev2.value = document.formname.selLevel2.options(document.formname.selLevel2.selectedIndex).text        
		document.formname.txtLev2.disabled = false  
		
		iArray = Split(document.formname.selLevel2.value,"-")
		SubSubID =  iArray(0)
		iModeType = iArray(1)
		isplit = split(document.formname.selLevel1.value,"-")
		 
	 
		If iModeType = "A" Then
			document.formname.optMode(1).checked = True  
			document.formname.hAccHead.value = iArray(2) 
			If isplit(2)<>"" Then   	
				id = "3"
				objhttp.Open "GET","AccSubAndSubSubID.asp?sOrgID="&sOrgID&"&EntryType="&iModeType&"&sFinyr="&sFinyr&"&sschedno="&sschedno&"&SubID="&isplit(0)&"&SubSubID="&SubSubID&"&id="&id, false
				objhttp.send
				IF objhttp.responsexml.XML <> "" Then
				OutData.loadXML objhttp.responseXML.xml
				Set Root = OutData.documentELement
				sExp = "//Root"
				Set HeadNode = Root.selectNodes(sExp)
				For Each HeadNode in Root.ChildNodes
					document.formname.hAccHead.value = HeadNode.Attributes.Item(0).nodeValue
					document.formname.txtAcHead.value = HeadNode.Attributes.Item(1).nodeValue
				Next
				End IF
			
			
			End If
			
			
		ElseIf iModeType = "D" Then
			document.formname.optMode(0).checked = true 
		ElseIf iModeType = "S" Then
			document.formname.optMode(2).checked = true 
		ElseIf iModeType = "N" Then
			document.formname.optMode(3).checked = true 
		End IF
		
	Else
	document.formname.txtLev2.value = "" 
	End If

End Function

Function ModeFun()
	if document.formname.optMode(1).checked = true then
		document.formname.ButAcHead.disabled = false
	else
		document.formname.ButAcHead.disabled = true  
		document.formname.txtAcHead.value = ""    
	end if
End Function
	
Function AccHeadClck()
	document.formname.txtAcHead.disabled = false 
End Function
	
Function popAccList()
	Dim iUnitNo,saTemp,iGlHead,sRetVal,OutValue,sAccHeadName,sNewAcc,sOrgId
	sOrgId = document.formname.hOrgId.value 
	OutValue= showModalDialog("ChgAccHeadName.asp?orgId="+sOrgId,"","dialogHeight:520px;dialogWidth:520px;center:Yes;help:No;resizable:No;status:No")
	'MsgBox OutValue
	arrTemp = split(OutValue,":")
				
	while UBound(arrTemp) = 0 
		OutValue = showModalDialog("ChgAccHeadName.asp?"&OutValue,"","dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No")
		arrTemp = split(OutValue,":")
	wend
	sRetVal = OutValue

	'MsgBox nEntNo
	'msgbox OutValue		
	IF UBound(arrTemp) <= 1 then exit function
	document.formname.hAccHead.value = arrTemp(0)
	sAccHeadName = arrTemp(1)
	document.formname.txtAcHead.value = sAccHeadName
	'MsgBox document.formname.hAccHead.value 
End Function 
	
Function SaveXML()
	Dim sExp,TempNode,recno
	IF CheckVal() Then
		Set Root = XmlData.documentElement
		Set objhttp = CreateObject("Microsoft.XMLHTTP")
		sExp = "//Details"
		Set TempNode = Root.selectNodes(sExp)
		recno = TempNode.length-1
	
		TempNode.Item(recno).Attributes.getNamedItem("OrgID").value = document.formname.hOrgId.Value
		TempNode.Item(recno).Attributes.getNamedItem("SchID").value = document.formname.selSch.value
	
		If document.formname.optLevel(0).checked = True Then  
			TempNode.Item(recno).Attributes.getNamedItem("LevelID").value = 0 
		Else
			TempNode.Item(recno).Attributes.getNamedItem("LevelID").value = 1
		End If
	
		TempNode.Item(recno).Attributes.getNamedItem("Level1ID").value = document.formname.selLevel1.value 
		TempNode.Item(recno).Attributes.getNamedItem("Level2ID").value = document.formname.selLevel2.value    
		TempNode.Item(recno).Attributes.getNamedItem("Level1Name").Value = document.formname.txtLev1.value 
		TempNode.Item(recno).Attributes.getNamedItem("Level2Name").value = document.formname.txtLev2.value 
	
		IF document.formname.optMode(0).checked = true then
			TempNode.Item(recno).Attributes.getNamedItem("ModeType").value = "D"
		elseif document.formname.optMode(1).checked = true then 
			 TempNode.Item(recno).Attributes.getNamedItem("ModeType").value = "A" 
		elseif document.formname.optMode(2).checked = true then 
			 TempNode.Item(recno).Attributes.getNamedItem("ModeType").value = "S" 
		elseif document.formname.optMode(3).checked = true then 
			 TempNode.Item(recno).Attributes.getNamedItem("ModeType").value = "N" 
		end if
	
		TempNode.Item(recno).Attributes.getNamedItem("AccHead").value = document.formname.hAccHead.value       
	
		IF TempNode.Item(recno).Attributes.getNamedItem("ModeType").value = "A" then
			TempNode.Item(recno).Attributes.getNamedItem("AccHeadName").value = document.formname.txtAcHead.value       
		End if
			TempNode.Item(recno).Attributes.getNamedItem("FinYear").value = document.formname.FinYear.value 
		IF document.formname.optCompMode(0).checked = true then
			   TempNode.Item(recno).Attributes.getNamedItem("ComputeMode").value = "+"
		Else
			   TempNode.Item(recno).Attributes.getNamedItem("ComputeMode").value = "-"
		End if
		
		objhttp.Open "Post","XMLSchHeadSave.asp?Name=SchedSubHeads&Mod=Acc", false
		objhttp.send XMLData.XMLDocument
	
		If objhttp.responseText <> "" then
			alert(objhttp.responseText)
		Else
			window.returnvalue = "Y"
			window.close()
		End IF
	Else
		Exit Function
	End IF
End Function

Function CheckVal()
	IF document.formname.selSch.value = 0 Then
		MsgBox "Select Schedule "
		document.formname.selSch.focus()
		CheckVal = False
		Exit Function
	End IF
	
	IF document.formname.selLevel1.selectedIndex = 0 and document.formname.txtLev1.value ="" Then
		MsgBox "Select Level 1 "
		document.formname.selLevel1.focus()
		CheckVal = False
		Exit Function
	End IF
	
	'IF document.formname.optLevel(1).checked = True and document.formname.selLevel2.selectedIndex = 0 Then
	If document.formname.txtLev2.value ="" and document.formname.selLevel2.selectedIndex = 0 and document.formname.selLevel2.disabled<>True Then     
		MsgBox "Select Level 2 "
		document.formname.selLevel2.focus()
		CheckVal = False
		Exit Function
	End IF
	
	IF document.formname.optMode(1).checked = True and document.formname.hAccHead.value = 0 Then
		MsgBox "Select Account Head "
		CheckVal = False
		Exit Function
	End IF
	CheckVal = True	
End Function


Function Del()
	Dim objhttp,Node1
	Dim sschedno,SubID,SubSubID,iArr1,iArr2,id,sOrgID,sFinyr
	Set objhttp = CreateObject("Microsoft.XMLHTTP")
	Set Root = TempData.documentElement

	sschedno = document.formname.sschedno.value  
	sOrgID = document.formname.hOrgId.Value
	sFinyr = document.formname.sfinyr.value  
	If document.formname.selLevel2.length>1 and document.formname.selLevel2.value<>"A" Then     
		 iArr2 = split(document.formname.selLevel2.value,"-")
		 SubSubID = iArr2(0) 
		 iArr1 = Split(document.formname.selLevel1.value,"-")
		 SubID = iArr1(0)      
	Else
		iArr = Split(document.formname.selLevel1.value,"-")     
		subID = iArr(0)
		SubSubID = iArr(1)
	End If
		id = "1"

	set Node1 = TempData.CreateElement("Schedule")
		Node1.setAttribute "id",id
	 	Node1.setAttribute "sschedno",sschedno
	 	Node1.setAttribute "SubID",SubID 
	 	Node1.setAttribute "SubSubID",SubSubID 
		Node1.setAttribute "sOrgID",sOrgID
		Node1.setAttribute "sFinyr",sFinyr 
	Root.Appendchild Node1
	MsgBox TempData.XML
	
	objhttp.Open "Post","XMLShdDelete_Update.asp?Name=SchdDelete&Mod=Acc", false
	objhttp.send TempData.XML
	If objhttp.responseText <> "" then
		alert(objhttp.responseText)
	Else
		window.returnvalue = "Y"
		window.close()
	End IF	
End Function
Function EditFields()
	If document.formname.selLevel1.value <>"A" Then 
		document.formname.txtLev1.disabled = False  
		document.formname.txtLev1.value = document.formname.selLevel1.options(document.formname.selLevel1.selectedIndex).text    
		document.formname.txtLev1.size = Len(Trim(document.formname.selLevel1.options(document.formname.selLevel1.selectedIndex).text)) + 7
	End If
	If document.formname.selLevel2.value <>"A" Then 
		document.formname.txtLev2.disabled = False  
		document.formname.txtLev2.value = document.formname.selLevel2.options(document.formname.selLevel2.selectedIndex).text    
		
	End If
		document.formname.btnEdit.disabled = True  
End Function

</Script>
<script language="javascript">
window.__itmsPopupCompat = { type: "scheduleSubHeadsPopup" };
</script>
<script language="javascript" src="../../scripts/itms-modern-compat.js"></script>
<script language="javascript" src="../../scripts/PopupModernCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0" >
<form method="POST" name="formname" >
<input type=hidden name=hSubHeadName value=""><input type=hidden name="hOrgId" value="<%=sOrgId%>">
<input type=hidden name="sschedno" value="<%=sschedno%>">
<input type=hidden name="sfinyr" value="<%=sfinyr%>">
<input type=hidden name="scatcode" value="<%=scatcode%>">
<input type=hidden name="hInsDate" value="<%=sInsDate%>">

<input type=hidden name="hAccHead" value="0">

<table border="0" width="100%" cellspacing="0" cellpadding="0" class="popupTable">
	<tr>
		<td align="center" class=PageTitle height="20"><p align="center">Add Schedule Sub Heads
          </td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
		<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="720">
				<TR>
					<TD class=TabBodyWithTopLine>
                        <table border="0" width="100%" cellspacing="0" cellpadding="0">
                          <tr>
                            <td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td width="100%">
						<table border="0" cellpadding="0" cellspacing="0">
                             <tr>
								<td class="FieldCellSub">Select Schedule</td>
                                <td class="FieldCellSub">
								
									<%
								
									sql = "SELECT ScheduleID, ScheduleHeading, FinYear FROM dbo.Acc_M_SchdSetupHeads Where ScheduleID = "&sSchedNo
									with Objrs1
										.CursorLocation = 3
										.CursorType = 3
										.ActiveConnection = con
										.Source = sql
										.Open
									End With
									Set Objrs1.ActiveConnection = Nothing 
										
									If not Objrs1.EOF Then
									%>					
											<input type=text name="selSch1" size=25 class="FormelemRead" align="Right" value="<%Response.Write(Objrs1(1))%>" readonly>
											<input type="Hidden" name="SelSch" size=25 class="Formelem" align="Right" value="<%Response.Write(Objrs1(0))%>">
									<%
									
									End IF
									Objrs1.Close
									
									%>
									
									</select>
									</td>
								  </tr>
							  <tr>
								<td class="FieldCellSub" valign="top">Add/Modify Level</td>
										<td class="FieldCell" valign="top">
											<Input type="radio" name="optLevel" value="Lev3" class="FormElem"  onclick="LevelFun()" Checked>Level1
											<Input type="radio" name="optLevel" value="Lev4" class="FormElem" onclick="LevelFun()">Level2&nbsp;
										</td>
								</tr>
							  <tr>
								<td class="FieldCellSub">Select Level 1</td>
										<td class="FieldCellSub">
											<select size="1" name="selLevel1" class="FormElem" OnChange="SetLevelFun()">
											<option Value="0">Select</option>
											<%
											sql = "Select Distinct ScheduleSubID,ScheduleSubSubID,isNull(SubHeadingName,'') "&_
													" SubHeadingName,EntryType,computemode,Hierarchy From Vw_Acc_SchSetup Where  OrganisationCode = '"&sOrgId&"' and "&_
													" FinYear = '"&sFinyr&"' and scheduleID = "&sSchedNo&" and schedulesubsubid = 0 Order By Hierarchy "
						
												Objrs1.Open sql,Con
												Do while not Objrs1.EOF 
											%>
											<option value="<%=Objrs1(0)&"-"&Objrs1(1)&"-"&Objrs1("EntryType")&"-"&Objrs1("computemode")%>"><%Response.Write (Left(Objrs1("SubHeadingName"),70)) %></b></option>
											<%Objrs1.MoveNext 
											loop
											Objrs1.Close
											
											%>
											<option Value="A">Add New</option>
											</select>
										</td>
									</tr>
									
									<tr>
									
									<td class="FieldCellSub">&nbsp;</td>
									<td class="FieldCellSub">
										<input type=text name=txtLev1 size=25 class="Formelem" align="Right" >
									</td>
									</tr>
									
								  <tr>
								<td class="FieldCellSub">Select Level 2</td>
										<td class="FieldCellSub">
										<select size="1" name="selLevel2" class="FormElem" disabled onchange="setlevelfun1()">
										<OPTION Value="0">Select</option>
										<OPTION Value="A">Add New</option>
									</select>
									
									</td>
								</tr>
								
								<tr>
									<td class="FieldCellSub">&nbsp;</td>
									<td class="FieldCellSub">
										<input type=text name=txtLev2 size=25 class="Formelem" align="Right" disabled>
									</td>
								</tr>
								
								<tr>
								<td class="FieldCellSub">Mode</td>
								<td class="FieldCell">
									<Input type=radio name=optMode value="D" class="FormElem" OnClick="ModeFun()">
                                    Data Entry
									<Input type=radio name=optMode value="A3" class="FormElem" OnClick="ModeFun()">
                                    A/c Heads
									<Input type=radio name=optMode value="A4" class="FormElem" OnClick="ModeFun()">
                                    Schedule
									<Input type=radio name=optMode value="A5" class="FormElem" checked OnClick="ModeFun()">
                                    Not Applicable
								</td>
								</tr>
								<tr>
								<td class="FieldCellSub">Select A/c Head</td>
								<td class="FieldCellSub">
								<Input type="Button" name="ButAcHead" value="A/cHead" class="ActionButton"  disabled OnClick="popAccList()" >
                                &nbsp;<Input type="text" name="txtAcHead" size="25" class="formElem" readonly value="" ></td>
								</tr>
								<tr>
								<td class="FieldCellSub">FinYear</td>
										<td class="FieldCellSub">
											<select size="1" name="FinYear" class="FormElem">
											<Option Value="<%=sFinyr%>" Selected><%=sFinyr%></Option> 
											</select>
										</td>
								</tr>
								<tr>
								<td class="FieldCellSub">Compute Mode</td>
								<td class="FieldCell">
									<Input type=radio name=optCompMode value="+" class="FormElem" onclick="" checked>Add
									<Input type=radio name=optCompMode value="-" class="FormElem" onclick="">Less
								</td>
								
								</tr>
		
						</table>
                            </td>
                            <td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                          </tr>
                          <tr>
                            <td colspan="3" class="MiddlePack"></td>
                          </tr>
                          <tr>
                            <td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                            <td>
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
										<tr>
											<td valign="middle" class="ActionCell">
                                                <p align="center"> 
												<Input type="Button" name="btnSave" value="Save" class="ActionButton" onclick = "SaveXML()" >&nbsp;
                                                <Input type="Button" name="btnDelete" value="Delete" class="ActionButton" onclick = "Del()" >&nbsp;
                                                <Input type="Button" name="btnClose" value="Close" class="ActionButton" onclick ="window.close()">
											</td>
										</tr>
									</table>
                            </td>
                            <td>
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
                            </td>
                          </tr>
                          <tr>
                            <td colspan="3" class="MiddlePack"></td>
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

