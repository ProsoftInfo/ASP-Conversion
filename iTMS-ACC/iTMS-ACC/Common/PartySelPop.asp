<%@language="VBScript"%>
<%Option Explicit%>
<%
	'Program Name				:	ParSelPop.asp
	'Module Name				:	Admin(Master)
	'Author Name				:	Ragavendran R
	'Created On					:	March 22,2013
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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

<!--#include file="../include/DatabaseConnection.asp"-->
<!--#include file="../include/populate.asp"-->
<!-- #include File="../include/CommonFunctions.asp" -->
<%

Dim sTable
Dim sIType,sOrgID,sFilter,sSearchBy,sSelectMode,sQuery
Dim sFinPeriod,sFinYearFrom,sFinYearTo,sTempMonYr,sMonYr,sEmpID
Dim iCurrentPage,iTotalPage,iPageCtr,lnPage,iCtr,iPageNo,iSNo
Dim dcrs,rsTemp,sTemp,sRequest,sEmpCode,sEmpName

Set dcrs = Server.CreateObject("ADODB.Recordset")
Set rsTemp = Server.CreateObject("ADODB.Recordset")

Const iPageSize = 15
Response.Write "<font color=red>"

sSearchBy = trim(Request.QueryString("SearchBy"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
Response.Write "<p><font color=red>"

sFilter = trim(Request.QueryString("Query"))&"%"

if Trim(sSearchBy)="" or IsNull(sSearchBy) then sSearchBy = "IC"

sRequest = "SearchBy="&sSearchBy&"&Query="&Request("Query")&"&hSelectMode="&sSelectMode

iCurrentPage=Request("Page")
if Trim(iCurrentPage)="" or IsNull(iCurrentPage) then iCurrentPage = 1
iCurrentPage = CInt(iCurrentPage)

if len(Month(date())) = 1 then
	sTempMonYr = "0"&Month(date())
else
	sTempMonYr = Month(date())
end if

sMonYr = sTempMonYr&Year(date())

sFinPeriod = split(Session("FinPeriod"),":") '
sFinYearFrom =  "01/04/"&sFinPeriod(0)       '
sFinYearTo = "31/03/"&sFinPeriod(1)          '

if trim(sSelectMode) = "" then sSelectMode = "R"

iSAApplicationPop = Session("iApplication")
iSAProcessPop = Session("iProcess")
iSAActivityPop = Session("iActivity")
iEmpNoPopulate = Session("employeenumber")

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Employee Selection</TITLE>
<base target="_self"></base>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<xml id="TempItem"><Root CurrPage="1" TotPage="1"></Root></xml>
<xml id="XMLPartySubType"><Root></Root></xml>
<xml id="PartyData"><Root></Root></xml>

<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/DataValidation.js"></SCRIPT>
<Script language="Javascript">
function DoKeyPress(sYesNo,iIntPart,iDecPart) {
	var sIntVal
	sIntVal=""
	eTD = window.event.srcElement;
	
	if (sYesNo == "N") {
		if ((window.event.keyCode < 48 || window.event.keyCode > 57) && window.event.keyCode !=60 && window.event.keyCode !=61 && window.event.keyCode !=62) {
			window.event.keyCode ="\b";
		}
	}
	else if (sYesNo == "Y") {
		if ((window.event.keyCode < 48 || window.event.keyCode > 57) && window.event.keyCode != 46 && window.event.keyCode !=60 && window.event.keyCode !=61 && window.event.keyCode !=62) {
			window.event.keyCode ="\b";
		}
	}
	
	sValue = new String(eTD.value);
	
	iDecPostion = sValue.indexOf(".");
	
	if (iDecPostion >= 0) {
		sDecVal = sValue.substring(iDecPostion + 1,sValue.length);
		sIntVal = sValue.substring(0,iDecPostion);
	}
	else {
		sDecVal="";
		sIntVal = sValue
	}

	if (sYesNo == "N") {
		if (sIntVal.length >= iIntPart)
			window.event.keyCode = "\b";
	}
	else if (sYesNo == "Y") {
		if (iDecPostion >= 0) {
			if (window.event.keyCode == 46 || (sDecVal.length >= iDecPart))
				window.event.keyCode = "\b";
		}
		else {
			if (sIntVal.length = iIntPart) {
				if (sDecVal.length >= iDecPart)
					window.event.keyCode = "\b";
			}
			if ((sIntVal.length >= iIntPart) && window.event.keyCode != 46)
				window.event.keyCode = "\b";
			
		}
		
	}
}	
</Script>

<SCRIPT LANGUAGE=vbscript>
dim sRet,sButtonPressed
sRet = "-1:0"
'**************************************************
Function AddFun()

	Dim Root,node1
	Dim n1,Arr1,sSelectMode,nItemRate,nItemStock
	set Root = PartyData.DocumentElement
	
	IF Root.haschildnodes then
		For each node in Root.childnodes
			if node.nodename = "Entry" then
				iCount = cint(Root.childNodes.length) + 1
			else
				iCount = 1
			end if
		Next
	else
		iCount = 1
	End IF
	For i = 1 to document.formname.hChkCount.value

		sChkObj = eval("document.formname.CHKZ"&i).checked
		
		If sChkObj = True  then
		    
			set Obj  = eval("document.formname.CHKZ"&i)
			
			sAttribVal = ""
			sValue = ""
			nText =""
	
	
	            sSelectMode = document.formname.hSelectMode.value
	            n1 = Obj.value
	            
	            Arr1 = split(n1,":")

                    set ndRoot =  TempItem.documentElement
		            sExp = "//Emp[@EmpID="&Arr1(1)&"]" 
		            set ndPartyNode = ndRoot.selectNodes(sExp)
		            sPartyName = ndPartyNode.Item(0).Attributes.getNamedItem("EmpFullName").value
		            sEmpName = ndPartyNode.Item(0).Attributes.getNamedItem("EmpName").value
                
	            if sSelectMode = "M" then
		            if Obj.checked then
		                for each temp in Root.childnodes
				            if Strcomp(temp.nodename,"Entry")=0 then
					            if trim(Temp.getAttribute("RetField1")) =  trim(Arr1(1)) then
						            exit function
					            end if
				            end if
			            next
			            
			            set node1 = PartyData.createElement("Entry")
			                node1.SetAttribute "RetField0", sPartyName
			                node1.SetAttribute "RetField1",Arr1(1) ' EmpID
			                node1.SetAttribute "RetField2",Arr1(0) 'EmpCode
			                node1.SetAttribute "RetField3",sEmpName 'EmpSpliTName
			                Root.appendchild node1
			            
		            else
			            for each temp in Root.childnodes
				            if Strcomp(temp.nodename,"Entry")=0 then
					            if trim(Temp.getAttribute("RetField1")) =  trim(Arr1(1)) then
						            Root.Removechild temp
					            end if
				            end if
			            next
		            end if

		            DispList()
	            else

		            for each temp in Root.childnodes
			            if Strcomp(temp.nodename,"Entry")=0 then
				            Root.Removechild temp
			            end if
		            next

		            if Obj.checked then
			            iCount  = iCount + 1

			           	set node1 = PartyData.createElement("Entry")
			                node1.SetAttribute "RetField0", sPartyName
			                node1.SetAttribute "RetField1",Arr1(1) ' EmpID
			                node1.SetAttribute "RetField2",Arr1(0) 'EmpCode
			                node1.SetAttribute "RetField3",sEmpName 'EmpSpliTName
			                Root.appendchild node1
		                
		            end if
		            DispList()
		            
	            end if
	        eval("document.formname.CHKZ"&i).checked = false
	        
	    end if 'If sChkObj = True  then
    Next 'For i = 1 to document.formname.hChkCount.value
    
	 'alert(Root.xml)
End Function
'********************************************************************************
Function XmlFun(Obj)

	Dim Root,node1,ndRoot,ndPartyNode
	Dim n1,Arr1,sSelectMode,nItemRate,nItemStock
	set Root = PartyData.DocumentElement
	'alert(Root.xml)

	IF Root.haschildnodes then
		For each node in Root.childnodes
			if node.nodename = "Entry" then
				iCount = cint(Root.childNodes.length) + 1
			else
				iCount = 1
			end if
		Next
	else
		iCount = 1
	End IF

	sSelectMode = document.formname.hSelectMode.value
	
	n1 = Obj.value
	Arr1 = split(n1,":")
	
	

		
		set ndRoot =  TempItem.documentElement
		sExp = "//Emp[@EmpID="&Arr1(1)&"]" 
		set ndPartyNode = ndRoot.selectNodes(sExp)
		sPartyName = ndPartyNode.Item(0).Attributes.getNamedItem("EmpFullName").value
		sEmpName = ndPartyNode.Item(0).Attributes.getNamedItem("EmpName").value
		
	if sSelectMode = "M" then
		if Obj.checked then
		    for each temp in Root.childnodes
				if Strcomp(temp.nodename,"Entry")=0 then
					if trim(Temp.getAttribute("RetField1")) =  trim(Arr1(1)) then
						exit function
					end if
				end if
			next

			set node1 = PartyData.createElement("Entry")
			node1.SetAttribute "RetField0", sPartyName
			node1.SetAttribute "RetField1",Arr1(1) ' EmpID
			node1.SetAttribute "RetField2",Arr1(0) 'EmpCode
			node1.SetAttribute "RetField3",sEmpName 'EmpSpliTName
			Root.appendchild node1
			
		else
			for each temp in Root.childnodes
				if Strcomp(temp.nodename,"Entry")=0 then
					if trim(Temp.getAttribute("RetField1")) =  trim(Arr1(1))then
						Root.Removechild temp
					end if
				end if
			next
		end if

		DispList()
	else

		for each temp in Root.childnodes
			if Strcomp(temp.nodename,"Entry")=0 then
				Root.Removechild temp
			end if
		next

		if Obj.checked then
			iCount  = iCount + 1

			set node1 = PartyData.createElement("Entry")
			node1.SetAttribute "RetField0", sPartyName
			node1.SetAttribute "RetField1",Arr1(1) ' EmpID
			node1.SetAttribute "RetField2",Arr1(0) 'EmpCode
			node1.SetAttribute "RetField3",sEmpName 'EmpSpliTName
			Root.appendchild node1
		    
		end if
		DispList()
	end if
	 'alert(Root.xml)
End Function
'********************************************************************************
Function DispList()
Dim s1

	s1 = "<br><TABLE class=""TableOutLineOnly"" cellspacing=""1"" width=""100%"">"
	set Root = PartyData.DocumentElement
		sQ = Root.getAttribute("PassQuery")
		sIType = right(sQ,3)
		'alert(ROot.xml)
	if Root.haschildnodes then
		for each temp in Root.childnodes
			if trim(temp.nodename) = trim("Entry") then
				s1= trim(s1) & "<tr><td class=ExcelDisplayCell >"
				s1= trim(s1) & "<input type=checkbox name=chk value='" & trim(temp.getAttribute("RetField1")) & "' checked onClick=RemoveNode(this)>"
				s1= trim(s1) & "</td>"
				s1= trim(s1) & "<td class=ExcelDisplayCell >" & trim(temp.getAttribute("RetField2")) & "</td>"
				s1= trim(s1) & "<td class=ExcelDisplayCell >" & replace(trim(temp.getAttribute("RetField0")),"~~",chr(34)) & "</td>"
				s1= trim(s1) & "</tr>"
			end if 'if Strcomp(temp.nodename,"Party")= 0 then
		next
	end if 'if Root.haschildnodes then
	
	s1 = trim(s1) + "</table><br>"
	idSelList.innerHTML = s1
End Function
'********************************************************************************
Function SendValue()
	sButtonPressed = "Done"
	set Root = PartyData.DocumentElement
	'alert(Root.xml)
	Root.SetAttribute "Action","Done"
	window.close
End Function
'********************************************************************************
Function window_onunload()
    set window.returnValue  = PartyData.documentelement
End Function
'********************************************************************************
function showpage(sArguments)
Dim sSICode,sSIName,sSPICode,sSPIName
Dim sSelectMode,objhttp,ndRoot,ndItem,iSLNO,iCounter
Dim sCompanyItemCode,sItemName,sClassName,iStock,sUOM,iItemCode,iClassCode
Dim sDecimal,sReceiptNum,sAttribList,sParItemCode,sParItemDesc
Dim iItemRate,sLocNo,sBinNo,sLocBinCount,iMarketPrice
Dim iTotPage,iCurrPage,iTabIndex
sSelectMode = document.formname.hSelectMode.value 
	EmpCode = document.formname.txtEmpCode.value
	EmpName = document.formname.txtEmpName.value
	
	sRet = sArguments&"&EmpCode="&EmpCode&"&EmpName="&EmpName
	sButtonPressed = "Page"
	set Root = PartyData.DocumentElement
	Root.SetAttribute "Action","Page"
	Root.SetAttribute "PassQuery",sRet
	
	
        set objhttp = CreateObject("Microsoft.XMLHTTP")
        sTempValues = document.formname.hTemp.value
        objhttp.open "GET","XMLGetParSel.asp?"&sRet,false
        objhttp.send
        
        if Trim(objhttp.responseXML.xml)<>"" then
            TempItem.loadXML(objhttp.responseXML.xml)
        else
            alert(objhttp.responseText)
        end if
        
        
    iTabIndex = 6
        set ndRoot = TempItem.documentElement
        'alert(ndRoot.xml)
        if ndRoot.hasChildNodes() then
        ClearTable()
            iCurrPage = ndRoot.getAttribute("CurrPage")
            iTotPage = ndRoot.getAttribute("TotPage")
            for each ndParty in ndRoot.childNodes
                if ndParty.nodeName="Emp" then
                    iSLNO = ndParty.getAttribute("SNo")
                    iCounter = ndParty.getAttribute("Counter")
                    sEmpCode = ndParty.getAttribute("EmpCode")
                    sEmpID =ndParty.getAttribute("EmpID")
                    sEmpName = ndParty.getAttribute("EmpFullName")
                                                
                    set iTblRow = document.all.tblItem.insertRow(document.all.tblItem.rows.length-3)
                    
                    set CurrCell = iTblRow.InsertCell()
			        if Trim(sSelectMode)="M" then
			            CurrCell.innerHtml = "<input type=checkbox name=ChkZ"& iSLNO &" value="& sEmpCode &":"& sEmpID &">"
			        else
			            CurrCell.innerHtml = "<input type=Radio name=ChkZ"& iCounter &" value="& sEmpCode &":"& sEmpID &" onclick=XMLFun(this)>"
			        end if
        	        CurrCell.ClassName = "ExcelDisplayCell"
			        CurrCell.Align	   = "center"
        			
			        set CurrCell = iTblRow.InsertCell()
			        CurrCell.innerHtml = sEmpCode
			        CurrCell.ClassName = "ExcelDisplayCell"
			        CurrCell.Align		= "Left"
        			
			        set CurrCell = iTblRow.InsertCell()
			        CurrCell.innerHtml = sEmpName
			        CurrCell.ClassName = "ExcelDisplayCell"
			        CurrCell.Align	   = "Left"
        			
	            end if 'if ndParty.nodeName="Party" then
            next
            document.formname.txtCurrPage.tabIndex = iTabIndex
            document.formname.txtCurrPage.value = iCurrPage
            
            iTabIndex= iTabIndex + 1
            document.formname.btnAddToList.tabIndex = iTabIndex
            iTabIndex= iTabIndex + 1
            document.formname.hChkCount.value = iSLNO 
            document.formname.btnDone.tabIndex = iTabIndex
            document.formname.hPage.value  = iCurrPage
            spanTotPage.innerText = iTotPage
        else
        ClearTable()
            set iTblRow = document.all.tblItem.insertRow(document.all.tblItem.rows.length-3)
            set CurrCell = iTblRow.InsertCell()
            CurrCell.innerText = "No Records Found"
	        CurrCell.ClassName = "ExcelDisplayCell"
	        CurrCell.Align	   = "center"
	        CurrCell.colspan = 7
	        document.formname.txtCurrPage.tabIndex = iTabIndex
	        document.formname.txtCurrPage.value = "0"
	        
	        iTabIndex= iTabIndex + 1
            document.formname.btnAddToList.tabIndex = iTabIndex
            iTabIndex= iTabIndex + 1
            document.formname.btnDone.tabIndex = iTabIndex
            document.formname.hPage.value  ="0"
            spanTotPage.innerText = "0"
        end if'if ndRoot.hasChildNodes() then 
        
    'end if 'if Trim(sSICode)<>"" or Trim(sSIName)<>"" or Trim(sSPICode)<>"" or Trim(sSPIName)<>"" then
end function
'-----------
Function CallSearchMain()
Dim TheKey,sRequest,sPage,sLastPage
TheKey = window.event.keyCode 
sRequest =  document.formname.hRequest.value 
sPage = document.formname.hPage.value 
sLastPage = CDbl(spanTotPage.innerText)
    if TheKey = 13 then
        sPage = document.formname.txtCurrPage.value 
        showpage(sRequest&"&Page="&sPage)
    elseif TheKey = 33 then
        if sPage > 1 then sPage = sPage - 1
        showpage(sRequest&"&Page="&sPage)
    elseif TheKey = 34 then
        if cdbl(sPage) < CDbl(sLastPage) then sPage = cdbl(sPage) + 1
        showpage(sRequest&"&Page="&sPage)
    end if
End Function
'--------------------------------
Function CallSearch()
Dim TheKey,sRequest,sPage,sLastPage
TheKey = window.event.keyCode 
sRequest =  document.formname.hRequest.value 
sPage = document.formname.hPage.value 
sLastPage = CDbl(spanTotPage.innerText)
'    if TheKey = 13 then
        sPage = document.formname.txtCurrPage.value 
        showpage(sRequest&"&Page="&sPage)
'    end if
End Function
'------------------------------------
'-----------------------------
Function Init()
Dim sSelectMode,objhttp,ndRoot,ndItem,iSLNO,iCounter
Dim sCompanyItemCode,sItemName,sClassName,iStock,sUOM,iItemCode,iClassCode
Dim sDecimal,sReceiptNum,sAttribList,sParItemCode,sParItemDesc
Dim iItemRate,sLocNo,sBinNo,sLocBinCount,iMarketPrice
Dim iTotPage,iCurrPage
set objhttp = CreateObject("Microsoft.XMLHTTP")
sSelectMode = document.formname.hSelectMode.value 

sTempValues = document.formname.hTemp.value
objhttp.open "GET","XMLGetParSel.asp?"&sTempValues,false
objhttp.send
'alert(objhttp.responseText)
if Trim(objhttp.responseXML.xml)<>"" then
    TempItem.loadXML(objhttp.responseXML.xml)
else
    alert(objhttp.responseText)
end if

set ndRoot = TempItem.documentElement
if ndRoot.hasChildNodes() then
ClearTable()
    iCurrPage = ndRoot.getAttribute("CurrPage")
    iTotPage = ndRoot.getAttribute("TotPage")
    for each ndParty in ndRoot.childNodes
        if ndParty.nodeName="Emp" then
            iSLNO = ndParty.getAttribute("SNo")
            iCounter = ndParty.getAttribute("Counter")
            sEmpCode = ndParty.getAttribute("EmpCode")
            sEmpName = ndParty.getAttribute("EmpFullName")
            sEmpID = ndParty.getAttribute("EmpID")
                                        
            set iTblRow = document.all.tblItem.insertRow(document.all.tblItem.rows.length-3)
            
            set CurrCell = iTblRow.InsertCell()
			if Trim(sSelectMode)="M" then
			    CurrCell.innerHtml = "<input type=checkbox name=ChkZ"& iSLNO &" value="& sEmpCode &":"& sEmpID &" >"
			else
			    CurrCell.innerHtml = "<input type=Radio name=ChkZ"& iCounter &" value="& sEmpCode &":"& sEmpID &" onclick=XMLFun(this)>"
			end if
        	CurrCell.ClassName = "ExcelDisplayCell"
			CurrCell.Align	   = "center"
			
			set CurrCell = iTblRow.InsertCell()
			CurrCell.innerHtml = sEmpCode
			CurrCell.ClassName = "ExcelDisplayCell"
			CurrCell.Align		= "Left"
			
			set CurrCell = iTblRow.InsertCell()
			CurrCell.innerHtml = sEmpName
			CurrCell.ClassName = "ExcelDisplayCell"
			CurrCell.Align	   = "Left"
			
	    end if 'if ndParty.nodeName="Party" then
    next
end if'if ndRoot.hasChildNodes() then 

document.formname.hChkCount.value = iSLNO 
document.formname.txtCurrPage.value = iCurrPage
document.formname.hPage.value  = iCurrPage
spanTotPage.innerText = iTotPage
DispList
End Function
'-----------------------------
Function ClearTable()
	Dim iNum
	For iNum = 2 to document.all.tblItem.rows.length - 4
        document.all.tblItem.deleteRow(2)
    Next
End Function
'********************************************************************************
function RemoveNode(this)
	Dim Root,node1
	Dim n1,Arr1,PartyCode
	PartyCode = this.value ' PartyCode
	
	set Root = PartyData.DocumentElement
	 if this.checked = false then
		for each temp in Root.childnodes
			if Strcomp(temp.nodename,"Entry")=0 then
				if trim(Temp.getAttribute("RetField1")) =  trim(PartyCode) then
					Root.Removechild temp
				end if
			end if
		next

		for i = 0 to document.FormName.elements.length - 1
			if document.FormName.elements(i).type = "checkbox"   then
				if document.FormName.elements(i).name = "ChkZ"   then
					n1 = trim(document.FormName.elements(i).value)
					'alert(n1)
					TempArr = split(n1,":")
					if trim(PartyCode) = trim(TempArr(1)) then
						document.FormName.elements(i).checked= false
						exit for
					end if 'if trim(Arr1(0)) = trim(TempArr(0)) then
				end if 'if document.FormName.elements(i).name = "pKey"   then
			end if 'if document.FormName.elements(i).type = "checkbox" then
		next
		DispList()
	end if 'if this.checked = false then
end Function
'********************************************************************************
</script>
<script language="javascript">
window.__itmsPartySelectorConfig = {
	kind: "employee",
	dataUrl: "XMLGetParSel.asp"
};
</script>
<script language="javascript" src="../../scripts/PartySelectorCompat.js"></script>
</HEAD>
<BODY leftMargin=0 topMargin=0 onload="Init();" onkeydown="CallSearchMain()">
<form method="POST" name="formname" class="PopupTable">
<input type="hidden" name="hTemp" value="<%=Request.QueryString%>">
<input type="hidden" name="hSelectMode" value="<%=sSelectMode%>">
<input type="hidden" name="hEmpID" value="<%=sEmpID%>">
<input type="hidden" name="hPage" value="0">
<input type="hidden" name="hRequest" value="<%=sRequest%>">
<input type="hidden" name="hChkCount" value="0">

<table border="0" width="98%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
	<tr>
	    <td width="10"></td>
		<td valign="top">
			<table id="tblItem" border="0" cellpadding="0" cellspacing="1" class="ExcelTable" width="100%">
			    <tr>
			        <td class="ExcelHeaderCell" align="center">Select</td>
			        <td class="ExcelHeaderCell" align="center">Employee Code</td>
			        <td class="ExcelHeaderCell" align="center">Employee Name</td>
			    </tr>
			    <tr>
			        <td class="ExcelHeaderCell" align="center"></td>
			        <td class="ExcelHeaderCell" align="center"><input type="text" name="txtEmpCode" class="FormElem" onblur="ShowPage('<%=sRequest%>&Page='+document.formname.hPage.value)" onkeyup="CallSearch()"></td>
			        <td class="ExcelHeaderCell" align="center"><input type="text" name="txtEmpName" class="FormElem" onblur="ShowPage('<%=sRequest%>&Page='+document.formname.hPage.value)" onkeyup="CallSearch()"></td>
			    </tr>
			    <tr>
                    <td valign="top" class="ExcelHeaderCell" align="center" colspan="8">Page&nbsp;
                    <input type=text class="FormElem" size=5 style="text-align:right" name="txtCurrPage" onblur="ShowPage('<%=sRequest%>&Page='+this.value)" onkeydown="CallSearch()" >&nbsp;
                    of&nbsp;<span id="spanTotPage"></span>
                    </td>
                </tr>
                <tr>
                    <td class="ExcelHeaderCell" colspan="8" align="center">
                        <%if sSelectMode ="M" then %>
                            <input type="button" name="btnAddToList" value="Add To List" class="ActionButtonX" onclick="AddFun()">
                            <input type="button" name="btnDone" value="Done" class="ActionButtonX" onclick="SendValue()">
                        <%else %>
                            <input type="button" name="btnAddToList" value="Add To List" disabled class="ActionButtonX">
                            <input type="button" name="btnDone" value="Done" class="ActionButtonX" onclick="SendValue()">
                        <%end if%>
                    </td>
                </tr>
                <tr>
                    <td class="ExcelHeaderCell" colspan="8">
                     Selected Entries<span id="idSelList"></span>
                    </td>
                </tr>
			</table>
		</td>
	</tr>
</table>
</form>
</BODY>
</html>
<%
Private Function PageQuery(ByRef pnPage)
		Dim lsQuery,iCount

		'lsQuery = Request.QueryString
		lsQuery = sRequest

		iCount = InStr(1,lsQuery,"&Query=")
		If cint(iCount) > 0 Then
			lsQuery = left(lsQuery,iCount - 1)
		End If

		lsQuery = Replace(lsQuery, "Page=" & lnPage, "")
		'Response.Write lsQuery

		If pnPage < 1 Then
			pnPage = 1
		ElseIf pnPage > iTotalPage Then
			pnPage = iTotalPage
		End If

		If lsQuery = "" Then
			lsQuery = "Page=" & pnPage
		ElseIf Right(lsQuery, 1) = "&" Then
			lsQuery = lsQuery & "Page=" & pnPage
	   	Else
			lsQuery = lsQuery & "&Page=" & pnPage
		End If

		'PageQuery = "?" & lsQuery
		PageQuery = lsQuery
		'Response.Write PageQuery

	End Function
%>
