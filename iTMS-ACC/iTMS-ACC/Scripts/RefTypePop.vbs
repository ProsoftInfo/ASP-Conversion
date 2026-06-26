' VBScript File
''selRefName,hRefType,hItemType hidden fields
''OutData  XML
Function RefTypeSelectionSupp(sRefType,sOrgID,sPartyCode,iStock,nFlag,bAddButton,sDispItem)
    if sRefType <>"N" then
        Set	ndItems = showModalDialog("/Common/DynamicNoSelection.asp?orgID="&sOrgID&"&RefType="&sRefType&"&ParCode="&sPartyCode,OutData,"dialogHeight:500px;dialogWidth:700px;status:no")
		sAct = UCase(trim(ndItems.getAttribute("Action")))
		sQuery = trim(ndItems.getAttribute("PassQuery"))
		if ucase(trim(sAct)) <> "CLOSE" then
			do while sAct <> "DONE"
				set ndItems = showModalDialog("/Common/DynamicNoSelection.asp?"&sQuery,OutData,"dialogHeight:500px;dialogWidth:700px;status:no")
				sAct = UCase(trim(ndItems.getAttribute("Action")))
				if ucase(Trim(sAct)) = "CLOSE" then exit do
				sQuery = trim(ndItems.getAttribute("PassQuery"))
			loop
		end if
		
		if nditems.xml<>"" then
			OutData.loadXML(nditems.xml)
		end if
    else
	    set OutValue=showModalDialog("../../Common/SuppItemSelectCommon.asp?orgID=" & sOrgID & "&Stock=" & iStock & "&hSelectMode=M&Flag="+cstr(nFlag)&"&hDispButt="&bAddButton&"&hDispItem="&sDispItem&"&hPartyCode="&sPartyCode,OutData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
	    'alert(OutValue.xml)
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    
	    'alert sQuery
	    if ucase(trim(sAct)) <> "CLOSE" then
		    do while sAct <> "DONE"
				    set OutValue=showModalDialog("../../Common/SuppItemSelectCommon.asp?" & sQuery,OutData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
				    sAct = UCase(trim(OutValue.getAttribute("Action")))
				    if ucase(Trim(sAct)) = "CLOSE" then exit do
				    sQuery = trim(OutValue.getAttribute("PassQuery"))
				    
		    loop
	    end if 'if ucase(trim(sAct)) <> "CLOSE" then
	end if 'if sRefType <>"N" then
End Function
'------------------------------------------------
Function RefTypeSelection(sRefType,sOrgID,sPartyCode,iStock,nFlag,bAddButton,sDispItem,sCallFrom)
    if sRefType <>"N" then
		sTempValWindowSize = GetWindowSizeForPopup("3")
		sArrTempValWindowSize = split(sTempValWindowSize,":")
		sProgramName = sArrTempValWindowSize(0)
		sPopupHeight = sArrTempValWindowSize(1)
		sPopupWidth = sArrTempValWindowSize(2)

    
        Set	ndItems = showModalDialog("/Common/"&sProgramName&"?orgID="&sOrgID&"&RefType="&sRefType&"&ParCode="&sPartyCode,OutData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
		sAct = UCase(trim(ndItems.getAttribute("Action")))
		sQuery = trim(ndItems.getAttribute("PassQuery"))
		if ucase(trim(sAct)) <> "CLOSE" then
			do while sAct <> "DONE"
				set ndItems = showModalDialog("/Common/"&sProgramName&"?"&sQuery,OutData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
				sAct = UCase(trim(ndItems.getAttribute("Action")))
				if ucase(Trim(sAct)) = "CLOSE" then exit do
				sQuery = trim(ndItems.getAttribute("PassQuery"))
			loop
		end if
		
		if nditems.xml<>"" then
			OutData.loadXML(nditems.xml)
		end if
    else
		
		sTempValWindowSize = GetWindowSizeForPopup("1")
		sArrTempValWindowSize = split(sTempValWindowSize,":")
		sProgramName = sArrTempValWindowSize(0)
		sPopupHeight = sArrTempValWindowSize(1)
		sPopupWidth = sArrTempValWindowSize(2)

		
        set OutValue = showModalDialog("../../Common/"&sProgramName&"?orgID="& sUnit &"&sIType=" & sIType & "&Stock=" & iStock & "&hSelectMode=M&Flag="+cstr(nFlag)&"&hDispButt="&bAddButton&"&hDispItem="&sDispItem&"&CallFrom="&sCallFrom,OutData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	   ' set OutValue=showModalDialog("../../Common/ItemSelectCommon.asp?orgID=" & sOrgID & "&Stock=" & iStock & "&hSelectMode=M&Flag="+cstr(nFlag)&"&hDispButt="&bAddButton&"&hDispItem="&sDispItem,OutData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
	    sAct = UCase(trim(OutValue.getAttribute("Action")))
	    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    
	    'alert sQuery
	  '  if ucase(trim(sAct)) <> "CLOSE" then
	'	    do while sAct <> "DONE"
	'			    set OutValue=showModalDialog("../../Common/"&sProgramName&"?" & sQuery,OutData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	'			    sAct = UCase(trim(OutValue.getAttribute("Action")))
	'			    if ucase(Trim(sAct)) = "CLOSE" then exit do
	'			    sQuery = trim(OutValue.getAttribute("PassQuery"))
	'	    loop
	 '   end if 'if ucase(trim(sAct)) <> "CLOSE" then
	end if 'if sRefType <>"N" then
End Function
'-----------------------------------------------------------------------------

Function MixSelection(sOrgID)

	set OutValue=showModalDialog("../../Common/MixSelectCommon.asp?orgID=" & sOrgID & "&hSelectMode=M",MixData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
	sAct = UCase(trim(OutValue.getAttribute("Action")))
	sQuery = trim(OutValue.getAttribute("PassQuery"))
	'alert sQuery
	if ucase(trim(sAct)) <> "CLOSE" then
	    do while sAct <> "DONE"
			    set OutValue=showModalDialog("../../Common/MixSelectCommon.asp?" & sQuery,MixData,"dialogHeight:500px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No")
			    sAct = UCase(trim(OutValue.getAttribute("Action")))
			    if ucase(Trim(sAct)) = "CLOSE" then exit do
			    sQuery = trim(OutValue.getAttribute("PassQuery"))
	    loop
	end if 'if ucase(trim(sAct)) <> "CLOSE" then
	    
End Function
'*********************************
Function popIssueTo()''Should declare in calling program hIssueToCode,hIssueToType,hIssueToSubCode as hidden and PartyData as XML
Dim OutValue,ObjValue,IssVal
Dim sTempValWindowSize,sArrTempValWindowSize,sProgramName,sPopupHeight,sPopupWidth
    
    sTempValWindowSize = GetWindowSizeForPopup("2")
    sArrTempValWindowSize = split(sTempValWindowSize,":")
    sProgramName = sArrTempValWindowSize(0)
    sPopupHeight = sArrTempValWindowSize(1)
    sPopupWidth = sArrTempValWindowSize(2)
    
	IssVal = document.formname.selIssueTo(document.formname.selIssueTo.selectedIndex).value
	
	'alert(issval)
	if lcase(trim(IssVal))=lcase("Party") then
		ObjValue = document.formname.hIssueToCode.value

		sOrgID = document.formname.hUnit.value
	    set	OutValue = showModalDialog("/Common/"&sProgramName&"?orgID="&sOrgID,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
	    sQuery = OutValue.getAttribute("PassQuery")
	    if OutValue.getAttribute("Action")="CLOSE" then exit function

		while OutValue.getAttribute("Action")<>"Done"
		set	OutValue = showModalDialog("/Common/"&sProgramName&"?"&sQuery,PartyData,"dialogHeight:"& sPopupHeight &"px;dialogWidth:"& sPopupWidth &"px;Status:No")
		    sQuery = OutValue.getAttribute("PassQuery")
	        if OutValue.getAttribute("Action")="CLOSE" then exit function
		wend
		if OutValue.hasChildNodes() then
		    For each ndChild in OutValue.childNodes
		        txtParty.innerText = ndChild.getAttribute("RetField0")
		        
		        document.formname.hIssueToType.value = "Party"
                document.formname.hIssueToCode.value = ndChild.getAttribute("RetField1")
		    Next
		end if
	elseif lcase(trim(IssVal))=lcase("dept:prd") then
	    set WorkCenter = showModalDialog("/Common/WorkCenterPopup.asp","","dialogHeight:150px;dialogWidth:300px;")
        if WorkCenter.hasChildNodes() then
            for each ndChild in WorkCenter.childNodes
                 sArrIssValue = split(IssVal,":")
                document.formname.hIssueToType.value = sArrIssValue(0)
                document.formname.hIssueToCode.value = sArrIssValue(1)
                
                document.formname.hIssueToSubCode.value = ndChild.getAttribute("Code")
                txtParty.innerText = ndChild.getAttribute("Name")
                exit for
            next
        end if
    elseif lcase(trim(IssVal))=lcase("unit") then
        alert("Select the Sub Level")
        document.formname.selIssueTo.selectedIndex = 0
        exit function
    elseif lcase(trim(IssVal))=lcase("pos") then
        alert("Select the Sub Level")
        document.formname.selIssueTo.selectedIndex = 0
        exit function
    else
        sArrIssValue = split(IssVal,":")
        document.formname.hIssueToType.value = sArrIssValue(0)
        document.formname.hIssueToCode.value = sArrIssValue(1)
        
	end if ' if lcase(trim(IssVal))=lcase("Party") then
End Function
'---------------------------------------------------------------
Function popIssueToWithOutSubLevel()''Should declare in calling program hIssueToCode,hIssueToType,hIssueToSubCode as hidden

    IssVal = document.formname.selIssueTo(document.formname.selIssueTo.selectedIndex).value
	if lcase(trim(IssVal))=lcase("Party") then
		document.formname.hIssueToType.value = "Party"
        document.formname.hIssueToCode.value = ""'ndChild.getAttribute("RetField1")
	elseif lcase(trim(IssVal))=lcase("dept:prd") then
        sArrIssValue = split(IssVal,":")
        document.formname.hIssueToType.value = sArrIssValue(0)
        document.formname.hIssueToCode.value = sArrIssValue(1)
    elseif lcase(trim(IssVal))=lcase("unit") then
		document.formname.hIssueToType.value = "unit"
        document.formname.hIssueToCode.value = ""
    elseif lcase(trim(IssVal))=lcase("pos") then
        document.formname.hIssueToType.value = "pos"
        document.formname.hIssueToCode.value = ""
    else
        sArrIssValue = split(IssVal,":")
        document.formname.hIssueToType.value = sArrIssValue(0)
        document.formname.hIssueToCode.value = sArrIssValue(1)
	end if ' if lcase(trim(IssVal))=lcase("Party") then
End Function