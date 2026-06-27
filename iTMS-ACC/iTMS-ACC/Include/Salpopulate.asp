<%'--------------------------------- Sales Related -----------------------------------%>
<%
'To Align the Text
Function Textalign(val,alen,str1)
	Dim vlen,str2,k
	vlen = CInt(len(val))
	if (vlen > alen) then
		val = mid(val,1,alen)
		vlen = CInt(len(val))
	end if
	k = (alen - vlen)
	if alen = vlen then
		str2 = val
		Textalign= str2
	elseif (str1="L") then
		str2 = val & String(k," ")
		Textalign= str2
	elseif (str1 = "R") then
		str2 = String(k," ") & val
		Textalign= str2
	end if
End function

%>
<%
Function ReplaceTest(Str1, ReplCharacter)
	  Dim Str2, ReplCtr
	  Str2 = Replace(Str1, "'", ReplCharacter)

	  ReplaceTest = Str2

End Function
%>

<%
	' Function to populate the UoM list
	Function populateUoM(Val)
		' Declaration of variables
		Dim oDom,fs,Root,PGNode
		dim sUoMID,sUoMName,sUoMShName
		Set oDOM = Server.CreateObject("Microsoft.XMLDOM")
		Set fs = CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(Server.MapPath("../../inventory/xmldata/UoM.xml")) then
		Response.Write "SD"
			oDOM.Load server.MapPath("../../inventory/xmldata/UoM.xml")
			Set Root = oDOM.documentElement
			if Root.HaschildNodes() then

				For Each PGNode In Root.childNodes
					sUoMID = trim(PGNode.Attributes.Item(0).nodeValue)
					sUoMName = trim(PGNode.Attributes.Item(1).nodeValue)
					sUoMShName = trim(PGNode.Attributes.Item(2).nodeValue)
					if sUoMID = Trim(Val) then
					Response.Write("<OPTION VALUE="""&trim(sUoMID)&""" Selected>"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
					else
					Response.Write("<OPTION VALUE="""&trim(sUoMID)&""" >"&trim(sUoMShName)&"</OPTION>" &vbcrlf)
					End if
				next
			end if
		end if
	End Function
%>

<%

	Function getFromFinYearSal()
			getFromFinYearSal="042003"
			'getFinancialYear=session("FromfinYear")
	End Function

	Function getToFinYearSal()
			getToFinYearSal="032004"
			'getFinancialYear=session("TofinYear")
	End Function
%>
<%
	' Function to populate the Item Type list
	Function popSelItemType(sel)
		' Declaration of variables
		Dim dcrs,stypID,stypName,sTypeNo
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.Source = "SELECT ITEMTYPEID,ITEMTYPENAME,ITEMTYPENO FROM INV_M_ITEMTYPE ORDER BY ITEMTYPEID"
			.ActiveConnection = con
			.Open
		end with
		set stypID = dcrs(0)
		set stypName = dcrs(1)
		set sTypeNo = dcrs(2)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
				If cstr(sel) = cstr(stypID) OR cstr(sel) = cstr(sTypeNo) Then
					Response.Write("<OPTION VALUE="""&trim(stypID)&""" Selected>"&trim(stypName)&"</OPTION>" &vbcrlf)
				Else
					Response.Write("<OPTION VALUE="""&trim(stypID)&""">"&trim(stypName)&"</OPTION>" &vbcrlf)
				End if
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close
		set dcrs.ActiveConnection = nothing
	End Function
%>
<%
	' Function to populate Purchase Type list
	Function popSelSaleTypes(sel)
		' Declaration of variables
		Dim dRSet,sPurTypeName,iPurTypeNum

		'Declaration of Objects
		Set dRSet = Server.CreateObject("ADODB.RecordSet")
		with dRSet
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT InvoiceType,InvTypeShortName,InvoiceTypeName FROM Sal_M_InvoiceTypes ORDER BY InvoiceType"
			.ActiveConnection = con
			.Open
		end with
		set dRSet.ActiveConnection = nothing

		set iPurTypeNum = dRSet(0)
		set sPurTypeName = dRSet(2)
		Do While Not dRSet.EOF
			if cstr(sel) = cstr(iPurTypeNum) then
				Response.Write("<OPTION VALUE="""&trim(cstr(iPurTypeNum))&""" Selected>"&trim(sPurTypeName)&"</OPTION>" &vbcrlf)
			else
				Response.Write("<OPTION VALUE="""&trim(cstr(iPurTypeNum))&""">"&trim(sPurTypeName)&"</OPTION>" &vbcrlf)
			end if
			dRSet.MoveNext
		Loop
		dRSet.Close
	End Function
%>
<%
	' Function to populate the Units list
	Function populateSelUnit(Val)
		' Declaration of variables
		Dim dcrs,sUnitID,sUnitName
		'Declaration of Objects
		Set dcrs = Server.CreateObject("ADODB.RecordSet")
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = "SELECT OUDEFINITIONID,ORGUNITDESCRIPTION FROM DCS_ORGANIZATIONUNITDEFINITIONS WHERE LEN(OUDEFINITIONID) > 4 ORDER BY OUDEFINITIONID"
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		set sUnitID = dcrs(0)
		set sUnitName = dcrs(1)
		If not dcrs.EOF then
			Do While Not dcrs.EOF
			if Trim(sUnitID) = Trim(val) then
				Response.Write("<OPTION VALUE="""&trim(sUnitID)&""" Selected>"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			Else
				Response.Write("<OPTION VALUE="""&trim(sUnitID)&""">"&trim(sUnitName)&"</OPTION>" &vbcrlf)
			End if
				dcrs.MoveNext
			Loop
		end if
		dcrs.Close

	End Function
%>
<%

Function IIf(AE)
		If AE = 1 then
			IIf = " "
		else
			IIf = "s "
		End if
End Function

Function AmountWordsSal(Amount)

    Dim Paise
    ReDim Ones(20)
    ReDim Tens(10)
    Dim Hundred
    ReDim Ws(5)

    Dim A1, S1, Crt, T1
    Dim ReturnStr
	Amount=cdbl(Amount)
    Paise = Amount - Int(Amount)
    Ones(1) = "One "
    Ones(2) = "Two "
    Ones(3) = "Three "
    Ones(4) = "Four "
    Ones(5) = "Five "
    Ones(6) = "Six "
    Ones(7) = "Seven "
    Ones(8) = "Eight "
    Ones(9) = "Nine "
    Ones(10) = "Ten "
    Ones(11) = "Eleven "
    Ones(12) = "Twelve "
    Ones(13) = "Thirteen "
    Ones(14) = "Fourteen "
    Ones(15) = "Fifteen "
    Ones(16) = "Sixteen "
    Ones(17) = "Seventeen "
    Ones(18) = "Eighteen "
    Ones(19) = "Nineteen "
    Tens(1) = "Ten "
    Tens(2) = "Twenty "
    Tens(3) = "Thirty "
    Tens(4) = "Fourty "
    Tens(5) = "Fifty "
    Tens(6) = "Sixty "
    Tens(7) = "Seventy "
    Tens(8) = "Eighty "
    Tens(9) = "Ninety "
    Hundred = "Hundred"
    Ws(1) = "Crore"
    Ws(2) = "Lakh"
    Ws(3) = "Thousand"

    A1 = Int(Amount)
    Crt = 9999999
    S1 = 1
    'ReturnStr = "Rupees "

    Do While A1 > 999
        If A1 > Crt Then
            T1 = Int(A1 / (Crt + 1))
            ReturnStr = ReturnStr + Some_Pro(T1, Ones, Tens)
            ReturnStr = ReturnStr + Ws(S1) + IIf(T1)
            A1 = Int(A1 Mod (Crt + 1))
        End If
        Crt = Int((Crt Mod (Crt + 1)) / 100)
        S1 = S1 + 1
    Loop

    If A1 > 99 Then
        T1 = Int(A1 / 100)
        'MsgBox T1
        ReturnStr = ReturnStr + Ones(T1) + Hundred + IIf(T1)
        A1 = A1 Mod 100
    End If

    If A1 > 0 Then
        ReturnStr = ReturnStr + Some_Pro(A1, Ones, Tens)
        'DO SOME_PRO WITH A1, returnStr
    End If

    If Int(Amount) > 0 Then  ReturnStr = ReturnStr


    If Paise <> 0 Then
    	If Int(Amount) > 0 Then ReturnStr = ReturnStr + "And "

		ReturnStr = ReturnStr + "Paise "
        ReturnStr = ReturnStr + Some_Pro(Paise * 100, Ones, Tens)

    End If

    ReturnStr = ReturnStr + "Only"
    AmountWordssal = ReturnStr

End Function

Function Some_Pro(TT1, Ons, Tes)

    Dim SReturnStr
    If TT1 < 20 Then
        SReturnStr = SReturnStr + Ons(TT1)
    Else
        SReturnStr = SReturnStr + Tes(Int(TT1 / 10))
        If TT1 Mod 10 <> 0 Then
            SReturnStr = SReturnStr + Ons(TT1 Mod 10)
        End If
    End If
    Some_Pro = SReturnStr

End Function

Function GetItemName(sItemCode,sClassCode)
	Dim sQuery,sItemName,dcrs
	IF CStr(sClassCode) = "0" Then
		sQuery = "Select ItemDescription From Ms_TemporaryItemMaster Where TempItemCode = "&sItemCode&" "
	Else
		sQuery = "Select ItemDescription From vwItem Where ItemCode = "&sItemCode&" and ClassificationCode = "&sClassCode&" "
	End IF
	
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	With dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	IF not dcrs.EOF Then
		sItemName = dcrs(0)
	End IF
	dcrs.close
	GetItemName = sItemName 
End Function

Function GetPartyName(sPartyCode)
	Dim sQuery,sPartyName,dcrs
	Response.Write sQuery
	sQuery = "Select PartyName From App_M_PartyMaster Where PartyCode = "&sPartyCode&" "
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	With dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	IF not dcrs.EOF Then
		sPartyName = dcrs(0)
	End IF
	dcrs.close
	GetPartyName = sPartyName 
End Function

Function GetTempItemName(sTempCode)
	Dim sQuery,sTempName,dcrs
	sQuery = "Select ItemDescription From Ms_TemporaryItemMaster Where TempItemCode = "&sTempCode&" "
	Set dcrs = Server.CreateObject("ADODB.RecordSet")
	With dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = sQuery
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	IF not dcrs.EOF Then
		sTempName = dcrs(0)
	End IF
	dcrs.close
	GetTempItemName = sTempName 
End Function
%>