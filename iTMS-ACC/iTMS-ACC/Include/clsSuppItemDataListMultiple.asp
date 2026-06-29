
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"

%>
<!-- #include File="CommonFunctions.asp" -->
<%
' ------------------------------------------------------------------------------
Class clsDataList
' ------------------------------------------------------------------------------

	Public AllowSorting					' Allow users to sort fields?
	Public TableName					' name of table to request data from
	Public PageSize						' Number of records on each page
	Public PrimaryKey					' Field identifying record
	Public sSql							' hidden String to execute
	Public PartyCode                    ' Party Code

	Private msReturnField
	Private msReturnFieldAry			' Caption to display for returned fields
	Private mnReturnFieldCount			' Number of fields to be returned


	Private msEnableField				' to enable or disable the Item Type List Box
	Private msOptDispField			    ' to display checkbox or radio button based parameter
	Private msDispAddNew                ' to display the Add New button based on the Parameter
	Private msDisplayField
	Private msDisplayFieldAry			' Caption to display for Display fields
	Private mnDisplayFieldCount			' Number of fields to be Displayed

	Private sSearchFor					' Search For Description
	Private msSearchField
	Private msSearchFieldAry			' Caption to display for Search Text fields
	Private msSearchValue
	Private msSearchValueAry			' Caption to display for Search Value fields
	Private mnSearchFieldCount			' Number of fields for Search
	
	Private msSearchTypeField
	Private msSearchTypeFieldAry
	Private msSearchTypeValue
	Private msSearchTypeValueAry
	Private mnSearchTypeFieldCount

	Private sSearchFor1					' Search For Description
	Private msSearchField1
	Private msSearchFieldAry1			' Caption to display for Search Text fields
	Private msSearchValue1
	Private msSearchValueAry1			' Caption to display for Search Value fields
	Private mnSearchFieldCount1			' Number of fields for Search

	Private msURLPrefix					' URL prefix is used to create other URL's
	Private msQuery						' Query user is searching for
	Private mnPage						' Page being viewed
	Private msSearch					' Search option user is searching for
	Private msItemType					' Selected ItemType
	Private mnLastPage					' Last page of results
	Private mnRecordCount				' Number of records found
	Private DisableBut					' AddToList Button
	Private msDispItem					' To Display the Item if > 0 then disabled else enabled
	
' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		mnReturnFieldCount = 0
		msSearch = Request.QueryString("SearchBy")
		msItemType = Request.QueryString("sIType")
		msQuery = Request.QueryString("Query")
		mnPage = Request.QueryString("Page")
		If mnPage = "" Or Not IsNumeric(mnPage) Then mnPage = 1 Else mnPage = CLng(mnPage)
		If mnPage < 1 Then mnPage = 1
		PageSize = 10
		AllowSorting = True
	End Sub
' ------------------------------------------------------------------------------
	Public Sub AddDisplayField(ByRef psCaption)

		' Appends another field to memory to be returned on the list
		msDisplayField = msDisplayField & ", " & psCaption

		mnDisplayFieldCount = mnDisplayFieldCount + 1

	End Sub ' AddDisplayField(ByRef psCaption, ByRef psFieldName)
' ------------------------------------------------------------------------------
	Public Sub AddEnableField(ByRef psCaption)
		' Appends another field to memory to be returned on the list
		msEnableField = psCaption
	End Sub ' AddDisplayField(ByRef psCaption, ByRef psFieldName)
' ------------------------------------------------------------------------------
	Public sub AddOptDispField(ByRef psCaption)
		msOptDispField   = psCaption
	End Sub
'----------------------------------------------------------------------------------
    Public sub AddDispAddButt(ByRef psCaption)
        msDispAddNew = psCaption
    End Sub
'------------------------------------------------------------------------
	Public sub AddDispItem(ByRef psCaption)
		msDispItem = psCaption
	End Sub 
' ------------------------------------------------------------------------------
	Public Sub AddReturnedField(ByRef psCaption)

		' Appends another field to memory to be returned on the list
		msReturnField = msReturnField & ", " & psCaption

		mnReturnFieldCount = mnReturnFieldCount + 1

	End Sub ' AddReturnedField(ByRef psCaption, ByRef psFieldName)
' ------------------------------------------------------------------------------
	Public Sub AddSearchField(ByRef psCaption,ByRef psValue)

		' Appends another field to memory to be returned on the list
		msSearchField = msSearchField & ", " & psCaption

		' Appends another field to memory to be returned on the list
		msSearchValue = msSearchValue & ", " & psValue

		mnSearchFieldCount = mnSearchFieldCount + 1

	End Sub ' AddSearchField(ByRef psCaption, ByRef psValue)
' ------------------------------------------------------------------------------
    Public Sub AddSearchTypeField(ByRef psCaption,ByRef psValue)

		' Appends another field to memory to be returned on the list
		msSearchTypeField = msSearchTypeField & ", " & psCaption

		' Appends another field to memory to be returned on the list
		msSearchTypeValue = msSearchTypeValue & ", " & psValue

		mnSearchTypeFieldCount = mnSearchTypeFieldCount + 1

	End Sub ' AddSearchField(ByRef psCaption, ByRef psValue)
' ------------------------------------------------------------------------------
	Public Sub AddSearchField1(ByRef psCaption,ByRef psValue)

		' Appends another field to memory to be returned on the list
		msSearchField1 = msSearchField1 & ", " & psCaption

		' Appends another field to memory to be returned on the list
		msSearchValue1 = msSearchValue1 & ", " & psValue

		mnSearchFieldCount1 = mnSearchFieldCount1 + 1

	End Sub ' AddSearchField(ByRef psCaption, ByRef psValue)
' ------------------------------------------------------------------------------
	Public Sub SearchForDesc(ByRef psCaption)

		sSearchFor = psCaption

	End Sub ' SearchForDesc(ByRef psCaption)
' ------------------------------------------------------------------------------
	Public Function GetTable(ByRef poDatabase)

		Dim lsSQL				' Structured Query Language
		Dim lvDataAry			' Data returned from database
		Dim lsTable				' Table of data to be returned
		' Validate that we have needed data and objects

		If Not msReturnField = "" Then
			msReturnFieldAry = Split(Mid(msReturnField, 3), ", ")
		End If

		If Not msDisplayField = "" Then
			msDisplayFieldAry = Split(Mid(msDisplayField, 3), ", ")
		End If

		If Not msSearchField = "" Then
			msSearchFieldAry = Split(Mid(msSearchField, 3), ", ")
			msSearchValueAry = Split(Mid(msSearchValue, 3), ", ")
		End If

		If Not msSearchField = "" Then
			msSearchFieldAry1 = Split(Mid(msSearchField1, 3), ", ")
			msSearchValueAry1 = Split(Mid(msSearchValue1, 3), ", ")
		End If
		
		IF Not msSearchTypeField = "" Then
		    msSearchTypeFieldAry = split(Mid(msSearchTypeField,3),", ")
		    msSearchTypeValueAry =  split(Mid(msSearchTypeValue,3),", ")
		End If
		'Response.Write "msItemType:"&mnDisplayFieldCount &"<BR>"
		' Build SQL to query database
		lsSQL = sSql
		'lsSearch = poDatabase.QueryFields(msSearchFieldAry, lsSQL)
		
		' Request page data
		poDatabase.AbsolutePage = mnPage
		poDatabase.PageSize = PageSize

		Call poDatabase.SetData(lsSQL, lvDataAry)

		mnLastPage = poDatabase.PageCount
		mnRecordCount = poDatabase.RecordsAffected


		' Build the table
		lsTable = "<Body OnLoad="" Init() "">"

		lsTable = lsTable & _
			"<table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" class=""PopupTable""><tr><td class=""MiddlePack""></td></tr>" & _
				"<tr>" & _
					"<td align=""center"" class=""TopPack"">" & _

			"<TABLE class=""ExcelTable"" cellspacing=""1"" width=""100%"">" & _

			GetHeader() & _
			"<FORM name=""FormName"" >"

		If Not PrimaryKey = "" Then
			lsTable = lsTable & "<INPUT type=""hidden"" name=""pKeyName"" value=""" & PrimaryKey & """>"
		End If
		lsTable = lsTable & _
			GetData(lvDataAry) & _
			GetFooter() & _
			"</FORM>" & _
			"</TD></TR></TABLE>"& _
			"</TABLE>"

		lsTable = lsTable & "</Body>"
		Response.Write lsTable

	End Function ' GetTable(ByRef poDatabase)
' ------------------------------------------------------------------------------
	Private Function GetHeader()
		Dim lsHTML
		Dim lnIndex
		Dim lnCount
		Dim lsFieldname
		Dim lsCaption
		Dim lnColSpan
		

		lnColSpan = mnDisplayFieldCount
		If Not PrimaryKey = "" Then lnColSpan = lnColSpan + 1
		If ShowTools() Then lnColSpan = lnColSpan + 1

		lsHTML = lsHTML & "<TR class=""ExcelHeaderCell"" align=""center"">"
		lsHTML = lsHTML & "<TD colSpan=""" & lnColSpan & """>"
		lsHTML = lsHTML & "Found " & mnRecordCount & " Records."
		lsHTML = lsHTML & " Displaying Page " & mnPage & " of " & mnLastPage & "."
		lsHTML = lsHTML & "</TD>"
		lsHTML = lsHTML & "</TR>"

		lsHTML = lsHTML & "<TR class=""ExcelHeaderCell"">"

		If Not PrimaryKey = "" Then
			lsHTML = lsHTML & "<TD align=""center"" width=""80"">Select</TD>"
		End If

		If AllowSorting Then
			lnCount = mnDisplayFieldCount - 1
			For lnIndex = 0 To lnCount
				lsCaption = msDisplayFieldAry(lnIndex)

				lsHTML = lsHTML & "<TD align=""center"">" & lsCaption & ""
				lsHTML = lsHTML & "</TD>"
			Next
		Else
			lsHTML = "<TD>" & Join(msDisplayFieldAry, "</TD><TD>") & "</TD>"
		End If
		lsHTML = lsHTML & "</TR>"
		GetHeader = lsHTML
	End Function
' ------------------------------------------------------------------------------
	Private Function GetData(ByRef pvDataAry)

		Dim lnRowIndex,lnRowCount,lnColCount,lnColIndex,lsHTML,lsClass,lsItemStock,lsPartyCode,lsPartyType,lsPartySubType,lsSuppItemCode,lsSuppItemDesc
		Dim lvPrimaryValue,lnMax,lsData,iCounter,iChkCnt,iEntNo,lsUOM
		Dim i,iOptValTemp,sOptNameTemp,iCtr,lsItmCode,sSql,sCheckNoOfBinAndLoc,sLocNo,sBinNo
		Dim dcrs,objrs
		set dcrs = Server.CreateObject("ADODB.Recordset")
		set objrs = Server.CreateObject("ADODB.Recordset")

		dim con,connFile,connString,connPath,connArray
		redim connArray(3)
		const connReading = 1
		set con = Server.CreateObject("ADODB.Connection")
		set connFile = Server.CreateObject("Scripting.FileSystemObject")
		connPath = Server.MapPath("/include/Settings.inf")
		set connFile = connFile.OpenTextFile(connPath,connReading,true)
		connFile.ReadLine()
		connString = connFile.ReadLine()
		connArray = split(connString,":")
		con.Open connArray(0),connArray(1),connArray(2)

		iChkCnt = 0
		iEntNo = 0
		'Response.Write MsitemType
		IF trim(Right(msItemType,3)) = "STO" or trim(Right(msItemType,3)) = ""  or trim(msOptDispField) = "S" then
			DisableBut =  "N"
		Else
			DisableBut =  "Y"
		End IF
		If IsArray(pvDataAry) Then
			lnRowCount = UBound(pvDataAry, 2)
		Else
			lnRowCount = -1
		End If

		lnMax = PageSize - 1
		lnColCount = mnDisplayFieldCount - 1
	 	For lnRowIndex = 0 To lnRowCount
	 		iChkCnt = iChkCnt + 1
			If lnRowIndex Mod 2 = 0 Then lsClass = "ExcelDisplayCell" Else lsClass = "ExcelDisplayCell"
			lsHTML = lsHTML & "<TR class=""" & lsClass & """>"
			lsUOM = pvDataAry(6,lnRowIndex)
			lsItemStock = pvDataAry(4,lnRowIndex)
			lsPartyCode = pvDataAry(12,lnRowIndex)
			lsPartyType = pvDataAry(13,lnRowIndex)
			lsPartySubType = pvDataAry(14,lnRowIndex)
			lsSuppItemCode = pvDataAry(2,lnRowIndex)
			lsSuppItemDesc = pvDataAry(3,lnRowIndex)
			
			If Not PrimaryKey = "" Then

				If lnRowIndex <= lnRowCount Then
					lvPrimaryValue = ""
					iCounter = 0
					for iCounter = 0 to mnReturnFieldCount - 1
						if iCounter = 4 then 'item name
							lvPrimaryValue = lvPrimaryValue & replace(pvDataAry(msReturnFieldAry(iCounter), lnRowIndex),chr(34),"~~") & ":"
						else
							lvPrimaryValue = lvPrimaryValue & pvDataAry(msReturnFieldAry(iCounter), lnRowIndex) & ":"
						end if
					next
					lvPrimaryValue = mid(lvPrimaryValue,1,len(lvPrimaryValue)-1)
					lsHTML = lsHTML & "<TD align=""center"">"
					'Response.Write "Dis="&DisableBut
					if msOptDispField = "M" then
						'Response.Write "DisableBut="&DisableBut

						if trim(DisableBut) = "N" then
							if lsItemStock = "0" and msDispItem = "1" then
								lsHTML = lsHTML & "<INPUT type=""checkbox"" class=""FormElem"" name=""pKey"&lnRowIndex&""" value=""" & lvPrimaryValue &""" OnClick=""XmlFun(this)"" disabled>"
							else
								lsHTML = lsHTML & "<INPUT type=""checkbox"" class=""FormElem"" name=""pKey"&lnRowIndex&""" value=""" & lvPrimaryValue &""" OnClick=""XmlFun(this)"">"
							end if
						elseif trim(DisableBut) = "Y" then
							if lsItemStock = "0" and msDispItem = "1" then
								lsHTML = lsHTML & "<INPUT type=""checkbox"" class=""FormElem"" name=""pKey"&lnRowIndex&""" value=""" & lvPrimaryValue &""" disabled> "
							else
								lsHTML = lsHTML & "<INPUT type=""checkbox"" class=""FormElem"" name=""pKey"&lnRowIndex&""" value=""" & lvPrimaryValue &"""> "
							end if
						end if
					else
						  ' Response.Write DisableBut &"<br><BR>"
						if trim(DisableBut) = "N" then
							if lsItemStock = "0" and msDispItem = "1" then
								lsHTML = lsHTML & "<INPUT type=""Radio"" class=""FormElem"" name=""pKey"&iCounter&""" Value=""" & lvPrimaryValue & """ OnClick=""XmlFun(this)"" disabled> "
							else
								lsHTML = lsHTML & "<INPUT type=""Radio"" class=""FormElem"" name=""pKey"&iCounter&""" Value=""" & lvPrimaryValue & """ OnClick=""XmlFun(this)""> "
							end if
						elseif trim(DisableBut) = "Y" then
							if lsItemStock = "0" and msDispItem = "1" then
								lsHTML = lsHTML & "<INPUT type=""Radio"" class=""FormElem"" name=""pKey"&iCounter&""" Value=""" & lvPrimaryValue & """ OnClick=""XmlFun(this)"" disabled>  "
							else
								lsHTML = lsHTML & "<INPUT type=""Radio"" class=""FormElem"" name=""pKey"&iCounter&""" Value=""" & lvPrimaryValue & """ OnClick=""XmlFun(this)"">  "
							end if
						end if
					end if

					lsHTML = lsHTML & "</TD>"


				Else
					lsHTML = lsHTML & "<TD>&nbsp;</TD>"

				End If

			End If

			'Response.Write "iChkCnt="&iChkCnt

			Dim lsArr,sAttName,sAttArr,sAttTemp,iOptVal,sOptName,iCnt,sAttributeList,nGetItemRate
			Dim sFinYrFrom,sFinYrTo,nGetMarketPrice
			sFinYrFrom = "01/04/"& Split(Session("FinPeriod"),":")(0)
			sFinYrTo = "31/03/" & Split(Session("FinPeriod"),":")(1)
				
			IF lnColIndex <> 0 then
				iCnt = lnColIndex - 1

			else

				'IF mnDisplayFieldCount = 6 then	iCnt = 5
				'IF mnDisplayFieldCount = 4 then	iCnt = 3
				iCnt = cint(mnDisplayFieldCount) - 1
			End IF

			For lnColIndex = 0 To lnColCount
				lsHTML = lsHTML & "<TD>"
				IF cint(lnColIndex) <>  cint(iCnt) then
					If lnRowIndex <= lnRowCount Then
						lsData = pvDataAry(lnColIndex, lnRowIndex)
						response.write "<font color=red>"
						If VarType(lsData) = vbNull Then
							lsHTML = lsHTML & "&nbsp;"
						Else
							If Len(lsData) > 255 Then
								lsData = lsData & Left(lsData, 255) & " ..."
							End If
							lsData = Server.HTMLEncode(lsData)
						'	Response.Write "<p>lsData="&lnColIndex & "--"& lsData 
							lsHTML = lsHTML & lsData
						End If
					Else
						lsHTML = lsHTML & "&nbsp;"
					End If
				ElseIF cint(lnColIndex) = cint(iCnt) then ' or  lnColIndex = 3

					lsData = pvDataAry(11, lnRowIndex)
					lsItmCode = pvDataAry(7, lnRowIndex)
					lsClass = pvDataAry(8, lnRowIndex)
					
					response.write "<font color=red>"
					'Response.Write "<p>lsData="&lsData
					'	Response.Write "<p>"&"lsItmCode = "&lsItmCode
				'		Response.Write "<p>"&"lsClassCode = "&lsClass
					
					IF  lsData = 0 or lsData = "" then
						lsHTML = lsHTML & " NA "
						'DisableBut =  "N"
					Else
						DisableBut =  "Y"

					End IF
					msDisplayFieldAry = split(lsData,",")
					'Response.Write lsData
					'Response.Write lsItmCode &"->"
					'response.write "<font color=red>stock = "& lsItemStock & "=="&lsUOM
					
					iCtr = 1

					with objrs
						.CursorLocation = 3
						.CursorType = 3
						.ActiveConnection = con
						.Source = "Select isNull(AttributeList,0) from INV_M_ItemMaster where ItemCode= "& lsItmCode &""
						.Open
					end with
					if not objrs.EOF then
						do while not objrs.EOF
							sAttributeList =  objrs(0)
							objrs.MoveNext
						loop
					end if
					objrs.Close
					'sAttributeList = sAttributeList
					with dcrs
						.CursorLocation = 3
						.CursorType = 3
						.ActiveConnection = con
						'.Source = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,ItemTypeID from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where O.ItemTypeAttributeID = A.ItemTypeAttributeID and ItemTypeID='"& trim(right(MsitemType,3)) &"' and A.ItemTypeAttributeID in ("& sAttributeList &") Group by A.ItemTypeAttributeID,A.ItemTypeAttributeName,ItemTypeID"
						.Source = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where O.ItemTypeAttributeID = A.ItemTypeAttributeID and A.ItemTypeAttributeID in ("& sAttributeList &") Group by A.ItemTypeAttributeID,A.ItemTypeAttributeName"
						' Response.Write dcrs.Source
						.Open
					end with
					if not dcrs.EOF then
						lsArr = 0
						do while not dcrs.EOF

							lsHTML = lsHTML & "<SELECT Name=""SelAttribList"&lsItmCode&lsClass&""" size=""1""  class=""FormElem"" "">"

							lsHTML = lsHTML & "<OPTION value="""& dcrs(0) &"#0:"&dcrs(1)&""" selected>"&dcrs(1)&"</OPTION>"
								lsHTML = lsHTML & "<OPTION value="""& dcrs(0) &"#0:"&dcrs(1)&""">"&dcrs(1)&"</OPTION>"
								'Response.Write lsHTML
								with objrs
									.CursorLocation = 3
									.CursorType = 3
									.ActiveConnection = con
									.Source = "Select A.ItemTypeAttributeID,A.ItemTypeAttributeName,O.OptionValue,O.OptionName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O where O.ItemTypeAttributeID = A.ItemTypeAttributeID and A.ItemTypeAttributeID = "& dcrs(0)
									'Response.Write objrs.Source
									.Open
								end with
								if not objrs.EOF then
									do while not objrs.EOF
										lsHTML = lsHTML & "<OPTION value="""& dcrs(0) &"#"&objrs(2)&":"&objrs(3)&""" >"&objrs(3)&"</OPTION>"
										objrs.MoveNext
									loop
								end if
								objrs.Close
								lsHTML = lsHTML & "</SELECT>"
								lsArr = lsArr  + 1
									lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hAttribVal"&lsItmCode&lsClass&""" value=""Y"">"
							dcrs.MoveNext
						loop
					else
						lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hAttribVal"&lsItmCode&lsClass&""" value=""N"">"
					end if
					dcrs.Close

				'	For lsArr = 0 TO UBOUND(msDisplayFieldAry)
				'
				'		IF msDisplayFieldAry(lsArr) <> 0 then
				'			sAttArr  = FunAttribVal(msDisplayFieldAry(lsArr))
				'			sAttTemp = split(sAttArr,":")
				'			sAttName = sAttTemp(0)
				'			iOptValTemp = split(sAttTemp(1),"-")
				'			sOptNameTemp = split(sAttTemp(2),"-")
				'			Response.Write msDisplayFieldAry(lsArr)
				'		'Response.Write sAttArr
				'			lsHTML = lsHTML & "<SELECT Name=""SelAttribList"&lsItmCode&lsArr&""" size=""1""  class=""FormElem"" "">" 'onclick=""FunSelAttrib("&lsItmCode&","&lsArr&")"">"
				'			If msDisplayFieldAry(lsArr) = 6 then lsHTML = lsHTML & "<OPTION value="""&msDisplayFieldAry(lsArr)&":DIA"" selected>DIA</OPTION>"
				'			If msDisplayFieldAry(lsArr) = 2 then lsHTML = lsHTML & "<OPTION value="""&msDisplayFieldAry(lsArr)&":SIZE"" selected>SIZE</OPTION>"
				'			If msDisplayFieldAry(lsArr) = 3 then lsHTML = lsHTML & "<OPTION value="""&msDisplayFieldAry(lsArr)&":COLOR"" selected>COLOR</OPTION>"
				'			If msDisplayFieldAry(lsArr) = 4 then lsHTML = lsHTML & "<OPTION value="""&msDisplayFieldAry(lsArr)&":TYPE"" selected>TYPE</OPTION>"
				'
				'			For i = 0 to UBOUND(iOptValTemp)
				'
				'				iOptVal = iOptValTemp(i)
				'				sOptName = sOptNameTemp(i)
				'			'	Response.Write "iOptVal ="& iOptVal
				'			'	Response.Write " sOptName ="& sOptName
				'				IF trim(iOptVal) = "" then lsHTML = lsHTML & "<OPTION value="""&msDisplayFieldAry(lsArr)&":"&sAttName&""" selected>"&sAttName&"</OPTION>"
				'				IF iOptVal <> "" then
				'					lsHTML = lsHTML & "<OPTION value="""&iOptVal&":"&sOptName&""" >"&sOptName&"</OPTION>"
				'				End IF
				'			Next
				'			lsHTML = lsHTML & "</SELECT>"
				'
				'		End IF
				'	Next
			'****	 Blocked by ragav
				'	lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hAttribVal"&lsItmCode&""" value=""" &lsArr & """>"
			'***
				sCheckNoOfBinAndLoc = "0"
				sSql = "Select count(*),LocationNumber,isNull(BinNumber,0) From Inv_T_ItemLocationStock where ItemCode="& lsItmCode &" and ClassificationCode="& lsClass&" and convert(DateTime,FinancialYearFrom,103) >= Convert(DateTime,'"& sFinYrFrom &"',103) and convert(DateTime,FinancialYearTo,103) <= Convert(DateTime,'"& sFinYrTo&"',103) Group by LocationNumber,BinNumber"
				objrs.Open sSql,con
				If Not objrs.EOF Then
					Do while Not objrs.EOF 
						sCheckNoOfBinAndLoc = sCheckNoOfBinAndLoc + 1
						sLocNo = objrs(1)
						sBinNo = objrs(2)
					objrs.MoveNext 
					Loop
				End IF
				objrs.Close 
				
			    'nGetItemRate = GetItemRate(Session("organizationcode"),Session("FinPeriod"),lsClass,lsItmCode,"")
				nGetItemRate = GetItemSalePrice(Session("organizationcode"),Date(),lsClass,lsItmCode,PartyCode)
				nGetMarketPrice = GetMarketPrice(Session("organizationcode"),lsClass,lsItmCode)
			    lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hItemRate"&lsItmCode&lsClass&""" value="""& nGetItemRate &""">"
			    lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hItemStock"&lsItmCode&lsClass&""" value="""& lsItemStock &""">"
			    lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hBinAndLocCheck"&lsItmCode&lsClass&""" value="""& sCheckNoOfBinAndLoc &""">"
			    lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hLocNo"&lsItmCode&lsClass&""" value="""& sLocNo &""">"
			    lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hBinNo"&lsItmCode&lsClass&""" value="""& sBinNo &""">"
			    lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hMarketPrice"&lsItmCode&lsClass&""" value="""& nGetMarketPrice &""">"
			    End IF
				lsHTML = lsHTML & "</TD>"
			Next

			lsHTML = lsHTML & "</TR>"
		Next
		lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hDisableBut"" value=""" &DisableBut& """>"
		lsHTML = lsHTML & "<INPUT type=""hidden"" name=""hChkCount"" value="""&iChkCnt&""">"
		GetData = lsHTML
	End Function
' ------------------------------------------------------------------------------
	Private Sub AppendLink(ByRef psContent, ByRef psLink, ByRef psDelimiter)
		If psContent = "" Then
			psContent = psLink
		Else
			psContent = psContent & psDelimiter & psLink
		End If
	End Sub
' ------------------------------------------------------------------------------
	Private Function ShowTools()
		ShowTools = True
	End Function
' ------------------------------------------------------------------------------
	Private Function Tools()
		Dim lsLinks
		Tools = lsLinks
	End Function
' ------------------------------------------------------------------------------
	Private Function PageQuery(ByRef pnPage)
		Dim lsQuery,iCount

		lsQuery = Request.QueryString

		iCount = InStr(1,lsQuery,"&Query=")
		If cint(iCount) > 0 Then
			lsQuery = left(lsQuery,iCount - 1)
		End If

		lsQuery = Replace(lsQuery, "Page=" & mnPage, "")

		If pnPage < 1 Then
			pnPage = 1
		ElseIf pnPage > mnLastPage Then
			pnPage = mnLastPage
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
' ------------------------------------------------------------------------------
	Private Function GetFooter()
		Dim lsHTML
		Dim lnColSpan
		Dim lsLinks
		Dim lsPageJump
		Dim lnPage,lnSearch
		Dim lsQueryString
		Dim sChkBut
		lnColSpan = mnReturnFieldCount
		If Not PrimaryKey = "" Then lnColSpan = lnColSpan + 1
		If ShowTools() Then lnColSpan = lnColSpan + 1

		sChkBut = DisableBut
		If mnLastPage >= 2 Then
			' Create the links.

			if mnPage = "1" then
				' First
				AppendLink lsLinks, "<span>First</span>", " | "
				' Previous
				AppendLink lsLinks, "<span>Previous</span>", " | "
			else
				' First
				AppendLink lsLinks, "<span style=""cursor: hand"" onClick=""showpage('" & PageQuery(1) & "')"">First</span>", " | "
				' Previous
				AppendLink lsLinks, "<span style=""cursor: hand"" onClick=""showpage('" & PageQuery(mnPage - 1) & "')"">Previous</span>", " | "
			end if

			If mnLastPage >= 2 Then
				lsQueryString = Request.QueryString
				lsQueryString = Replace(lsQueryString, "Page=" & mnPage, "")
				If lsQueryString = "" Then
					lsQueryString = lsQueryString & "Page="
				ElseIf Right(lsQueryString, 1) = "&" Then
					lsQueryString = lsQueryString & "Page="
				Else
					lsQueryString = lsQueryString & "&Page="
				End If
				lsQueryString = lsQueryString
				lsPageJump = "<SELECT class=""FormElem"" onChange=""showpage('" & lsQueryString & "' + this[this.selectedIndex].value)"">"
				For lnPage = 1 To mnLastPage
					If lnPage = mnPage Then
						lsPageJump = lsPageJump & "<OPTION value=""" & lnPage & """ selected>Page " & lnPage & " of " & mnLastPage & "</OPTION>"
					Else
						lsPageJump = lsPageJump & "<OPTION value=""" & lnPage & """>Page " & lnPage & "</OPTION>"
					End If
				Next
				lsPageJump = lsPageJump & "</SELECT>"
				AppendLink lsLinks, lsPageJump, " | "
			End If
			if mnPage = mnLastPage then
				' Next
				AppendLink lsLinks, "<span>Next</span>", " | "
				' Last
				AppendLink lsLinks, "<span>Last</span>", " | "
			else
				' Next
				AppendLink lsLinks, "<span style=""cursor:hand"" onClick=""showpage('" & PageQuery(mnPage + 1) & "')"">Next</span>", " | "
				' Last
				AppendLink lsLinks, "<span style=""cursor:hand"" onClick=""showpage('" & PageQuery(mnLastPage) & "')"">Last</span>", " | "
			end if

			lsHTML = lsHTML & "<TR class=""ExcelFooterCell"" align=""center"">"
			lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
			lsHTML = lsHTML & lsLinks
			lsHTML = lsHTML & "</TD>"
			lsHTML = lsHTML & "</TR>"
			'Response.Write "sChkBut="&sChkBut
			lsHTML = lsHTML & "<TR class=""ExcelFooterCell"" align=""center"">"
			lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"

			IF trim(sChkBut) = "Y" then
				lsHTML = lsHTML & "<INPUT TYPE=BUTTON VALUE=""Add To List"" ONCLICK=""AddFun()"" class=""ActionButtonx"" id=BUTTON1 name=BUTTON1 >"
			Else
				lsHTML = lsHTML & "<INPUT TYPE=BUTTON VALUE=""Add To List"" ONCLICK=""AddFun()"" class=""ActionButtonx"" id=BUTTON1 name=BUTTON1 Disabled>"
			End IF

			lsHTML = lsHTML & "&nbsp;&nbsp;<INPUT TYPE=BUTTON VALUE=""Done"" ONCLICK=""sendValue()"" class=""ActionButton"">"
			lsHTML = lsHTML & "</TD>"
			lsHTML = lsHTML & "</TR>"


		elseIf mnLastPage = 1 Then
			lsHTML = lsHTML & "<TR class=""ExcelFooterCell"" align=""center"">"
			lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
			lsHTML = lsHTML & lsLinks
			lsHTML = lsHTML & "</TD>"
			lsHTML = lsHTML & "</TR>"
			'Response.Write "sChkBut="&sChkBut
			lsHTML = lsHTML & "<TR class=""ExcelFooterCell"" align=""center"">"
			lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
			IF trim(sChkBut) = "Y" then
				lsHTML = lsHTML & "<INPUT TYPE=BUTTON VALUE=""Add To List"" ONCLICK=""AddFun()"" class=""ActionButtonx"" id=BUTTON1 name=BUTTON1>"
			Else
				lsHTML = lsHTML & "<INPUT TYPE=BUTTON VALUE=""Add To List"" ONCLICK=""AddFun()"" class=""ActionButtonx"" id=BUTTON1 name=BUTTON1 Disabled>"
			End IF
			lsHTML = lsHTML & "&nbsp;&nbsp;<INPUT TYPE=BUTTON VALUE=""Done"" ONCLICK=""sendValue()"" class=""ActionButton"" id=BUTTON2 name=BUTTON2>"
			lsHTML = lsHTML & "</TD>"
			lsHTML = lsHTML & "</TR>"

		End If
		
		

		' Do Search
		lsLinks = ""
		lsHTML = lsHTML & "<TR class=""ExcelFooterCell"">"
		lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
		lsHTML = lsHTML & "Search By <select size=""1"" name=""SearchType"" class=""FormElem"" onChange=""document.FormName.Query.value=''"">"
		For lnSearch = 0 To mnSearchTypeFieldCount - 1
			if msSearchTypeValueAry(lnSearch) = Server.HTMLEncode(msSearch) then
				lsHTML = lsHTML & "<OPTION value=""" & msSearchTypeValueAry(lnSearch) & """ selected>" & msSearchTypeFieldAry(lnSearch) & "</OPTION>"
			else
				lsHTML = lsHTML & "<OPTION value=""" & msSearchTypeValueAry(lnSearch) & """>" & msSearchTypeFieldAry(lnSearch) & "</OPTION>"
			end if
		Next
		lnSearch = 0
		lsHTML = lsHTML & "</Select>"
		
		lsHTML = lsHTML & "  <select size=""1"" name=""SearchBy"" class=""FormElem"" onChange=""document.FormName.Query.value=''"">"
		For lnSearch = 0 To mnSearchFieldCount - 1
			if msSearchValueAry(lnSearch) = Server.HTMLEncode(msSearch) then
				lsHTML = lsHTML & "<OPTION value=""" & msSearchValueAry(lnSearch) & """ selected>" & msSearchFieldAry(lnSearch) & "</OPTION>"
			else
				lsHTML = lsHTML & "<OPTION value=""" & msSearchValueAry(lnSearch) & """>" & msSearchFieldAry(lnSearch) & "</OPTION>"
			end if
		Next
		lnSearch = 0
		lsHTML = lsHTML & "</Select>"
				
		'Response.write "msSearchFieldAry1(lnSearch)="&msSearchFieldAry1(lnSearch)
		if msEnableField = "2" then

		'	lsHTML = lsHTML & " Item Type <select size=""1"" name=""selIType"" class=""FormElem"" onChange=""DeleteNodes();showpage('" & PageQuery(0) & "')"" DISABLED>"
		else
		'	lsHTML = lsHTML & " Item Type <select size=""1"" name=""selIType"" class=""FormElem"" onChange=""DeleteNodes();showpage('" & PageQuery(0) & "')"">"
		end if
		For lnSearch = 0 To mnSearchFieldCount1 - 1
			if msSearchValueAry1(lnSearch) = Server.HTMLEncode(msItemType) then
				lsHTML = lsHTML & "<OPTION value=""" & msSearchValueAry1(lnSearch) & """ selected>" & msSearchFieldAry1(lnSearch) & "</OPTION>"
			else
				lsHTML = lsHTML & "<OPTION value=""" & msSearchValueAry1(lnSearch) & """>" & msSearchFieldAry1(lnSearch) & "</OPTION>"
			end if
		Next

		lsHTML = lsHTML & "</Select>"
		lsHTML = lsHTML & "</TD>"

		lsHTML = lsHTML & "</TR>"

		lsHTML = lsHTML & "<TR class=""ExcelFooterCell"">"
		lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
		lsHTML = lsHTML & "" & sSearchFor & " <INPUT type=""text"" class=""FormElem"" name=""Query"" value=""" & Server.HTMLEncode(msQuery) & """ size=""20"">"
		lsHTML = lsHTML & "&nbsp;&nbsp;&nbsp; <INPUT type=""button"" value=""Search"" class=""ActionButton"" onClick=""showpage('" & PageQuery(0) & "')"" id=1 name=1>"
		if trim(msDispAddNew)="Y" then
		lsHTML = lsHTML & "&nbsp;&nbsp;&nbsp; <INPUT type=""button"" value=""Add New"" class=""ActionButton"" onClick=""WithOutMat()"" id=1 name=1>"
		end if '  if trim(msDispAddNew)="Y" then
		lsHTML = lsHTML & "<INPUT type=""Hidden"" name=""hSelectMode"" value=""" & msOptDispField & """ >"

		lsHTML = lsHTML & "</TD>"
		lsHTML = lsHTML & "</TR>"


		'if msOptDispField = "M" then
			lsHTML = lsHTML & "<TR class=""ExcelFooterCell"">"
			lsHTML = lsHTML & "<TD colspan=""" & lncolSpan & """>"
			lsHTML = lsHTML & "Selected Entries &nbsp;&nbsp;  <Span id=idSelList></span>"
			lsHTML = lsHTML & "</TD>"
			lsHTML = lsHTML & "</TR>"

		'end if 'if msOptDispField = "M" then

		lsHTML = lsHTML & "</FORM>"

		GetFooter = lsHTML
	End Function

End Class
' ------------------------------------------------------------------------------
Function FunAttribVal(sVal)
Dim rs,sSql,sAttribName,iOptVal,sOptName
set rs=server.CreateObject("ADODB.Recordset")

	sSql = "Select A.ItemTypeAttributeName,O.OptionValue,O.OptionName from INV_M_ItemTypeAttributes as A,Inv_M_ItemTypeOptions as O "&_
		   "where A.ItemTypeAttributeID = "&sVal&" And O.ItemTypeAttributeID = A.ItemTypeAttributeID "
	rs.Open sSql,con
	'Response.Write ssql
	'Response.Write rs.EOF
	Do while not rs.EOF
		sAttribName = rs(0)
		iOptVal		= iOptVal &"-"&rs(1)
		sOptName	= sOptName &"-"&rs(2)

		FunAttribVal = sAttribName	&":"& iOptVal &":"&  sOptName
	rs.MoveNext
	loop
	rs.Close

End Function
' ------------------------------------------------------------------------------
%>
<SCRIPT id="Datalist.Scripts">
	<!-- // hide from old browsers (this still needed today?)
	function sendValue1()
	{
		var loForm = new Object(document.FormName)
		if(loForm.pKeyName)
		{
			lsName = loForm.pKeyName.value;
			for(var i=0;i<loForm.pKey.length;i++)
			{
				if(loForm.pKey[i].checked)
				{
					sRet = loForm.pKey[i].value
					window.close();
				}
			}
		}
	}
	// -->
</SCRIPT>
<!--#include file="SupplierItemSelectorCompat.asp"-->

