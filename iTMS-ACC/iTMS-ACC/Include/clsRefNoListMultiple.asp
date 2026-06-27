<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
' ------------------------------------------------------------------------------
Class clsDataList
' ------------------------------------------------------------------------------

	Public AllowSorting					' Allow users to sort fields?
	Public TableName					' name of table to request data from
	Public PageSize						' Number of records on each page
	Public PrimaryKey					' Field identifying record
	Public sSql							' hidden String to execute

	Private msReturnField
	Private msReturnFieldAry			' Caption to display for returned fields
	Private mnReturnFieldCount			' Number of fields to be returned


	Private msEnableField				' to enable or disable the Item Type List Box
	Private msOptDispField			    ' to display checkbox or radio button based parameter
	Private msDisplayField
	Private msDisplayFieldAry			' Caption to display for Display fields
	Private mnDisplayFieldCount			' Number of fields to be Displayed

	Private sSearchFor					' Search For Description
	Private msSearchField
	Private msSearchFieldAry			' Caption to display for Search Text fields
	Private msSearchValue
	Private msSearchValueAry			' Caption to display for Search Value fields
	Private mnSearchFieldCount			' Number of fields for Search

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
	Private msLinkField                 'View Path

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
    Public Sub LinkField(ByRef psCaption)
        msLinkField = psCaption
    End Sub
'----------------------------------------------------------------------------
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

		' Build SQL to query database
		lsSQL = sSql
		'lsSearch = poDatabase.QueryFields(msSearchFieldAry, lsSQL)

		' Request page data
		poDatabase.AbsolutePage = mnPage
		poDatabase.PageSize = PageSize
		'Response.Write "lvDataAry="&lsSQL
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
			lsHTML = lsHTML & "<TD class=""ExcelHeaderCell"" align=""center"" width=""80"">Select</TD>"
		End If

		If AllowSorting Then
			lnCount = mnDisplayFieldCount - 1
			For lnIndex = 0 To lnCount
				lsCaption = msDisplayFieldAry(lnIndex)
				'Response.Write "lsCaption="&lsCaption & ":"& lnIndex &"<Br><br>"
				'IF lnIndex = 0 then
				'	lsHTML = lsHTML & "<TD class=""ExcelHeaderCell"" align=""center"" width=""100"">" & lsCaption & ""
				'	lsHTML = lsHTML & "</TD>"
				'ElseIf lnIndex = 1 then
				'	lsHTML = lsHTML & "<TD align=""center"" width=""100"">" & lsCaption & ""
				'	lsHTML = lsHTML & "</TD>"
				'ElseIF lnIndex = 2 then
					lsHTML = lsHTML & "<TD class=""ExcelHeaderCell""  align=""center"">" & lsCaption & ""
					lsHTML = lsHTML & "</TD>"
				'End IF
			Next
		Else
			lsHTML = "<TD>" & Join(msDisplayFieldAry, "</TD><TD>") & "</TD>"
		End If
		lsHTML = lsHTML & "</TR>"
		GetHeader = lsHTML
	End Function
' ------------------------------------------------------------------------------
	Private Function GetData(ByRef pvDataAry)

		Dim lnRowIndex,lnRowCount,lnColCount,lnColIndex,lsHTML,lsClass
		Dim lvPrimaryValue,lnMax,lsData,iCounter
		
		If IsArray(pvDataAry) Then
		
			lnRowCount = UBound(pvDataAry, 2)
		Else
			lnRowCount = -1
		End If

		lnMax = PageSize - 1
		lnColCount = mnDisplayFieldCount - 1

		For lnRowIndex = 0 To lnRowCount
			If lnRowIndex Mod 2 = 0 Then lsClass = "ExcelDisplayCell" Else lsClass = "ExcelDisplayCell"
			lsHTML = lsHTML & "<TR class=""" & lsClass & """>"

			If Not PrimaryKey = "" Then

				If lnRowIndex <= lnRowCount Then
					lvPrimaryValue = ""
					iCounter = 0
					for iCounter = 0 to mnReturnFieldCount  -1					
						lvPrimaryValue = lvPrimaryValue & pvDataAry(msReturnFieldAry(iCounter),lnRowIndex) & ":"						
					next
					lvPrimaryValue = mid(lvPrimaryValue,1,len(lvPrimaryValue)-1)					
					
					lsHTML = lsHTML & "<TD align=""center"">"

							if msOptDispField = "M" then
								lsHTML = lsHTML & "<INPUT " &_
									" type=""checkbox""" &_
									" class=""FormElem""" &_
									" name=""pKey""" &_
									" value=""" & lvPrimaryValue &"""" &_
									" OnClick=XmlFun(this)"
							else 
								lsHTML = lsHTML & "<INPUT" &_
									" type=""Radio""" &_
									" class=""FormElem""" &_
									" name=""pKey""" &_
									" Value=""" & lvPrimaryValue & """" & _
									" OnClick=XmlFun(this)"
								'If lnRowIndex = 0 Then lsHTML = lsHTML & " checked"
					end if
					lsHTML = lsHTML & "></TD>"
				Else
					lsHTML = lsHTML & "<TD>&nbsp;</TD>"
				End If
			End If
			
			For lnColIndex = 0 To lnColCount			
				IF cint(lnColIndex) = 0 then
				    if trim(msLinkField)<>"" then
				        lsHtml = lsHTML & "<TD width=""130""><a href=""#"" class=""ExcelDisplayLink"" onClick=""ViewDetails("& pvDataAry(msLinkField,lnRowIndex)&","& lvPrimaryValue &")"">"
				    else
					    lsHTML = lsHTML & "<TD width=""130"">"
					end if 
				ElseIF cint(lnColIndex) = 1  then
					lsHTML = lsHTML & "<TD width=""100"" align=""center"">"
				Else
					lsHTML = lsHTML & "<TD>"
				End IF				
				'lsHTML = lsHTML & "<TD>"
				
				If lnRowIndex <= lnRowCount Then
					
					lsData = pvDataAry(lnColIndex, lnRowIndex)
					If VarType(lsData) = vbNull Then
						lsHTML = lsHTML & "&nbsp;"
					Else
						If Len(lsData) > 255 Then
							lsData = lsData & Left(lsData, 255) & " ..."
						End If
						lsData = Server.HTMLEncode(lsData)
						'Response.Write "lsData="&lsData
						lsHTML = lsHTML & lsData
					End If
				Else
					lsHTML = lsHTML & "&nbsp;"
				End If
				lsHTML = lsHTML & "</TD>"
			Next

			lsHTML = lsHTML & "</TR>"
		Next

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

	End Function
' ------------------------------------------------------------------------------
	Private Function GetFooter()
		Dim lsHTML
		Dim lnColSpan
		Dim lsLinks
		Dim lsPageJump
		Dim lnPage,lnSearch
		Dim lsQueryString

		lnColSpan = mnReturnFieldCount
		If Not PrimaryKey = "" Then lnColSpan = lnColSpan + 1
		If ShowTools() Then lnColSpan = lnColSpan + 1


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

			lsHTML = lsHTML & "<TR class=""ExcelFooterCell"" align=""center"">"
			lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
			lsHTML = lsHTML & "<INPUT TYPE=BUTTON VALUE=""Done"" ONCLICK=""sendValue()"" class=""ActionButton"">"
			lsHTML = lsHTML & "</TD>"
			lsHTML = lsHTML & "</TR>"


		elseIf mnLastPage = 1 Then
			lsHTML = lsHTML & "<TR class=""ExcelFooterCell"" align=""center"">"
			lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
			lsHTML = lsHTML & lsLinks
			lsHTML = lsHTML & "</TD>"
			lsHTML = lsHTML & "</TR>"

			lsHTML = lsHTML & "<TR class=""ExcelFooterCell"" align=""center"">"
			lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
			lsHTML = lsHTML & "<INPUT TYPE=BUTTON VALUE=""Done"" ONCLICK=""sendValue()"" class=""ActionButton"" id=BUTTON1 name=BUTTON1>"
			lsHTML = lsHTML & "</TD>"
			lsHTML = lsHTML & "</TR>"

		End If

		' Do Search
		lsLinks = ""
		lsHTML = lsHTML & "<TR class=""ExcelFooterCell"">"
		lsHTML = lsHTML & "<TD colspan=""" & lnColSpan & """>"
		lsHTML = lsHTML & "Search By <select size=""1"" name=""SearchBy"" class=""FormElem"" onChange=""document.FormName.Query.value=''"">"
		For lnSearch = 0 To mnSearchFieldCount - 1
			if msSearchValueAry(lnSearch) = Server.HTMLEncode(msSearch) then
				lsHTML = lsHTML & "<OPTION value=""" & msSearchValueAry(lnSearch) & """ selected>" & msSearchFieldAry(lnSearch) & "</OPTION>"
			else
				lsHTML = lsHTML & "<OPTION value=""" & msSearchValueAry(lnSearch) & """>" & msSearchFieldAry(lnSearch) & "</OPTION>"
			end if
		Next
		lnSearch = 0
		lsHTML = lsHTML & "</Select>"

		'if msEnableField = "2" then
		'	lsHTML = lsHTML & " Item Type <select size=""1"" name=""selIType"" class=""FormElem"" onChange=""DeleteNodes();showpage('" & PageQuery(0) & "')"" DISABLED>"
		'else
		'	lsHTML = lsHTML & " Item Type <select size=""1"" name=""selIType"" class=""FormElem"" onChange=""DeleteNodes();showpage('" & PageQuery(0) & "')"">"
		'end if
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

		lsHTML = lsHTML & "<INPUT type=""Hidden"" name=""hSelectMode"" value=""" & msOptDispField & """ >"

		lsHTML = lsHTML & "</TD>"
		lsHTML = lsHTML & "</TR>"


		if msOptDispField = "M" then
			lsHTML = lsHTML & "<TR class=""ExcelFooterCell"">"
			lsHTML = lsHTML & "<TD colspan=""" & lncolSpan & """>"
			lsHTML = lsHTML & "Selected Entries &nbsp;&nbsp;  <Span id=idSelList></span>"
			lsHTML = lsHTML & "</TD>"
			lsHTML = lsHTML & "</TR>"
		end if 'if msOptDispField = "M" then

		lsHTML = lsHTML & "</FORM>"

		GetFooter = lsHTML
	End Function

End Class
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
<SCRIPT LANGUAGE=vbscript>
Dim ObjTemp

Set ObjTemp = window.dialogArguments

'********************************************************************************

dim sRet,sButtonPressed
sRet = "-1:0"
function showpage(sArguments)
	sRet = sArguments&"&Query="&trim(document.FormName.Query.value)&"&SearchBy="&trim(document.FormName.SearchBy.value)
	'alert(sret)
	sButtonPressed = "Page"
	set Root = ObjTemp.DocumentElement
	Root.SetAttribute "Action","Page"
	Root.SetAttribute "PassQuery",sRet
	window.close
end function

'********************************************************************************
Function SendValue()
	sButtonPressed = "Done"
	set Root = ObjTemp.DocumentElement
	Root.SetAttribute "Action","Done"
	window.close
End Function
'XML Function
'********************************************************************************
Function XmlFun(Obj)

	Dim Root,node1
	Dim n1,Arr1,sSelectMode


	sSelectMode = document.FormName.hSelectMode.value

	' alert(Obj.value)
	n1 = Obj.value
	Arr1 = split(trim(n1),":")

	set Root = ObjTemp.DocumentElement
	if sSelectMode = "M" then
		if Obj.checked then
			set node1 = ObjTemp.createElement("Reference")
			node1.SetAttribute "ReferenceCode",Arr1(0)
			node1.SetAttribute "ReferenceDate",Arr1(1)
			node1.SetAttribute "ReferenceType",Arr1(2)
			node1.SetAttribute "OtherReference",Arr1(3)
			node1.SetAttribute "Remarks",Arr1(4)
			node1.SetAttribute "ReferenceNo",Arr1(5)
			node1.SetAttribute "OtherRefNoDate",Arr1(6)
			Root.appendchild node1
		else
			for each temp in Root.childnodes
				if Strcomp(temp.nodename,"Reference")=0 then
					if trim(Temp.getAttribute("ReferenceCode")) =  trim(Arr1(0))  then
						Root.Removechild temp
					end if
				end if
			next
		end if
		DispList()
	else

		for each temp in Root.childnodes
			if Strcomp(temp.nodename,"Reference")=0 then
				Root.Removechild temp
			end if
		next

		if Obj.checked then
			set node1 = ObjTemp.createElement("Reference")
			node1.SetAttribute "ReferenceCode",trim(Arr1(0))
			node1.SetAttribute "ReferenceDate",Arr1(1)
			node1.SetAttribute "ReferenceType",Arr1(2)
			node1.SetAttribute "OtherReference",Arr1(3)
			node1.SetAttribute "Remarks",Arr1(4)
			node1.SetAttribute "ReferenceNo",Arr1(5)
			Root.appendchild node1
		end if
	end if
	'alert(Root.xml)
End Function
'********************************************************************************
Function Init()

	sButtonPressed = ""
	set Root = ObjTemp.DocumentElement
	 'alert(Root.xml)
	for i = 0 to document.FormName.elements.length - 1
		if document.FormName.elements(i).type = "checkbox" then
			if Root.haschildnodes then
				for each temp in Root.childnodes
					if ucase(temp.nodename) = ucase("Reference") then
						n1 = trim(document.FormName.elements(i).value)
						TempArr = split(n1,":")
						'alert(trim(Temp.getAttribute("ReferenceNo"))&" = "&  trim(TempArr(3)))
						if trim(Temp.getAttribute("ReferenceNo")) =  trim(TempArr(5)) then
							document.FormName.elements(i).checked=true
							exit for
						end if 'if trim(temp.getAttribute("ItemCode")) = trim(TempArr(1)) then

					end if 'if Strcomp(temp.nodename,"Reference")= 0 then
				next
			end if 'if Root.haschildnodes then

		end if 'if document.FormName.elements(i).type = "checkbox" then
	next

	sSelectMode = document.FormName.hSelectMode.value
	if sSelectMode = "S" then
		DispList()
	end if

End function
'********************************************************************************
Function DeleteNodes()
	set Root = ObjTemp.DocumentElement
	for each temp in Root.childnodes
		if Strcomp(temp.nodename,"Reference")=0 then
			Root.Removechild temp
		end if
	next
End function
'********************************************************************************
function RemoveNode(this)
	'alert(this.value)
	Dim Root,node1
	Dim n1,Arr1

	n1 = this.value ' company item code : item code : class code : class name : item name
	Arr1 = split(n1,":")

	set Root = ObjTemp.DocumentElement
	if this.checked = false then
		for each temp in Root.childnodes
			if Strcomp(temp.nodename,"Reference")=0 then
				if trim(Temp.getAttribute("ReferenceNo")) =  trim(Arr1(0))  then
					Root.Removechild temp
				end if
			end if
		next

		for i = 0 to document.FormName.elements.length - 1
			if document.FormName.elements(i).type = "checkbox"   then
				if document.FormName.elements(i).name = "pKey"   then
					n1 = trim(document.FormName.elements(i).value)
					TempArr = split(n1,":")

					if trim(Arr1(1)) = trim(TempArr(1)) and trim(Arr1(2)) = trim(TempArr(2))  then
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

Function DispList()
Dim s1

	s1 = "<br><TABLE class=""TableOutLineOnly"" cellspacing=""1"" width=""100%"">"
	set Root = ObjTemp.DocumentElement
	'alert(Root.xml)
	if Root.haschildnodes then
		sQ = Root.getAttribute("PassQuery")
		sIType = right(sQ,3)
		'setIndex document.FormName.selIType,sIType
		
		for each temp in Root.childnodes
			if ucase(temp.nodename) = ucase("Reference") then

				s1= trim(s1) & "<tr><td class=ExcelDisplayCell >"
	
				s1= trim(s1) & "<input type=Checkbox name=chk value='" & trim(temp.getAttribute("ReferenceCode")) & ":" & trim(temp.getAttribute("ReferenceDate")) & ":" & trim(temp.getAttribute("Remarks")) & "' checked onClick=RemoveNode(this) >"
				s1= trim(s1) & "</td>"
				s1= trim(s1) & "<td class=ExcelDisplayCell >" & trim(temp.getAttribute("ReferenceCode")) & "</td>"
				s1= trim(s1) & "<td class=ExcelDisplayCell >" & replace(trim(temp.getAttribute("ReferenceDate")),"~~",chr(34)) & "</td>"
				s1= trim(s1) & "<td class=ExcelDisplayCell >" & trim(temp.getAttribute("Remarks")) & "</td>"
				s1= trim(s1) & "</tr>"
			end if 'if Strcomp(temp.nodename,"Reference")= 0 then
		next
	end if 'if Root.haschildnodes then
	'if right(s1,1) = "," then  s1 = mid(s1,1,len(s1) - 1 )

	s1 = trim(s1) + "</table><br>"
	
	idSelList.innerHTML = s1

End Function
'********************************************************************************
Function window_onunload()
	set Root = ObjTemp.DocumentElement
	if trim(sButtonPressed) = "" then
		Root.SetAttribute "Action","CLOSE"

		for each temp in Root.childnodes
			if Strcomp(temp.nodename,"Reference")=0 then
				Root.Removechild temp
			end if
		next
	end if
	
	'alert(ObjTemp.Xml)
	set window.returnValue = ObjTemp.DocumentElement
end Function
'********************************************************************************
Function setIndex(obj,sTemp)
	dim i
	for i = 0 to obj.length - 1
		if trim(sTemp) = trim(obj.options(i).value) then
			obj.selectedIndex = i
			exit function
		end if
	next
End Function
'********************************************
Function ViewDetails(URLValue,RefNo)
    showModalDialog URLValue&RefNo
End Function
'********************************
</SCRIPT>
