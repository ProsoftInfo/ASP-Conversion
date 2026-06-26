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

	Private msURLPrefix					' URL prefix is used to create other URL's
	Private msQuery						' Query user is searching for
	Private mnPage						' Page being viewed
	Private msSearch					' Search option user is searching for
	Private mnLastPage					' Last page of results
	Private mnRecordCount				' Number of records found

' ------------------------------------------------------------------------------
	Private Sub Class_Initialize()
		mnReturnFieldCount = 0
		msSearch = Request.QueryString("SearchBy")
		if msSearch = "" then msSearch = "IA"
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
			lsHTML = lsHTML & "<TD align=""center"" width=""20"">&nbsp;</TD>"
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
					 
					for iCounter = 0 to mnReturnFieldCount - 1
						if iCounter = 4 then 'item name
							lvPrimaryValue = lvPrimaryValue & replace(pvDataAry(msReturnFieldAry(iCounter), lnRowIndex),chr(34),"~~") & ":"
						else
							lvPrimaryValue = lvPrimaryValue & pvDataAry(msReturnFieldAry(iCounter), lnRowIndex) & ":"
						end if 	
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
				lsHTML = lsHTML & "<TD>"
				If lnRowIndex <= lnRowCount Then
					lsData = pvDataAry(lnColIndex, lnRowIndex)
					If VarType(lsData) = vbNull Then
						lsHTML = lsHTML & "&nbsp;"
					Else
						If Len(lsData) > 255 Then
							lsData = lsData & Left(lsData, 255) & " ..."
						End If
						lsData = Server.HTMLEncode(lsData)
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
Function SendValue_OLD()
	sButtonPressed = "Done"
	set Root = ObjTemp.DocumentElement
	Root.SetAttribute "Action","Done"
	window.close 
End Function
'XML Function
Function SendValue()
	Dim Node,nValue,sUnit,sParTy,sParSubType,iAccHEad,objhttp,sTemp
	sButtonPressed = "Done"
	set Root = ObjTemp.DocumentElement
	AccHead = "0"
	If Root.hasChildNodes Then
		For Each Node in Root.childNodes 
			nValue = Node.getAttribute("RetField5")
			sParTy = Node.getAttribute("RetField3")
			sParSubType = Node.getAttribute("RetField4")
		Next
	End IF
	
	If nValue = "0" or nValue = "" Then
		sTemp = sParTy & ":" & sParSubType 
		
		set objhttp = CreateObject("MSXML2.XMLHTTP")
		objhttp.Open "GET","../../include/GetAccHead.asp?Data="&sTemp , false
		objhttp.send
		
		iAccHEad = trim(objhttp.responseText)
		Root.setAttribute "AccHead",iAccHEad	
		If iAccHEad = "0" Then
			alert("This party can not be selected as the party subtype is not mapped to any party control account.")
			Exit Function
		End IF
	End IF
	
	Root.SetAttribute "Action","Done"
	window.close 
End Function

'XML Function
'********************************************************************************
Function XmlFun(Obj)

	Dim Root,node1
	Dim n1,Arr1,sSelectMode
	
	
	sSelectMode = document.FormName.hSelectMode.value
	
	'alert(Obj.value)
	n1 = Obj.value
	Arr1 = split(n1,":")	
	set Root = ObjTemp.DocumentElement
	if sSelectMode = "M" then
		if Obj.checked then
			set node1 = ObjTemp.createElement("Entry")
			For iCnt = 0 to UBound(Arr1)
			    node1.SetAttribute "RetField"&iCnt ,Arr1(iCnt)
			Next
			Root.appendchild node1
		else
			for each temp in Root.childnodes 
				if Strcomp(temp.nodename,"Entry")=0 then
					if trim(Temp.getAttribute("RetField1")) =  trim(Arr1(1)) and trim(Temp.getAttribute("RetField2")) =  trim(Arr1(2)) then
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
			set node1 = ObjTemp.createElement("Entry")
			For iCnt = 0 to UBound(Arr1)
			    node1.SetAttribute "RetField"&iCnt ,Arr1(iCnt)
			Next
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
					if ucase(temp.nodename) = ucase("Entry") then
						n1 = trim(document.FormName.elements(i).value)
						TempArr = split(n1,":")
						if trim(Temp.getAttribute("RetField1")) =  trim(TempArr(1)) and trim(Temp.getAttribute("RetField2")) =  trim(TempArr(2)) then
							document.FormName.elements(i).checked=true
							exit for
						end if 'if trim(temp.getAttribute("ItemCode")) = trim(TempArr(1)) then
						
					end if 'if Strcomp(temp.nodename,"Item")= 0 then
				next
			end if 'if Root.haschildnodes then

		end if 'if document.FormName.elements(i).type = "checkbox" then
	next
	
	sSelectMode = document.FormName.hSelectMode.value
	if sSelectMode = "M" then
		DispList()
	end if	
	
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
			
			if Strcomp(temp.nodename,"Entry")=0 then
				if trim(Temp.getAttribute("RetField1")) =  trim(Arr1(1)) and trim(Temp.getAttribute("RetField2")) =  trim(Arr1(2)) then
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
	if Root.haschildnodes then
		for each temp in Root.childnodes
			if ucase(temp.nodename) = ucase("Entry") then
				
				s1= trim(s1) & "<tr><td class=ExcelDisplayCell >"
				
				s1= trim(s1) & "<input type=checkbox name=chk value='" & trim(temp.getAttribute("RetField0")) & ":" & trim(temp.getAttribute("RetField1")) & ":" & trim(temp.getAttribute("RetField2")) & ":" & trim(temp.getAttribute("RetField3")) & ":" & replace(trim(temp.getAttribute("RetField4")),"~~",chr(34)) & ":" & trim(temp.getAttribute("RetField5")) & "' checked onClick=RemoveNode(this) >" 
				s1= trim(s1) & "</td>"
				s1= trim(s1) & "<td class=ExcelDisplayCell >" & trim(temp.getAttribute("RetField0")) & "</td>"
				s1= trim(s1) & "<td class=ExcelDisplayCell >" & replace(trim(temp.getAttribute("RetField3")),"~~",chr(34)) & "</td>"
				's1= trim(s1) & "<td class=ExcelDisplayCell >" & trim(temp.getAttribute("RetField4")) & "</td>"
				s1= trim(s1) & "</tr>"
			end if 'if Strcomp(temp.nodename,"Item")= 0 then
		next
	end if 'if Root.haschildnodes then
	'if right(s1,1) = "," then  s1 = mid(s1,1,len(s1) - 1 )
	
	s1 = trim(s1) + "</table><br>"
	idSelList.innerHTML = s1			
			
End Function
'********************************************************************************
Function window_onunload()
	if trim(sButtonPressed) = "" then
		set Root = ObjTemp.DocumentElement
		Root.setAttribute "AccHead","0"
		Root.SetAttribute "Action","CLOSE"
		
		for each temp in Root.childnodes	
			if Strcomp(temp.nodename,"Entry")=0 then
				Root.Removechild temp
			end if
		next
		
	end if  
	'alert(ObjTemp.Xml)
set window.returnValue = ObjTemp.DocumentElement
end Function
'********************************************************************************
</SCRIPT>
<SCRIPT LANGUAGE=javascript>
(function (window, document) {
	"use strict";

	var objTemp = window.dialogArguments;
	var sButtonPressed = "";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.FormName || document.forms.FormName || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function asArray(collection) {
		return Array.prototype.slice.call(collection || []);
	}

	function attr(node, name) {
		return node && node.getAttribute ? node.getAttribute(name) || "" : "";
	}

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function childElements(node, name) {
		var wanted = name && String(name).toLowerCase();
		return asArray(node && node.childNodes).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function root() {
		return objTemp && objTemp.documentElement || objTemp && objTemp.XMLDocument && objTemp.XMLDocument.documentElement || null;
	}

	function createEntry(values) {
		var node;
		if (objTemp && objTemp.createElement) {
			node = objTemp.createElement("Entry");
		} else if (objTemp && objTemp.XMLDocument && objTemp.XMLDocument.createElement) {
			node = objTemp.XMLDocument.createElement("Entry");
		} else {
			node = document.implementation.createDocument("", "", null).createElement("Entry");
		}
		values.forEach(function (value, index) {
			setAttr(node, "RetField" + index, value);
		});
		return node;
	}

	function dialogId() {
		var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
		return match ? decodeURIComponent(match[1]) : "";
	}

	function returnValue(value) {
		var id;
		window.returnValue = value;
		window.returnvalue = value;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(value);
			return;
		}
		id = dialogId();
		if (id && window.opener && window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
			window.opener.ITMSModernCompat._receiveDialogValue(id, value);
		}
	}

	function closeWithRoot() {
		returnValue(root());
		window.close();
	}

	function selectedMode() {
		var mode = field("hSelectMode");
		return mode ? trim(mode.value).toUpperCase() : "";
	}

	function matchingEntry(entry, parts) {
		return trim(attr(entry, "RetField1")) === trim(parts[1]) && trim(attr(entry, "RetField2")) === trim(parts[2]);
	}

	function removeMatchingEntries(parts) {
		var currentRoot = root();
		if (!currentRoot) {
			return;
		}
		childElements(currentRoot, "Entry").forEach(function (entry) {
			if (matchingEntry(entry, parts)) {
				currentRoot.removeChild(entry);
			}
		});
	}

	function escapeHtml(value) {
		return String(value == null ? "" : value)
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;");
	}

	function escapeAttr(value) {
		return escapeHtml(value).replace(/'/g, "&#39;");
	}

	function getText(value) {
		return trim(value).replace(/~~/g, '"');
	}

	window.showpage = function (sArguments) {
		var query = field("Query");
		var searchBy = field("SearchBy");
		var currentRoot = root();
		var ret = String(sArguments || "") + "&Query=" + encodeURIComponent(trim(query && query.value)) + "&SearchBy=" + encodeURIComponent(trim(searchBy && searchBy.value));
		sButtonPressed = "Page";
		setAttr(currentRoot, "Action", "Page");
		setAttr(currentRoot, "PassQuery", ret);
		closeWithRoot();
	};

	window.SendValue_OLD = function () {
		sButtonPressed = "Done";
		setAttr(root(), "Action", "Done");
		closeWithRoot();
	};

	window.SendValue = function () {
		var currentRoot = root();
		var lastNodeValue = "";
		var parType = "";
		var parSubType = "";
		var xhr;
		var accHead;
		sButtonPressed = "Done";
		childElements(currentRoot).forEach(function (node) {
			lastNodeValue = attr(node, "RetField5");
			parType = attr(node, "RetField3");
			parSubType = attr(node, "RetField4");
		});
		if (lastNodeValue === "0" || trim(lastNodeValue) === "") {
			xhr = new XMLHttpRequest();
			xhr.open("GET", "../../include/GetAccHead.asp?Data=" + encodeURIComponent(parType + ":" + parSubType), false);
			xhr.send(null);
			accHead = trim(xhr.responseText);
			setAttr(currentRoot, "AccHead", accHead);
			if (accHead === "0") {
				alert("This party can not be selected as the party subtype is not mapped to any party control account.");
				return;
			}
		}
		setAttr(currentRoot, "Action", "Done");
		closeWithRoot();
	};
	window.sendValue = window.SendValue;
	window.sendValue_OLD = window.SendValue_OLD;

	window.XmlFun = function (obj) {
		var currentRoot = root();
		var values = String(obj && obj.value || "").split(":");
		if (!currentRoot) {
			return;
		}
		if (selectedMode() === "M") {
			if (obj.checked) {
				currentRoot.appendChild(createEntry(values));
			} else {
				removeMatchingEntries(values);
			}
			window.DispList();
			return;
		}
		childElements(currentRoot, "Entry").forEach(function (entry) {
			currentRoot.removeChild(entry);
		});
		if (obj.checked) {
			currentRoot.appendChild(createEntry(values));
		}
	};

	window.Init = function () {
		var currentRoot = root();
		var checks = asArray(document.querySelectorAll('input[type="checkbox"]'));
		sButtonPressed = "";
		checks.forEach(function (check) {
			var parts = String(check.value || "").split(":");
			childElements(currentRoot, "Entry").forEach(function (entry) {
				if (matchingEntry(entry, parts)) {
					check.checked = true;
				}
			});
		});
		if (selectedMode() === "M") {
			window.DispList();
		}
	};

	window.RemoveNode = function (obj) {
		var values = String(obj && obj.value || "").split(":");
		if (obj && obj.checked) {
			return;
		}
		removeMatchingEntries(values);
		asArray(document.querySelectorAll('input[type="checkbox"][name="pKey"]')).forEach(function (check) {
			var parts = String(check.value || "").split(":");
			if (trim(values[1]) === trim(parts[1]) && trim(values[2]) === trim(parts[2])) {
				check.checked = false;
			}
		});
		window.DispList();
	};

	window.DispList = function () {
		var currentRoot = root();
		var host = document.getElementById("idSelList") || window.idSelList;
		var html = '<br><TABLE class="TableOutLineOnly" cellspacing="1" width="100%">';
		childElements(currentRoot, "Entry").forEach(function (entry) {
			var value = [
				attr(entry, "RetField0"),
				attr(entry, "RetField1"),
				attr(entry, "RetField2"),
				attr(entry, "RetField3"),
				getText(attr(entry, "RetField4")),
				attr(entry, "RetField5")
			].join(":");
			html += '<tr><td class="ExcelDisplayCell">';
			html += '<input type="checkbox" name="chk" value="' + escapeAttr(value) + '" checked onclick="RemoveNode(this)">';
			html += '</td><td class="ExcelDisplayCell">' + escapeHtml(attr(entry, "RetField0")) + '</td>';
			html += '<td class="ExcelDisplayCell">' + escapeHtml(getText(attr(entry, "RetField3"))) + '</td></tr>';
		});
		html += "</table><br>";
		if (host) {
			host.innerHTML = html;
		}
	};

	window.window_onunload = function () {
		var currentRoot = root();
		if (trim(sButtonPressed) === "") {
			setAttr(currentRoot, "AccHead", "0");
			setAttr(currentRoot, "Action", "CLOSE");
			childElements(currentRoot, "Entry").forEach(function (entry) {
				currentRoot.removeChild(entry);
			});
		}
		returnValue(currentRoot);
	};

	window.addEventListener("beforeunload", function () {
		window.window_onunload();
	});
}(window, document));
</SCRIPT>


