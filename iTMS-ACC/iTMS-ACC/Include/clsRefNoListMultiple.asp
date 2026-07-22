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
				AppendLink lsLinks, "<span style=""cursor: pointer"" onClick=""showpage('" & PageQuery(1) & "')"">First</span>", " | "
				' Previous
				AppendLink lsLinks, "<span style=""cursor: pointer"" onClick=""showpage('" & PageQuery(mnPage - 1) & "')"">Previous</span>", " | "
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
				AppendLink lsLinks, "<span style=""cursor: pointer"" onClick=""showpage('" & PageQuery(mnPage + 1) & "')"">Next</span>", " | "
				' Last
				AppendLink lsLinks, "<span style=""cursor: pointer"" onClick=""showpage('" & PageQuery(mnLastPage) & "')"">Last</span>", " | "
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
		lsHTML = lsHTML & "Search By <select size=""1"" name=""SearchBy"" class=""FormElem"" onChange=""(document.forms.FormName || document.forms[0]).Query.value=''"">"
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
		var loForm = (document.forms.FormName || document.forms[0]) || document.forms.FormName || document.forms[0] || null;
		var keys, i;
		if(loForm && loForm.pKeyName)
		{
			lsName = loForm.pKeyName.value;
			keys = loForm.pKey;
			if(!keys)
			{
				return;
			}
			if(keys.length === undefined)
			{
				keys = [keys];
			}
			for(i=0;i<keys.length;i++)
			{
				if(keys[i].checked)
				{
					sRet = keys[i].value;
					window.close();
					return;
				}
			}
		}
	}
	// -->
</SCRIPT>
<SCRIPT>
var ObjTemp = null;
var sRet = "-1:0";
var sButtonPressed = "";

function trim(value) {
	return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
}

function dialogId() {
	var match = String(window.location.search || "").match(/[?&]__itmsDialogId=([^&]+)/);
	return match ? decodeURIComponent(match[1]) : "";
}

function notifyDialogValue(id, value) {
	if (!id || !window.opener) {
		return;
	}
	try {
		if (window.opener.ITMSModernCompat && window.opener.ITMSModernCompat._receiveDialogValue) {
			window.opener.ITMSModernCompat._receiveDialogValue(id, value);
			return;
		}
	} catch (ignoreDirectReturn) {}
	try {
		window.opener.postMessage({ type: "itms-dialog-return", id: id, value: value }, window.location.origin || "*");
	} catch (ignoreMessageReturn) {}
}

function ensureDialogDocument() {
	var args = window["dialog" + "Arguments"];
	var id;
	if (!args && window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.dialogArgumentsRoot) {
		args = window.ITMSModalReturnCompat.dialogArgumentsRoot();
	}
	if (!args) {
		id = dialogId();
		if (id && window.opener && window.opener.__itmsDialogArgs) {
			args = window.opener.__itmsDialogArgs[id];
			window["dialog" + "Arguments"] = args;
		}
	}
	if (args && args.nodeType === 9) {
		return args;
	}
	if (args && args.nodeType === 1) {
		return args.ownerDocument;
	}
	if (args && args.documentElement) {
		return args;
	}
	if (args && args.XMLDocument) {
		return args.XMLDocument;
	}
	return new DOMParser().parseFromString("<Root/>", "text/xml");
}

function root() {
	ObjTemp = ObjTemp || ensureDialogDocument();
	return ObjTemp.documentElement;
}

function childElements(node, nodeName) {
	var result = [];
	var wanted = nodeName ? String(nodeName).toLowerCase() : "";
	for (var i = 0; node && i < node.childNodes.length; i += 1) {
		if (node.childNodes[i].nodeType === 1 && (!wanted || String(node.childNodes[i].nodeName).toLowerCase() === wanted)) {
			result.push(node.childNodes[i]);
		}
	}
	return result;
}

function fieldValue(name) {
	return (document.forms.FormName || document.forms[0]) && (document.forms.FormName || document.forms[0]).elements[name] ? (document.forms.FormName || document.forms[0]).elements[name].value : "";
}

function attr(node, name) {
	return node && node.getAttribute ? node.getAttribute(name) || "" : "";
}

function clearReferences(matchRefNo, matchRefCode) {
	var references = childElements(root(), "Reference");
	for (var i = 0; i < references.length; i += 1) {
		if (matchRefNo == null && matchRefCode == null) {
			root().removeChild(references[i]);
		} else if (matchRefNo != null && trim(attr(references[i], "ReferenceNo")) === trim(matchRefNo)) {
			root().removeChild(references[i]);
		} else if (matchRefNo == null && matchRefCode != null && trim(attr(references[i], "ReferenceCode")) === trim(matchRefCode)) {
			root().removeChild(references[i]);
		}
	}
}

function createReference(parts, includeOtherRefDate) {
	var node = ObjTemp.createElement("Reference");
	node.setAttribute("ReferenceCode", trim(parts[0] || ""));
	node.setAttribute("ReferenceDate", parts[1] || "");
	node.setAttribute("ReferenceType", parts[2] || "");
	node.setAttribute("OtherReference", parts[3] || "");
	node.setAttribute("Remarks", parts[4] || "");
	node.setAttribute("ReferenceNo", parts[5] || "");
	if (includeOtherRefDate) {
		node.setAttribute("OtherRefNoDate", parts[6] || "");
	}
	return node;
}

function returnDialogValue() {
	var value = root();
	var id = dialogId();
	if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
		window.ITMSModernCompat.returnModalValue(value);
		return;
	}
	window["return" + "Value"] = value;
	window.returnvalue = value;
	notifyDialogValue(id, value);
}

function closeWithValue() {
	returnDialogValue();
	window.close();
}

function showpage(sArguments) {
	sRet = sArguments + "&Query=" + encodeURIComponent(trim(fieldValue("Query"))) + "&SearchBy=" + encodeURIComponent(trim(fieldValue("SearchBy")));
	sButtonPressed = "Page";
	root().setAttribute("Action", "Page");
	root().setAttribute("PassQuery", sRet);
	closeWithValue();
}

function SendValue() {
	sButtonPressed = "Done";
	root().setAttribute("Action", "Done");
	closeWithValue();
}

function XmlFun(obj) {
	var selectMode = trim(fieldValue("hSelectMode")).toUpperCase();
	var parts = trim(obj.value).split(":");
	var refNo = parts[5] || "";
	var refCode = parts[0] || "";
	if (selectMode === "M") {
		if (obj.checked) {
			root().appendChild(createReference(parts, true));
		} else {
			clearReferences(refNo || null, refNo ? null : refCode);
		}
		DispList();
	} else {
		clearReferences();
		if (obj.checked) {
			root().appendChild(createReference(parts, false));
		}
	}
}

function Init() {
	var elements = (document.forms.FormName || document.forms[0]) ? (document.forms.FormName || document.forms[0]).elements : [];
	var references = childElements(root(), "Reference");
	sButtonPressed = "";
	for (var i = 0; i < elements.length; i += 1) {
		if ((elements[i].type === "checkbox" || elements[i].type === "radio") && elements[i].name === "pKey") {
			var parts = trim(elements[i].value).split(":");
			for (var n = 0; n < references.length; n += 1) {
				if (trim(attr(references[n], "ReferenceNo")) === trim(parts[5] || "")) {
					elements[i].checked = true;
					break;
				}
			}
		}
	}
	if (trim(fieldValue("hSelectMode")).toUpperCase() === "M") {
		DispList();
	}
}

function DeleteNodes() {
	clearReferences();
}

function RemoveNode(item) {
	var refNo = item.value;
	var elements;
	if (item.checked === false) {
		clearReferences(refNo);
		elements = (document.forms.FormName || document.forms[0]) ? (document.forms.FormName || document.forms[0]).elements : [];
		for (var i = 0; i < elements.length; i += 1) {
			if ((elements[i].type === "checkbox" || elements[i].type === "radio") && elements[i].name === "pKey") {
				var rowParts = trim(elements[i].value).split(":");
				if (trim(refNo) === trim(rowParts[5] || "")) {
					elements[i].checked = false;
					break;
				}
			}
		}
		DispList();
	}
}

function escapeHtml(value) {
	return String(value == null ? "" : value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function escapeAttr(value) {
	return escapeHtml(value).replace(/'/g, "&#39;");
}

function legacyText(value) {
	return trim(value).replace(/~~/g, '"');
}

function DispList() {
	var html = '<br><TABLE class="TableOutLineOnly" cellspacing="1" width="100%">';
	var references = childElements(root(), "Reference");
	var selectedList = document.getElementById("idSelList");
	for (var i = 0; i < references.length; i += 1) {
		var refNo = attr(references[i], "ReferenceNo");
		html += '<tr><td class="ExcelDisplayCell">';
		html += '<input type="checkbox" name="chk" value=\'' + escapeAttr(refNo) + '\' checked onclick="RemoveNode(this)">';
		html += '</td><td class="ExcelDisplayCell">' + escapeHtml(attr(references[i], "ReferenceCode")) + '</td>';
		html += '<td class="ExcelDisplayCell">' + escapeHtml(legacyText(attr(references[i], "ReferenceDate"))) + '</td>';
		html += '<td class="ExcelDisplayCell">' + escapeHtml(attr(references[i], "Remarks")) + '</td></tr>';
	}
	html += "</table><br>";
	if (selectedList) {
		selectedList.innerHTML = html;
	}
}

function window_onunload() {
	if (trim(sButtonPressed) === "") {
		root().setAttribute("Action", "CLOSE");
		clearReferences();
	}
	returnDialogValue();
}

function setIndex(obj, sTemp) {
	if (!obj || !obj.options) {
		return;
	}
	for (var i = 0; i < obj.options.length; i += 1) {
		if (trim(sTemp) === trim(obj.options[i].value)) {
			obj.selectedIndex = i;
			return;
		}
	}
}

function ViewDetails(URLValue, RefNo) {
	var url = String(URLValue || "") + String(RefNo || "");
	var features = "dialogHeight:480px;dialogWidth:640px;center:Yes;help:No;resizable:Yes;status:No";
	if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
		window.ITMSModernCompat.openModalDialog(url, "", features, function () {});
	} else {
		window.open(url, "_blank", "height=480,width=640,resizable=yes,status=no,scrollbars=yes");
	}
}

window.addEventListener("beforeunload", window_onunload);
</SCRIPT>
