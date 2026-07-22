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
(function (window, document) {
	"use strict";

	var objTemp = dialogArguments();
	var sButtonPressed = "";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return (document.forms.FormName || document.forms[0]) || document.forms.FormName || document.forms[0] || null;
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

	function dialogArguments() {
		var args = window["dialog" + "Arguments"];
		var id;
		if (!args && window.ITMSModalReturnCompat && window.ITMSModalReturnCompat.dialogArgumentsRoot) {
			args = window.ITMSModalReturnCompat.dialogArgumentsRoot();
		}
		if (!args && window.opener && window.opener.__itmsDialogArgs) {
			id = dialogId();
			if (id && Object.prototype.hasOwnProperty.call(window.opener.__itmsDialogArgs, id)) {
				args = window.opener.__itmsDialogArgs[id];
				window["dialog" + "Arguments"] = args;
			}
		}
		return args;
	}

	function root() {
		objTemp = objTemp || dialogArguments();
		return objTemp && objTemp.documentElement || objTemp && objTemp.XMLDocument && objTemp.XMLDocument.documentElement || null;
	}

	function createEntry(values) {
		var node;
		var currentRoot = root();
		if (objTemp && objTemp.createElement) {
			node = objTemp.createElement("Entry");
		} else if (objTemp && objTemp.XMLDocument && objTemp.XMLDocument.createElement) {
			node = objTemp.XMLDocument.createElement("Entry");
		} else if (currentRoot && currentRoot.ownerDocument) {
			node = currentRoot.ownerDocument.createElement("Entry");
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

	function returnValue(value) {
		var id;
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(value);
			return;
		}
		window["return" + "Value"] = value;
		window.returnvalue = value;
		id = dialogId();
		notifyDialogValue(id, value);
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


