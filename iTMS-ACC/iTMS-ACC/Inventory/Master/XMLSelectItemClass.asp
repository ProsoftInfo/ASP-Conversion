<%@ Language=VBScript %>
<%	option explicit	%>
<%
'Response.ExpiresAbsolute=#Now#
Response.Expires=10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	XMLSelectItemClass.asp
	'Module Name				:	Inventory (Reports)
	'Author Name				:	TAJUDEEN S
	'Created On					:	May 28, 2004
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	ClassificationCode, OrganizationID
	'							:
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

<!--#include virtual="/include/DBConnectionLogin.asp"-->

<%
	dim dcrs,OutData,Root,newElem
	dim sOrgID, sClass, sCategory, sText, arrClass, i,arrOrg

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set dcrs = Server.CreateObject("ADODB.Recordset")

	sOrgID = Request.QueryString("sOrgID")
	arrOrg = Split(sOrgID,":")

	sText = Request.QueryString("sText")

	Set Root = OutData.createElement("Classification")
	OutData.appendChild Root

	sOrgID = ""
	for i = 0 to UBound(arrOrg)
		sOrgID = sOrgID&","&"'"&arrOrg(i)&"'"
	next

	sOrgID = trim(mid(sOrgID,2))

	sClass = GetClass(sText)
	arrClass = Split(sClass,",")

	i = 0

	for i = 0 to UBound(arrClass)
		GetChildClass arrClass(i),False
	next

'	OutData.save server.MapPath("../temp/master/Class.xml")
	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>

<%
	Function GetChildClass(iCode,Flag)
		dim dcrs, dcrs1, Sql, iChildCount, iClassCode, sClassName

		set dcrs = Server.CreateObject("ADODB.Recordset")
		set dcrs1 = Server.CreateObject("ADODB.Recordset")

		if Flag then 'Class Contains Child Class
			Sql = "SELECT CHILDCOUNT,GROUPCODE,GROUPNAME FROM INV_M_CLASSIFICATION WHERE PARENTGROUP = " & iCode & " AND GROUPCODE <> PARENTGROUP"
		else 'Class Does not Contains Child Class
			Sql = "SELECT CHILDCOUNT,GROUPCODE,GROUPNAME FROM INV_M_CLASSIFICATION WHERE GROUPCODE = " & iCode
		end if
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = Sql
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		do while not dcrs.eof
			iChildCount = trim(dcrs(0))
			iClassCode =  trim(dcrs(1))
			sClassName =  trim(dcrs(2))
'			Response.Write "ClassCode "  & iClassCode & " ClassName " & sClassName & " Child " & iChildCount & " Called By " & iCode  & "<BR>"

			if iChildCount = 0 then
				'Checking for Item
			'	with dcrs1
			'		.CursorLocation = 3
			'		.CursorType = 3
			'		.Source = "SELECT DISTINCT ITEMCODE FROM INV_M_ITEMORGNGROUP WHERE CLASSIFICATIONCODE =" & iClassCode & " AND ORGANISATIONCODE IN (" & sOrgID & ")"
			'		.ActiveConnection = con
			'		.Open
			'	end with
			'	set dcrs1.ActiveConnection = nothing
'
''				if not dcrs1.eof then
'					'Getting the Category
					GetCategory(iClassCode)

					'Adding Class to the XML
					set NewElem = OutData.createElement("CLASS")
					newElem.setAttribute "CLASSCODE", iClassCode
					newElem.setAttribute "CLASSNAME", sClassName
					newElem.setAttribute "CATEGORY",  sCategory
					Root.appendChild newElem
'				else
''					Response.Write "No Items <BR>"
'				end if
'				dcrs1.close

			elseif iChildCount > 0 then
				'Calling function to find Child Group
				GetChildClass iClassCode , True
			end if
			dcrs.movenext
		loop
		dcrs.close

	End Function
%>

<%
	Function GetCategory(iCode)
		dim dcrs, Sql, iParent, sGroupCategeory

		set dcrs = Server.CreateObject("ADODB.Recordset")

		Sql = "SELECT PARENTGROUP,ISNULL(GROUPCATEGORY,'') FROM INV_M_CLASSIFICATION WHERE GROUPCODE = " & iCode
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			.Source = Sql
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing

		if not dcrs.eof then
			iParent = trim(dcrs(0))
			sGroupCategeory = trim(dcrs(1))
		end if
		dcrs.Close

		if iParent <> iCode then
			'Calling function to find the Parent Class
			GetCategory(iParent)
		else
			sCategory = sGroupCategeory
		end if

	End Function
%>

<%
	Function GetClass(sText)
		dim arrText, arrClassCode(), arrCategory(), arrLen(), arrTemp
		dim sTemp, sTempCategory, i, j, l, iLoc, sClassCode

		'Splitting the given text
		arrText = Split(sText,"|")
		l = ubound(arrText)

		'Splitting Class, Category, Length of the text
		for i= 0 to l
			redim preserve arrLen(i+1), arrClassCode(i+1), arrCategory(i+1)

			arrTemp = Split(arrText(i),":")

			arrLen(i) = len(arrText(i))
			arrClassCode(i) = arrTemp(ubound(arrTemp))
			arrCategory(i)  = arrTemp(1)
			arrTemp = ""
		next

		'Sorting the array based on Length or Category
		for i = 0 to l-1
			for j = 0 to l-1
				if arrLen(j) > arrLen(j+1) or arrCategory(j) > arrCategory(j+1) then
					sTemp = arrText(j)
					arrText(j) = arrText(j+1)
					arrText(j+1) = sTemp

					sTemp = arrCategory(j)
					arrCategory(j) = arrCategory(j+1)
					arrCategory(j+1) = sTemp

					sTemp = arrLen(j)
					arrLen(j) = arrLen(j+1)
					arrLen(j+1) = sTemp

					sTemp = arrClassCode(j)
					arrClassCode(j) = arrClassCode(j+1)
					arrClassCode(j+1) = sTemp
				end if
			next
		next

'		for i = 0 to l
'			Response.Write "Cat " & arrCategory(i) & " Text " & arrText(i) & " Class " & arrClassCode(i) & " Len " & arrLen(i) & " <BR>"
'		next

		sTempCategory = ""

		for i = 0 to l
			if sTempCategory <> arrCategory(i) then
'				Response.Write arrClassCode(i) & ","
				sClassCode = sClassCode & arrClassCode(i) & ","
				sTempCategory = arrCategory(i)
				iLoc = i
			else
				if arrLen(i) = arrLen(iLoc) and arrClassCode(i) <> arrClassCode(iLoc) then
'					Response.Write arrClassCode(i) & ","
					sClassCode = sClassCode & arrClassCode(i) & ","
				end if
			end if
		next

		if sClassCode <> "" then GetClass = left(sClassCode,len(sClassCode)-1)

	End Function
%>

