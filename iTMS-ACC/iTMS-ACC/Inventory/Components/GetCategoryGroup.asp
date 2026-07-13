<%@ Language=VBScript %>

<%
	Response.Expires=10
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "no-cache"
%>
<%
	'Program Name				:	GetCategoryGroup.asp
	'Module Name				:	Inventory (Item Creation)
	'Author Name				:	MOHAMMED ASIF
	'Created On					:	November 18, 2002
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
	'							:
	'Connects To				:	Component(Tree view for Classification)
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

<!-- #include File="../../include/DBConnectionLogin.asp" -->
<%
	dim dcrs,sql,scatCode,scatName,sGCode,sGName,sGChildCount,sPGroup,sGroupCat,stempkey
	dim scategorycode,sItemType,sIType,sOrgID
	dim OutData,newElem,newElem1,arrTemp,iCtr,sClassIn

	set dcrs = Server.CreateObject("ADODB.Recordset")

	Set OutData = Server.CreateObject("Microsoft.XMLDOM")

	sIType = trim(Request("sIT"))
	sOrgID = trim(Request("sOrgID"))

	if sIType = "NO" and sOrgID = "NO" then
		Set newElem = OutData.createElement("Root")
		OutData.appendChild newElem

		Response.ContentType="text/xml"
		Response.Write OutData.xml
		Response.End
	end if

	if sIType = "NO" then sIType = ""
	if sOrgID = "" or sOrgID = "NO" then sOrgID = "NULL"

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		if sIType = "" or sIType = "select" then
			.Source = "SELECT * FROM INV_M_CLASSIFICATIONCATEGORY ORDER BY CATEGORYCODE"
		else
			.Source = "SELECT * FROM INV_M_CLASSIFICATIONCATEGORY WHERE CATEGORYCODE IN (SELECT GROUPCATEGORY FROM INV_M_CLASSIFICATION WHERE ITEMTYPEID = '" & sIType & "' AND GROUPCATEGORY IS NOT NULL) ORDER BY CATEGORYCODE"
		end if
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing
	set scatCode = dcrs(0)
	set scatName = dcrs(1)

	Set newElem = OutData.createElement("Root")

	do while not dcrs.EOF

		Set newElem1 = OutData.createElement("Group")

		newElem1.setAttribute "Code", "GRP"
		newElem1.setAttribute "Key", "CAT"  & trim(scatCode)
		newElem1.setAttribute "Text", trim(scatName)

		newElem.appendChild newElem1

		dcrs.MoveNext
	loop
	dcrs.close

	if sOrgID <> "NULL" then
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			if sIType = "" or sIType = "select" then
				.Source = "SELECT DISTINCT ITEMPATH FROM INV_M_ITEMORGNGROUP WHERE ORGANISATIONCODE = '" & sOrgID & "' AND CLASSIFICATIONCODE IN (SELECT DISTINCT CLASSIFICATIONCODE FROM VWITEM WHERE ORGANISATIONCODE = '" & sOrgID & "')"
			else
				.Source = "SELECT DISTINCT ITEMPATH FROM INV_M_ITEMORGNGROUP WHERE ORGANISATIONCODE = '" & sOrgID & "' AND CLASSIFICATIONCODE IN (SELECT DISTINCT CLASSIFICATIONCODE FROM VWITEM WHERE ORGANISATIONCODE = '" & sOrgID & "' AND ITEMTYPEID = '" & sIType & "')"
			end if
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
	    Do While Not dcrs.EOF
			iCtr = 2
			arrTemp = split(trim(dcrs(0)),":")
			for iCtr = 2 to UBound(arrTemp)
				sClassIn = sClassIn & "," & arrTemp(iCtr)
			next
			Erase arrTemp
		dcrs.MoveNext
		loop
		dcrs.close
		sClassIn = mid(sClassIn,2)
	else
		with dcrs
			.CursorLocation = 3
			.CursorType = 3
			if sIType = "" or sIType = "select" then
				.Source = "SELECT DISTINCT GROUPCODE FROM INV_M_CLASSIFICATION"
			else
				.Source = "SELECT DISTINCT GROUPCODE FROM INV_M_CLASSIFICATION WHERE ITEMTYPEID = '" & sIType & "'"
			end if
			.ActiveConnection = con
			.Open
		end with
		set dcrs.ActiveConnection = nothing
	    Do While Not dcrs.EOF
			sClassIn = sClassIn & "," & trim(dcrs(0))
		dcrs.MoveNext
		loop
		dcrs.close
		sClassIn = mid(sClassIn,2)
	end if

	sClassIn = uniquearray(sClassIn)
	'Response.Write sClassIn

	with dcrs
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT GROUPCODE,GROUPNAME,CHILDCOUNT,PARENTGROUP,GROUPCATEGORY,ITEMTYPEID FROM INV_M_CLASSIFICATION WHERE GROUPCODE IN (" & sClassIn & ") ORDER BY GROUPCODE"
		.ActiveConnection = con
		.Open
	end with
	set dcrs.ActiveConnection = nothing

	set sGCode = dcrs(0)
	set sGName = dcrs(1)
	set sGChildCount = dcrs(2)
	set sPGroup = dcrs(3)
	set sGroupCat = dcrs(4)
	set sItemType = dcrs(5)

    Do While Not dcrs.EOF
        stempkey = Trim(sGCode)
		If stempkey = Trim(sPGroup) Then

			set newElem1 = OutData.createElement("Group")
			newElem1.setAttribute "Code", "CAT" & Trim(sGroupCat)
			newElem1.setAttribute "Key", Trim(sItemType) & ":" & Trim(sGroupCat) & ":" & Trim(sPGroup)
			newElem1.setAttribute "Text", trim(sGName)
			newElem.appendChild newElem1

		    scategorycode = "" & Trim(sItemType) & ":" & Trim(sGroupCat) & ":" & Trim(sPGroup)

		    stempkey = Trim(sGCode)

		    child stempkey, scategorycode,sOrgID,newElem
		End If
    dcrs.MoveNext
    Loop
    dcrs.Close

	OutData.appendChild newElem

	Response.ContentType="text/xml"
	Response.Write OutData.xml
%>

<%
Private Sub child(sid,smyKey,sOrgID,oNodParent)
    Dim sptempkey, sctempkey,dcrs1
    dim stGCode,stGName

    sptempkey = smyKey

	set dcrs1 = Server.CreateObject("ADODB.Recordset")

	with dcrs1
		.CursorLocation = 3
		.CursorType = 3
		.Source = "SELECT GROUPCODE,GROUPNAME,CHILDCOUNT,PARENTGROUP,GROUPCATEGORY,ITEMTYPEID FROM INV_M_CLASSIFICATION WHERE PARENTGROUP = " & sid & " AND GROUPCODE <> " & sid & " AND GROUPCODE IN (" & sClassIn & ") ORDER BY GROUPCODE"
		.ActiveConnection = con
		.Open
	end with
	set dcrs1.ActiveConnection = nothing

	set stGCode = dcrs1(0)
	set stGName = dcrs1(1)

    Do While Not dcrs1.EOF
        sctempkey = sptempkey & ":" & Trim(stGCode)

		set newElem2 = OutData.createElement("Group")
		newElem2.setAttribute "Code", smyKey
		newElem2.setAttribute "Key", sctempkey
		newElem2.setAttribute "Text", trim(stGName)
		oNodParent.appendChild newElem2

        child Trim(stGCode), smyKey & ":" & Trim(stGCode),sOrgID,oNodParent

    dcrs1.MoveNext
    Loop
    dcrs1.Close
End Sub

%>
<%
function uniquearray(delurparr)
	dim i,temp,newlist,j,f,doc1,arrTemp,bFlag
	i = 0
	temp = ""
	newlist = ""
	arrTemp = split(delurparr,",")

	for j=0 to cint(ubound(arrTemp))
		doc1 = arrTemp(j)
		bFlag = false
		for k=j+1 to cint(ubound(arrTemp))
			temp = arrTemp(k)
			if temp = doc1 then
				bFlag = true
				exit for
			end if
		next
		if not bFlag then
			newlist = newlist&","&doc1
		end if
	next
	newlist = mid(newlist,2)
	'Response.Write "aaaaaaa " & newlist & "  "
	uniquearray = newlist
end function
%>