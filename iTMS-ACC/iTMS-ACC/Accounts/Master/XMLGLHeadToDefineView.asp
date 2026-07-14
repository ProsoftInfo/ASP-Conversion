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
	'Program Name				:	XMLGLHeadToDefineView.asp
	'Module Name				:	Accounts (Master)
	'Author Name				:	Ragavendran R
	'Created On					:	Jan 25,2011
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:	None
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

<!--#include virtual="/include/DatabaseConnection.asp"-->
<html>
<title>GL Heads</title>
<head>
<SCRIPT SRC="../../Common/XMLTreeView.js"></SCRIPT>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css"/>
</head>
<BODY leftMargin="20" topMargin="10" MARGINHEIGHT="0" MARGINWIDTH="0"></BODY>
</html>

<%
	dim rsObj,rsObj1,rsObj2,sSql,OutData,sorgID,Root,newElem,newElem1,newElem2,sTemp,rsConn
	dim iCatagoryCode,sQuery,iAccParentGroup,CurPoss,iAccGroup1,iChildCount,iSubLed
	dim iContra,iETDS,iSumPosting,bSummaryPosting,bEligibleForTDS,bPartySubLed,bContra
	Dim iAccRowCount,iCnt
	dim iDivCounter

    sOrgID = Session("organizationcode")
	Set OutData = Server.CreateObject("Microsoft.XMLDOM")
	set rsObj = Server.CreateObject("ADODB.Recordset")
	set rsObj1 = Server.CreateObject("ADODB.Recordset")
	set rsObj2 = Server.CreateObject("ADODB.Recordset")
	
	iCatagoryCode = "01"
	iDivCounter = 0
    iDivCounter = cint(iDivCounter) + 1
    Set Root = OutData.createElement("menu")
    Root.setAttribute "caption", "GLHead"
    Root.setAttribute "opened", "false"
    Root.setAttribute "icon1", "../../assets/images/home.gif"
    Root.setAttribute "DIVID", iDivCounter
    OutData.appendChild Root
    
	
	sQuery = "Select AccountsGroupCode,AccountsGroupName,ChildCount,GroupCategory,AccountsParentGroup from ACC_M_AccountGroups where AccountsGroupCode = AccountsParentGroup and GroupCategory = '"& iCatagoryCode &"'"
	with rsObj1
		.CursorLocation = 3
    	.CursorType = 3
		.ActiveConnection = con
		.source = sQuery
		.Open
	end with
			do while not rsObj1.EOF
				iAccParentGroup =  rsObj1(0)
				    iDivCounter = cint(iDivCounter) + 1
                        Set newElem = OutData.createElement("menuItem")
                        newElem.setAttribute "caption", trim(rsObj1(1))
                        newElem.setAttribute "opened", "false"
                        newElem.setAttribute "icon1", "../../assets/images/folder-closed.gif"
                        newElem.setAttribute "DIVID", iDivCounter
                        newElem.setAttribute "Description", "UNIT"
                        Root.appendChild newElem
				

					set rsConn = con
					rsConn.CursorLocation = 3
					sQuery = "Select AccountsGroupCode,AccountsGroupName,ChildCount,GroupCategory,AccountsParentGroup from ACC_M_AccountGroups "
					set rsObj = rsConn.execute(sQuery)
					set rsObj.ActiveConnection = Nothing

			'		rsConn.close
					rsObj.Filter = "AccountsGroupCode ='"& iAccParentGroup &"'"
					set rsConn = Nothing

					do while not rsObj.EOF
						CurPoss = rsObj.AbsolutePosition

						iAccGroup1 = rsObj(0)
						iChildCount = rsObj(2)
						
								if trim(sOrgID)<>"0" and trim(sOrgID)<>"" then
									sQuery = "Select AccountHead,AccountHeadCode,AccountDescription,AccountsGroupCode,EligibleForTDS,"&_
									         "SummaryPosting,OUDefinitionID,GroupCategory,AccountsGroupName,SubLedger,EligibleForContras from vwOrgGLHeads where GroupCategory = '"& iCatagoryCode &"' and AccountsGroupCode = '"& iAccGroup1 &"'  and OUDefinitionID = '"& sOrgID &"' "
								else
									sQuery = "Select Distinct AccountHead,AccountHeadCode,AccountDescription,AccountsGroupCode,EligibleForTDS,"&_
											"SummaryPosting,'0',GroupCategory,AccountsGroupName,SubLedger,EligibleForContras from vw_GLHeads where GroupCategory = '"& iCatagoryCode &"' and AccountsGroupCode = '"& iAccGroup1 & "'  "
								end if 'if trim(sOrgID)<>"" then

								if iSubLed = "1" then sQuery = sQuery & " and SubLedger = 1 "

								if iContra = "1" then sQuery = sQuery & " and EligibleForContras = 1 "

								if iETDS = "1" then sQuery = sQuery & " and EligibleForTDS = 1 "

								if iSumPosting = "1" then sQuery = sQuery & " and  SummaryPosting = 1 "

								sQuery= sQuery & " Order By AccountHead"
						'Response.Write sQuery
						rsObj2.Open sQuery,con
						if not rsObj2.EOF then
							do while not rsObj2.EOF
								bSummaryPosting = rsObj2(5)
								bEligibleForTDS = rsObj2(4)
								bPartySubLed = rsObj2(9)
								bContra = rsObj2(10)

								iAccRowCount = iAccRowCount + 1
								    
                                   iDivCounter = cint(iDivCounter) + 1
			                        Set newElem1 = OutData.createElement("menuItem")
			                        newElem1.setAttribute "caption", trim(rsObj2(1))
			                        newElem1.setAttribute "opened", "false"
			                        newElem1.setAttribute "icon1", "../../assets/images/folder-closed.gif"
			                        newElem1.setAttribute "DIVID", iDivCounter
			                        newElem1.setAttribute "Description", "UNIT"
			                        newElem.appendChild newElem1

                                    
								rsObj2.MoveNext
							loop
						else
							if iChildCount="0" then
							end if
						end if
						rsObj2.Close

						For iCnt = 1 to iChildCount
							iAccountGroupCode = iAccGroup1 & Right("0"&cstr(iCnt),2)
							GLAccountHead iAccountGroupCode,2
							set iAccountGroupCode = Nothing
						Next

						rsObj.Filter = "AccountsGroupCode='"& iAccParentGroup &"'"
						rsObj.AbsolutePosition  = CurPoss
						rsObj.MoveNext
					loop

				rsObj1.MoveNext
			loop

	dim xsl,xslt,xslProc

	set xsl = Server.CreateObject("Msxml2.FreeThreadedDOMDocument")
	xsl.async = False
	xsl.Load (Server.MapPath("../../Common/DisplayDetails.xsl"))

	Set xslt = Server.CreateObject("Msxml2.XSLTemplate")
	xslt.stylesheet = xsl
	set xslProc = xslt.createProcessor()
	xslProc.input = OutData

	'APPLY THE TRANSFORMATION AND WRITE THE RESULTS AS OUTPUT
	xslProc.transform()
	Response.Write(xslProc.output)

'	Response.Write OutData.xml

%>
<%

Sub GLAccountHead(iAccGroupCode,Level)
	Dim iChildCnt,iAccGrpCode,iCurPoss,iCount
	rsObj.Filter = "AccountsGroupCode='"&iAccGroupCode &"'"
	do while not rsObj.EOF
		iCurPoss = rsObj.AbsolutePosition
		iAccGroup2 = rsObj(0)
		sSpace = ""

		For iSpCnt = 1 to (nSpace*Level)
			sSpace = sSpace & "&nbsp;"
		Next

		iChildCnt = rsObj(2)

		Response.Write "<tr>"
		Response.Write "<td colspan=8 class=ExcelDisplayCell>"& sSpace & rsObj(1) &"</td></tr>"

		'Response.Write "iChildCnt = "& iChildCnt
		if trim(sOrgID)<>"0" and trim(sOrgID)<>"" then
			sQuery = "Select AccountHead,AccountHeadCode,AccountDescription,AccountsGroupCode,EligibleForTDS,"&_
						"SummaryPosting,OUDefinitionID,GroupCategory,AccountsGroupName,SubLedger,EligibleForContras from vwOrgGLHeads where GroupCategory = '"& iCatagoryCode &"' and AccountsGroupCode = '"& iAccGroup2 & "'  and OUDefinitionID = '"& sOrgID &"' "
		else
			sQuery = "Select Distinct AccountHead,AccountHeadCode,AccountDescription,AccountsGroupCode,EligibleForTDS,"&_
						 "SummaryPosting,'0',GroupCategory,AccountsGroupName,SubLedger,EligibleForContras from vw_GLHeads where GroupCategory = '"& iCatagoryCode &"' and AccountsGroupCode = '"& iAccGroup2 & "' "
		end if 'if trim(sOrgID)<>"" then

		if iSubLed = "1" then sQuery = sQuery & " and SubLedger = 1 "

		if iContra = "1" then sQuery = sQuery & " and EligibleForContras = 1 "

		if iETDS = "1" then sQuery = sQuery & " and EligibleForTDS = 1 "

		if iSumPosting = "1" then sQuery = sQuery & " and  SummaryPosting = 1 "

		sQuery= sQuery & " Order By AccountHead"

		'Response.Write sQuery
		rsObj2.Open sQuery,con
		if not rsObj2.EOF then
			do while not rsObj2.EOF
				bSummaryPosting = rsObj2(5)
				bEligibleForTDS = rsObj2(4)
				bPartySubLed = rsObj2(9)
				bContra = rsObj2(10)
				iAccRowCount = iAccRowCount + 1
				
				iDivCounter = cint(iDivCounter) + 1
                Set newElem1 = OutData.createElement("menuItem")
                newElem1.setAttribute "caption", trim(rsObj2(1))
                newElem1.setAttribute "opened", "false"
                newElem1.setAttribute "icon1", "../../assets/images/folder-closed.gif"
                newElem1.setAttribute "DIVID", iDivCounter
                newElem1.setAttribute "Description", "UNIT"
                newElem.appendChild newElem1
				

				rsObj2.MoveNext
			loop
		else
			if iChildCnt="0" then
			end if
		end if
		rsObj2.Close



		For iCount = 1 to iChildCnt
			iAccGrpCode = iAccGroup2 &right("0"&cstr(iCount),2)
			GLAccountHead iAccGrpCode,Level+1
		Next
		rsObj.Filter = "AccountsGroupCode='"& iAccGroupCode &"'"
		rsObj.AbsolutePosition = iCurPoss
		rsObj.MoveNext
	loop
End Sub
 %>
