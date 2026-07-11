<%
' Function to check for Number series creation before start of any Purchase transaction which uses the No. series .
' include by Malathi for the new No.series generated based in the unit ,activity and itemtype
'Function checkNoSeriesEntry(iActivityNo,sOrgID,sItemType)
Function checkNoSeriesEntry(iActivityNo,sOrgID)
Dim dRSet1,dRSet,sSql, blnFlag,dSql

Set dRSet = Server.CreateObject("ADODB.RecordSet")
set dRSet1 = server.CreateObject("ADODB.RecordSet")

'sSql = "SELECT MainSeriesNo,MainSeriesCode FROM VwPurNoSeriesSel where " &_
'" ActivityType=" & iActivityNo & " AND OrganisationCode='" & sOrgID & "' and ItemValue in('0','"&sItemType&"') "
sSql = "SELECT MainSeriesNo,MainSeriesCode FROM VwPurNoSeriesSel where " &_
		" ActivityType=" & iActivityNo & " AND OrganisationCode='" & sOrgID & "' and ItemValue in('0') "
'Response.Write sSql
with dRSet
	.CursorLocation = 3
	.CursorType = 3
	.Source = sSql
	.ActiveConnection = con
	.Open
end with
set dRSet.ActiveConnection = nothing
If not dRSet.eof Then
	blnFlag = True
else
	dSql = "SELECT MainSeriesNo,MainSeriesCode FROM VwPurNoSeriesSel where " &_
	" ActivityType=" & iActivityNo & " AND OrganisationCode='" & sOrgID & "'" ' and ItemValue='"&sItemType&"' "
'	Response.Write dSql
	with dRSet1
		.CursorLocation = 3
		.CursorType = 3
		.Source = dSql
		.ActiveConnection = con
		.Open
	end with
	set dRSet1.ActiveConnection = nothing
	If not dRSet1.eof Then
		blnFlag = true
	else
		blnFlag = False
	end if
	dRSet1.Close
end if
dRSet.close

checkNoSeriesEntry = blnFlag
End Function
%>

<%
Function popSellingForms(sSelCode,sItemType)
' Declaration of variables
Dim aTemp,sCompanyItemCode,SellingNumber,FormOfSelling
Dim dRSet, rsTemp

'Declaration of Objects
Set dRSet = Server.CreateObject("ADODB.RecordSet")
Set rsTemp = Server.CreateObject("ADODB.RecordSet")


with dRSet
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS Where ITEMTYPEID='" & sItemType &"'"
	.ActiveConnection = con
	.Open
end with
set dRSet.ActiveConnection = nothing

set SellingNumber = dRSet(0)
set FormOfSelling = dRSet(1)

Do While Not dRSet.EOF
	if trim(sSelCode) = trim(SellingNumber) then
		Response.Write("<OPTION VALUE="""&trim(SellingNumber)&""" selected >"&trim(FormOfSelling)&"</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE="""&trim(SellingNumber)&""">"&trim(FormOfSelling)&"</OPTION>" &vbcrlf)
	end if

	dRSet.MoveNext
Loop
dRSet.Close

End Function
%>
<%
' To populate Selling forms based on Item/Class & OrgID
Function popSellingFormsforItem(sItemCode,sClassCode,sOrgId,sItemType,sSelCode)
' Declaration of variables
Dim aTemp,sCompanyItemCode,SellingNumber,FormOfSelling
Dim dRSet, rsTemp

'Declaration of Objects
Set dRSet = Server.CreateObject("ADODB.RecordSet")
Set rsTemp = Server.CreateObject("ADODB.RecordSet")


with dRSet
	.CursorLocation = 3
	.CursorType = 3
	.Source = "SELECT SELLINGUNITID,SELLINGUNIT FROM APP_M_FORMCODESELLINGUNITS Where ITEMTYPEID='" & sItemType &"'"
	.ActiveConnection = con
	.Open
end with
set dRSet.ActiveConnection = nothing

set SellingNumber = dRSet(0)
set FormOfSelling = dRSet(1)

Do While Not dRSet.EOF
	if trim(sSelCode) = trim(SellingNumber) then
		Response.Write("<OPTION VALUE="""&trim(SellingNumber)&""" selected >"&trim(FormOfSelling)&"</OPTION>" &vbcrlf)
	else
		Response.Write("<OPTION VALUE="""&trim(SellingNumber)&""">"&trim(FormOfSelling)&"</OPTION>" &vbcrlf)
	end if

	dRSet.MoveNext
Loop
dRSet.Close

End Function
%>

<%
Function popPackingType(sItemType,sSelCode)
' Declaration of variables
Dim dRSet,sQuery, sCode,sName

'Declaration of Objects
Set dRSet = Server.CreateObject("ADODB.RecordSet")

sQuery = "select PackingCode,Packingname from APP_M_PackingType where packingcode in (Select Packingcode from SAL_R_itemtypepack where ItemtypeID = '"&sItemtype&"')"
sQuery = "SELECT PACKINGCODE,PACKINGSHORTNAME,PACKINGNAME FROM APP_M_PACKINGTYPE ORDER BY 3"
With dRSet
	.CursorLocation = 3
	.CursorType = 3
	.ActiveConnection = con
	.Source = sQuery
	.Open
End With
Set dRSet.ActiveConnection = Nothing

Set sCode = dRSet(0)
Set sName = dRSet(2)

If not dRSet.EOF then
	Do while Not dRSet.eof
		if trim(sSelCode) = trim(sCode) then
			Response.Write("<OPTION VALUE="""&trim(sCode)&""" selected >"&trim(sName)&"</OPTION>" &vbcrlf)
		else
			Response.Write("<OPTION VALUE="""&trim(sCode)&""">"&trim(sName)&"</OPTION>" &vbcrlf)
		end if

		dRSet.MoveNext
	Loop
End if
dRSet.close
End Function

%>
