<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	NoSeriesNoSettings.asp
	'Module Name				:	Admin  (Master)
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	FEB 17,2010
	'Modified On				:
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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
<!--#include file="../../include/DatabaseConnection.asp"-->
<!--#include file="../../include/populate.asp"-->
<!--#include file="../../include/getCurrentDate.asp"-->
<%

	Dim sOrgUnit,dcrs, drSet, sNumberType, iSeriesNo, sNumberDef, sNoSeries,sSql,sSelProd,ictr,sSelPack,sQuery,iSeriesCode
	Dim sSelItem,sSelItemProd,sProduct, sPeriod, drSet1, sProdType, iNoRecord,Objrs,Objrs1
	Dim sTempSeriesCodes,sTempSeriesNos,sTempModule,sTempActivity,sTempItemValue,sItemType,sModule
	Dim sCatCode,sClassCode,sClassName,sCatName
	sProdType = 0
	iNoRecord = 0
	set dcrs = Server.CreateObject("ADODB.RecordSet")
	Set drSet = Server.CreateObject("ADODB.RecordSet")
	Set drSet1 = Server.CreateObject("ADODB.RecordSet")
	set Objrs = Server.CreateObject("ADODB.RecordSet")
	set Objrs1 = Server.CreateObject("ADODB.RecordSet")

	Response.Write "<font color=red>"

	sOrgUnit = Request.Form("selUnit")

	sTempModule = Request.Form("hSelModule")
	stempActivity = Request.Form("hSelActivity")
	sItemType= Request.Form("hItemValue")
	sClassCode = Request.Form("hClassCode")
	sCatCode = Request.Form("hCatCode")
	'Response.Write "<p>CatCode = "& sCatCode
	if sTempModule ="" then
		sTempModule = Request.QueryString("hSelModule")
		sTempActivity = Request.QueryString("hSelActivity")
		sItemType= Request.QueryString("hItemValue")
		sClassCode = Request.QueryString("hClassCode")
	    sCatCode = Request.QueryString("hCatCode")
	end if
	if Trim(sCatCode)="" then sCatCode= Request.Form("selCategory")
	sTempSeriesCodes = Request.Form("hSeriesCodes")
	sTempSeriesNos = Request.Form("hSeriesNos")
	sTempItemValue = sItemType
	sPeriod = GetPeriodInterval(FormatDate(getCurrentDate()),"Y")

	sModule = Request.QueryString("Module")
	if sTempModule ="" then
		sTempModule = sModule
	end if
	if sTempItemValue ="" then
		sTempItemValue = "0"
	end if

	if trim(sOrgUnit)="0" or IsNull(sOrgUnit) or trim(sOrgUnit)="" then sOrgUnit = "010101"
	sClassName = ""
	if Trim(sClassCode)<>"" and Trim(sClassCode)<>"0" then
    	sQuery ="Select GroupName from INV_M_Classification where GroupCode in ("& sClassCode &")"
    	dcrs.Open squery,con
    	if not dcrs.EOF then
    	    sClassName = dcrs(0)
    	end if
    	dcrs.Close
    end if 'if Trim(sClassCode)<>"" and Trim(sClassCode)<>"0" then

    if Trim(sCatCode)<>"" and Trim(sCatCode)<>"0" and Trim(sClassName)="" then
    	sQuery ="Select CategoryName from INV_M_ClassificationCategory where CategoryCode in ("& sCatCode &")"
    	dcrs.Open squery,con
    	if not dcrs.EOF then
    	    sClassName = dcrs(0)
    	end if
    	dcrs.Close
    end if 'if Trim(sClassCode)<>"" and Trim(sClassCode)<>"0" then



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE> Number Series Settings</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<XML ID="SeriesNoData" src="../../NoSeries/xmldata/SeriesNumberDetail.xml"></XML>
<XML ID="NumData"><root/></XML>
<XML id="activity"></XML>
<xml id="SeriesList"></XML>
<XML id="NoSeries"></XML>
<SCRIPT LANGUAGE=javascript SRC="../../scripts/rolloverout.js"></SCRIPT>
<SCRIPT LANGUAGE="vbscript" SRC="../../scripts/ExcelFunctions.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript">
Dim Root,subnode
'---------------------------------------------
Function SubmitItem()
End Function
'**************************************
Function SelectClassifcation()

	OrgID = document.formname.selUnit(document.formname.selUnit.selectedIndex).value
	'ITypeName = document.formname.selIType(document.formname.selIType.selectedIndex).text
	IType = 1
	ReturnData = showModalDialog("/include/ClassificationSelectPop.asp?sIType="&IType&"&sOrgID="&OrgID&"&sITypename="&ITypeName&"&SelMode=M","Classification","dialogHeight:500px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No")
	'ReturnData = window.open("/include/ClassificationSelectPop.asp?sIType="&IType&"&sOrgID="&OrgID&"&sITypename="&ITypeName&"&SelMode=M","Classification","Height=500px;Width=650px;center=Yes;help=No;resizable=No;status=No")
	'ALERT(ReturnData)
	arrTemp = split(ReturnData,"*****")

	if arrTemp(0) = "-1" then exit function


	for i = 0 to ubound(arrTemp)  - 1
		j = 0
		arrTempValue =  split(arrTemp(0),"|")
		for j = 0 to ubound(arrTempValue)
			arrTempClass =  split(arrTempValue(j),":")
			if UBound(ArrTempClass)>0 then
				sClass = sClass & "," & arrTempClass(ubound(arrTempClass))
				sCategory = sCategory & "," & arrTempClass(1)
			else
				'sClass = sClass & "," & mid(arrTempClass(0),4)
				sClass = ""
				sCategory = sCategory & "," & mid(arrTempClass(0),4)
			end if
		next

		arrTempName =  split(arrTemp(1),"|||")
		for z = 0 to ubound(arrTempName)
			arrTempClassName =  split(arrTempName(z),":")
			sClassName = sClassName & "," & arrTempClassName(ubound(arrTempClassName))
		next
	next

	sClass = mid(sClass,2)
	sClassName = mid(sClassName,2)
	sTemp = mid(sTemp,3)
	sCategory = Mid(sCategory,2)

	txtClass.innerText  = sClassName
	document.formname.hClassCode.value = sclass
	document.formname.hCatCode.value = sCategory
	  set objCat = eval("document.formname.selCategory")
	For iCnt = 0 to objCat.length-1
	    if trim(objCat(iCnt).value)=trim(sCategory) then
	        objCat.selectedIndex= iCnt
	        objCat.disabled = true
	        exit for
	    end if
	Next

End Function
'-----------------------------------------------
Function PopulateData()
	if document.formname.selUnit.value = "0" then
		alert("Select Unit")
		document.formname.selUnit.focus
		exit function
	end if
	document.formname.action="NoSeriesNoSettings.asp"
	document.formname.submit
End Function
'--------------------------------------------------------
Function PopulateDetailOne()
Dim sTemp,sTempSerNo,sItemVal,sSeriesCode,sUnit,sActivity,sModule

sTempSerNo = ""
sSeriesCode  =""
sItemVal  =""

	sModule = document.formname.SelModule.value
	sUnit = document.formname.selUnit.value
	sActivity = document.formname.SelActivity.value

	if trim(sActivity)="" then
		alert("No Details are available for this Activity")
		exit function
	end if

	document.formname.hSelModule.value=document.formname.SelModule.value
	document.formname.hSelActivity.value=document.formname.SelActivity.value



	document.formname.action ="NoSeriesNoSettings.asp"
	document.formname.submit
End Function
'------------------------------------------------
Function PopulateDetail()
Dim sTemp,sTempSerNo,sItemVal,sSeriesCode,sUnit,sActivity,sModule

sTempSerNo = ""
sSeriesCode  =""
sItemVal  =""

	sModule = document.formname.SelModule.value
	sUnit = document.formname.selUnit.value
	sActivity = document.formname.SelActivity.value
	set objCat = eval("document.formname.selCategory")
	sCategoryCoode = objCat(objCat.selectedIndex).value
	if sModule <>"6" then
		if trim(sActivity)="" then
			alert("No Details are available for this Activity")
			exit function
		end if
	end if

	document.formname.hSelModule.value=document.formname.SelModule.value
	document.formname.hSelActivity.value=document.formname.SelActivity.value

'	sCatCode = document.formname.hCatCode.value
'	if Trim(sCatCode)="" or Trim(sCatCode)="0" then
'	    document.formname.hCatCode.value = sCategoryCoode
'	end if
	document.formname.action ="NoSeriesNoSettings.asp"
	document.formname.submit
End Function
'----------------------------------------------
Function PopulateActivity(sPara)
	Dim sCount
	if document.formname.selUnit.value ="0" then
		alert("Select Unit")
		document.formname.selUnit.focus
		document.formname.SelModule.value="0"
		exit function
	end if
	if document.formname.SelModule.value="0" then
		alert("Select Module")
		document.formname.SelModule.focus
		exit function
	end if

	Set objhttp = CreateObject("MSXML2.XMLHTTP")
	'alert(document.formname.SelActivity.value)
	objhttp.Open "GET","XMLModuleActivity.asp?sUnitId="& document.formname.hdOrgUnit.value &"&hApplicationNo="& document.formname.SelModule.value , false
	objhttp.send
	'alert(objhttp.responseText)
	if objhttp.responseXML.Xml <> "" then
		activity.loadXML objhttp.responseXML.Xml
	end if
	document.formname.SelActivity.length = 0
	set Root = activity.documentElement
		sCount = document.formname.SelActivity.length
	for each subnode in Root.childnodes
		sCount= sCount + 1
		document.formname.SelActivity.length = sCount
		document.formname.SelActivity(sCount-1).value = subnode.getAttribute("ActCode")
		document.formname.SelActivity(sCount-1).Text = subnode.getAttribute("ActName")
		if trim(sPara)<>"" then
			if strcomp(subnode.getAttribute("ActCode"),sPara)=0 then
				document.formname.SelActivity.selectedIndex  =  sCount - 1
			end if
		end if
	next

End Function
'---------------------------------------------------
Function closeMe()
	window.location.href = "../index_admin.asp"
End Function
'---------------------------------------------
Function updateMe()
Dim hRow,nCount,NoRoot,NoSubNode
Dim sPackNum
Dim nSerNo,nSerCode,nPeriod,nPrefix,nSuffix,nLastPack,nNewPack
hRow = cint(document.formname.hRowCount.value)

set NoRoot = NoSeries.createElement("Root")
NoRoot.setAttribute "OrgCode",document.formname.selUnit.value
NoSeries.appendChild NoRoot

for nCount = 1 to hRow
	set NoSubNode = NoSeries.createElement("Series")
		nSerNo = eval("document.formname.hdSeriesNoZ"&nCount).value
		nSerCode = eval("document.formname.hdSeriesCodeZ"&nCount).value
		nPeriod = eval("document.formname.hdPeriodZ"&nCount).value
		nPrefix = eval("document.formname.hdPrefixZ"&nCount).value
		nSuffix = eval("document.formname.hdSuffixZ"&nCount).value
		nLastPack = eval("document.formname.hdLastPackNumZ"&nCount).value
		nNewPack = eval("document.formname.txtPackNumZ"&nCount).value

		NoSubNode.setAttribute "SeriesNo",nSerNo
		NoSubNode.setAttribute "SeriesCode",nSerCode
		NoSubNode.setAttribute "Period",nPeriod
		NoSubNode.setAttribute "Prefix",nPrefix
		NoSubNode.setAttribute "Suffix",nSuffix
		NoSubNode.setAttribute "LastPack",nLastPack
		if trim(nNewPack)<>"" then
			NoSubNode.setAttribute "NewPack",nNewPack
		else
			NoSubNode.setAttribute "NewPack",""
		end if
		NoRoot.appendChild NoSubNode
next


	for nCount = 1 to hRow
		sPackNum = eval("document.formname.txtPackNumZ"&nCount).value
			If isNaN(Trim(eval("document.formname.txtPackNumZ"&nCount).value)) Then
				MsgBox "Specify numeric value for the packing number"
				eval("document.formname.txtPackNum"&nCount).select()
				Exit Function
			End If
	next

		set objhttp = CreateObject("Microsoft.XMLHTTP")
			objhttp.Open "POST","XMLSave.asp?Name=NoSeriesSettings", false
			objhttp.send NoSeries.XMLDocument
			if objhttp.responseText <> "" then
				Msgbox(objhttp.responseText)
			else
				document.formname.action="NoSeriesNoSettingsUpdate.asp?Module="&document.formname.hselModule.value &"&Activity="& document.formname.hSelActivity.value &"&ClassCode="&document.formname.hClassCode.value&"&CatCode="&document.formname.hCatCode.value
				document.formname.submit()
			end if

End Function
</SCRIPT>
<SCRIPT LANGUAGE="vbscript">
Function selUnit()
Dim sTemp
	document.formname.selUnit.value = document.formname.hdOrgUnit.value
	document.formname.SelModule.value = document.formname.hselModule.value
	if document.formname.SelModule.value <>"" then
		if trim(document.formname.hSelActivity.value)<>"" then
			PopulateActivity(document.formname.hSelActivity.value)
		else
			PopulateActivity("0")
		end if
		document.formname.SelActivity.value = document.formname.hSelActivity.value
	else
		document.formname.SelModule.value = 0
	end if

	set objCat = eval("document.formname.selCategory")
	For iCnt = 0 to objCat.length-1
	    if trim(objCat(iCnt).value)=trim(document.formname.hCatCode.value) then
	        objCat.selectedIndex = icnt
	        exit for
	    end if
	Next

	sClassName = document.formname.hClassName.value
	if Trim(sClassName)="" then
		sClassName =  document.formname.hCatName.Value
	end if
	txtClass.innerText = sClassName
End Function
</Script>
</HEAD>

<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" onLoad="selUnit()" MARGINWIDTH="0">
<form method="POST" name="formname">
<input type="hidden" name="hdOrgUnit" value="<%=sOrgUnit%>">
<input type="hidden" name="hdNoSeries" value="<%=iSeriesNo%>">
<input type=hidden name="hSeriesType" value="<%=sNumberType%>">
<input type=hidden name="hSelProd" value="<%=Request.Form("selProd")%>">
<input type=hidden name="hselPack" value="<%=Request.Form("selPack")%>">
<input type=hidden name="hSeriesLen" value="">
<input type=hidden name="hActivityName" value="">
<input type=hidden name="hItem" value="<%=Request.Form("selItem")%>">
<input type=hidden name="hItemProd" value="<%=Request.Form("selItemProd")%>">
<input type=hidden name="hSeriesNos" value="<%=sTempSeriesNos%>">
<input type=hidden name="hSeriesCodes" value="<%=sTempSeriesCodes%>">
<input type=hidden name="hItemValue" value="<%=sTempItemValue%>">
<input type=hidden name="hSelModule" value="<%=sTempModule%>">
<input type=hidden name="hSelActivity" value="<%=sTempActivity%>">
<input type="hidden" name="hClassCode" value="<%=sClassCode%>" />
<input type="hidden" name="hCatCode" value="<%=sCatCode%>" />
<input type="hidden" name="hClassName" value="<%=sClassName%>" />
<input type="hidden" name="hCatName" value="<%=sCatName%>" />
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr><td height="1px"></td></tr>
	<tr>
	<td class="PageTitle">No Series Setting Change
	</td>
    </tr>
	<tr>
	<td align="center" class="TopPack">
	</td>
	</tr>
	<tr>
	<td valign="top">
			<TABLE id=Table16 cellSpacing=0 cellPadding=0 border=0 width="100%">
				<TR>
					<TD class=TabBodyWithTopLine>
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td align="center" colspan="3" class="MiddlePack">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
							</tr>
							<tr>
								<td align="center">
								</td>
								<td valign="top" width="100%">
									<table cellpadding="0" cellspacing="0">
										<tr>
											<td class="FieldCellSub">Select Unit</td>
											<td class="FieldCellSub" colspan="3">
												<select size="1" name="selUnit" class="FormElem">
													<option value="0">Select</option>
													<%	'Calling the Function which populates Organization Unit list
														populateUnit
													%>
												</select>
											</td>
										</tr>

											<tr>
												<td class="FieldCellSub">Select Module</td>
												<td class="FieldCellSub">
												<Select size="1" name="SelModule" class="FormElem" onChange="PopulateActivity('0')" disabled="true">
													<option value="0">Select</option>
													<%
														populateModule(sTempModule)
													%>
												</Select>
												</td>
											</tr>
											<tr>
											<td class="FieldCellSub">Classification</td>
											<td class="FieldCellSub">
											    <span id="txtClass" class="DataOnly">&nbsp;</span>&nbsp;&nbsp;<a href="#" onclick="SelectClassifcation()"><img style="cursor: hand" src="../../assets/images/iTMS%20Icons/EntryIcon.gif" align="top" width="11" height="11" alt="Select Classification"></a>
											        <select id="selCategory" class="FormElem">
											        <option value="0">Select Category</option>
											        <%
											            sQuery = "Select CategoryCode,CategoryName from Inv_M_ClassificationCategory Order By CategoryCode"
											            Objrs.open sQuery,con
											            if not Objrs.eof then
											                do while not Objrs.eof
											                    Response.Write "<option value="& Objrs(0) &">"& Objrs(1) &"</option>"
											                    Objrs.movenext
											                loop
											            end if
											            Objrs.close
											        %>
											    </select>
											</td>
										</tr>
										<tr>
										    <td class="FieldCellSub">Select Activity</td>
												<td class="FieldCellSub">
												<Select size="1" name="SelActivity" class="FormElem" onChange="PopulateDetail()">
													<option value="0">Select</option>
												</Select>
												</td>
										</tr>

										</table>
										</td>
										</tr>
									<tr>
									<td align="center" colspan="3" class="BottomPack">
									</td>
						        </tr>
								<tr>
								<td align="center">
								</td>
								<td align="center" valign="top">
                                         <table id="tblBook" border="0" cellspacing="1" class="ExcelTable">
                                         <tr>
											<td class="ExcelHeaderCell" align="center" width="10">S.No.</td>
											<td class="ExcelHeaderCell" align="center" width="75">Period</td>
											<td class="ExcelHeaderCell" align="center" width="100">Prefix</td>
											<td class="ExcelHeaderCell" align="center" width="100">Suffix</td>
											<td class="ExcelHeaderCell" align="center" width="100">Last Number</td>
											<td class="ExcelHeaderCell" align="center" width="100">Next Number</td>
											<td class="ExcelHeaderCell" align="center" width="100">Set Number</td>
                                         </tr>
	 							<%

	 							    Response.Write "<font color=red>"
	 							    'Response.Write "sTempModule = "& sTempModule
	 									if sTempModule = 2 then
	 										sQuery = "Select NoSeriesTransactionNo,OrganisationCode,NumberFor,isNull(H.SeriesNo,0),isNull(H.SeriesCode,0) "&_
											 "From PUR_M_Noseries H left join PUR_M_NoSeriesClass D on H.SeriesCode = D.SeriesCode Where H.NoSeriesStatus <> '1'  and OrganisationCode = '"&sOrgUnit&"' and ActivityType = '"&sTempActivity&"' "

											if Trim(sCatCode)<>"" then
	                                            sQuery= sQuery&" and CatCode in ("& sCatCode &")"
	                                        end if
	                                        if Trim(sClassCode)<>"" then
	                                            sQuery= sQuery&" and ClassCode in ("& sClassCode &")"
	                                        end if

											' Response.Write "<p>"& squery
												with Objrs
													.CursorLocation = 3
													.CursorType = 3
													.Source = squery
													.ActiveConnection = con
													.Open
												end with
												if not objrs.eof then
													Do While Not Objrs.EOF
														iSeriesNo = Objrs(3)
														iSeriesCode = objrs(4)

															squery = "Select isNull(ItemType,0),isNull(MainSeriesNo,0),isNull(MainSeriesCode,0),isnull(ItemValue,0) "&_
															 " From VwPurNoSeriesSel Where NoSeriesTransactionNo = "&Objrs(0)&" "

														'	Response.Write  "<p>"& squery
															With Objrs1
																.CursorType = 3
																.CursorLocation = 3
																.ActiveConnection = Con
																.Source = sQuery
																.Open
															End With
															Set Objrs1.ActiveConnection = Nothing
															IF Not Objrs1.EOF Then
																sTempSeriesNos = sTempSeriesNos & ","& objrs1(1)
																sTempSeriesCodes = sTempSeriesCodes & ","& objrs1(2)
															End IF
															objrs1.close
														Objrs.MoveNext
													Loop
												else
													sTempSeriesNos = ""
													sTempSeriesCodes = ""
												end if
												Objrs.Close
	 									end if

	 									if sTempModule = 3 then
	 										sQuery = "Select H.NoSeriesTransactionNo,OrganisationCode,NumberFor,isNull(H.SeriesNo,0),isNull(H.SeriesCode,0) "&_
													" From Sal_M_NoSeries H left join Sal_M_NoSeriesDetails D on H.NoSeriesTransactionNo = D.NoSeriesTransactionNo "&_
                                                    " left join Sal_M_NoSeriesClass C on D.SeriesCode=C.SeriesCode Where H.NoSeriesStatus <> '1'  and "&_
                                                    " OrganisationCode = '"&sOrgUnit&"' and ActivityType = '"&sTempActivity&"' "

                                                    if Trim(sCatCode)<>"" then
	                                                    sQuery= sQuery&" and C.CatCode in ("& sCatCode &")"
	                                                end if
	                                                if Trim(sClassCode)<>"" then
	                                                    sQuery= sQuery&" and C.ClassCode in ("& sClassCode &")"
	                                                end if

												'Response.Write "<p>"&squery
												with Objrs
													.CursorLocation = 3
													.CursorType = 3
													.Source = squery
													.ActiveConnection = con
													.Open
												end with
												if not objrs.eof then
													Do While Not Objrs.EOF

														iSeriesNo = Objrs(3)
														iSeriesCode = objrs(4)

														IF CStr(iSeriesNo) = "0" and CStr(iSeriesCode) = "0" Then

															sQuery = "Select ItemType,InvoiceType,SaleType,CommissionAgent,isNull(DetailsSeriesNo,0), "&_
																	 "isNull(DetailsSeriesCode,0),ItemValue,InvoiceValue,SaleValue,isNull(AgentCode,0), "&_
																	 "isNull(AddSeriesNo,0),isNull(AddSeriesCode,0) From VwSalNumberSeriesSel "&_
																	 "Where NoSeriesTransactionNo = "&Objrs(0)

															'if sItemType<>"select" then squery = squery & " and ItemValue ='"& sItemType&"'"

															'Response.Write "<p>"&squery
															With Objrs1
																.CursorType = 3
																.CursorLocation = 3
																.ActiveConnection = Con
																.Source = sQuery
																.Open
															End With
															Set Objrs1.ActiveConnection = Nothing
															IF Not Objrs1.EOF Then
																sTempSeriesNos = sTempSeriesNos & ","& objrs1(4)
																sTempSeriesCodes = sTempSeriesCodes & ","& objrs1(5)
															End IF

														else
															sTempSeriesNos = sTempSeriesNos & ","& iSeriesNo
														sTempSeriesCodes = sTempSeriesCodes & ","& iSeriesCode
														end if ' IF CStr(iSeriesNo) = "0" and CStr(iSeriesCode) = "0" Then
														objrs1.close
														Objrs.MoveNext
													loop
												else
													sTempSeriesNos = ""
													sTempSeriesCodes = ""
												end if
												Objrs.close
	 									end if


	 									if sTempModule = 1 or sTempModule = 4 or sTempModule = 5 or sTempModule = 6 or sTempModule = 8 or sTempModule = 9 then

	 										if trim(sTempActivity) ="" then sTempActivity = "0"

	 										if sTempModule = 1 then squery = "SELECT DrSeriesNo, DrSeriesCode FROM Acc_M_BookNumberSeries WHERE OUDefinitionID = '" & sOrgUnit & "' and BookCode = "& sTempActivity

	 										if sTempModule = 4 then  squery = "SELECT SERIESNO, SERIESCODE FROM INV_M_NUMBERSERIES WHERE ORGANISATIONCODE = '" & sOrgUnit & "' AND ACTIVITYTYPE = '" & sTempActivity & "'"

	 										if sTempModule = 5 then squery = "SELECT SERIESNO, SERIESCODE FROM MTN_M_NoSERIES WHERE ORGANISATIONCODE = '" & sOrgUnit & "' AND ACTIVITYTYPE = '" & sTempActivity & "'"

	 										if sTempModule = 6 then squery = "SELECT SERIESNO, SERIESCODE FROM PRD_M_PACKINGNUMBERSERIES WHERE ORGANISATIONCODE = '"&sOrgUnit&"'"

	 										if sTempModule = 8 then squery = "SELECT SERIESNO, SERIESCODE FROM FDP_M_NoSERIES WHERE ORGANISATIONCODE = '" & sOrgUnit & "' AND ACTIVITYTYPE = '" & sTempActivity & "'"

	 										if sTempModule = 9 then squery = "SELECT SERIESNO, SERIESCODE FROM TDP_M_NoSERIES WHERE ORGANISATIONCODE = '" & sOrgUnit & "' AND ACTIVITYTYPE = '" & sTempActivity & "'"

	 										if sTempModule = 4  and trim(sItemType)<>"select" then
 												 squery = squery  & " AND ITEMTYPE = '" & sItemType & "'"
	 										end if
	 										if sTempModule = 6 and trim(sItemType)<>"select" then
	 											 squery = squery & " AND ITEMTYPEID = '"&sItemType&"'"
	 										end if
 										'Response.Write squery
	 										with dcrs
												.CursorLocation = 3
												.CursorType = 3
												.Source = squery
												.ActiveConnection = con
												.Open
											end with
											if not dcrs.EOF then
												sTempSeriesCodes =""
												sTempSeriesNos=""
												sTempItemValue  =""
												do while not dcrs.EOF
													sTempSeriesNos = sTempSeriesNos & ","& dcrs(0)
													sTempSeriesCodes = sTempSeriesCodes & ","& dcrs(1)
													dcrs.MoveNext
												loop
											else
												sTempSeriesCodes =""
												sTempSeriesNos=""
												sTempItemValue  =""
											end if
									end if ' if sTempModule = 4 then

	 								sTempSeriesNos = mid(sTempSeriesNos,2)
	 								sTempSeriesCodes = mid(sTempSeriesCodes,2)
	 								sTempItemValue = mid(sTempItemValue,4)
	 								sTempItemValue = pack(sTempItemValue)

	 								if trim(sTempSeriesCodes)=""then
	 										sTempSeriesCodes = 0
	 								end if
	 								if trim(sTempSeriesNos)="" then
	 									sTempSeriesNos = 0
	 								end if

	 								sSql  = "Select SeriesNo,SeriesCode,Period,EntryNo,Prefix,Suffix,Number-1,Number from APP_R_NoSeriesModuleEntry where OUDefinitionID = '"& sOrgUnit &"'  and Period = '"& sPeriod&"'"
	 								sSql = sSql & " and  SeriesCode in("& sTempSeriesCodes&") and SeriesNo in("& sTempSeriesNos&")"
	 									'	Response.Write sSql

									If ((sOrgUnit <> "") And (sSelProd <> "select"))Then
											With drSet
												.CursorLocation = 3
												.CursorType = 3
												.Source = sSQl
												.ActiveConnection = Con
												.Open
											End With
										Set drSet.ActiveConnection = Nothing
										ictr=0
													If Not drSet.EOF Then
														Do While Not drSet.EOF
															ictr=ictr+1
 											%>
 													 <tr>
														<td class="ExcelHeaderCell" align="center" width="25"><%=ictr%></td>
														<td class="ExcelDisplayCell" align="center" width="75"><%=drSet(2)%></td>
														<td class="ExcelDisplayCell" align="center" width="100"><%=drSet(4)%></td>
														<td class="ExcelDisplayCell" align="center" width="100"><%=drSet(5)%></td>
														<td class="ExcelDisplayCell" align="center" width="100"><%=drSet(6)%></td>
														<td class="ExcelDisplayCell" align="center" width="100"><%=drSet(7)%></td>
														<td class="ExcelInputCell" align="left" width="100">
														<input type="text" name="txtPackNumZ<%=ictr%>" size="8" maxlength="8" class="FormElem"></td>
														<input type="hidden" name="hdLastPackNumZ<%=ictr%>" value="<%=drSet(6)%>">
														<input type="hidden" name="hdPeriodZ<%=ictr%>" value="<%=drSet(2)%>">
														<input type="hidden" name="hdSeriesNoZ<%=ictr%>" value="<%=drSet(0)%>">
														<input type="hidden" name="hdSeriesCodeZ<%=ictr%>" value="<%=drSet(1)%>">
														<input type="hidden" name="hdPrefixZ<%=ictr%>" value="<%=drSet(4)%>">
														<input type="hidden" name="hdSuffixZ<%=ictr%>" value="<%=drSet(5)%>">
													</tr>

											<%
														drSet.MoveNext
														Loop
													Else
											%>
													<tr>
														<td class="DataOnly" align="center" colspan="7">
											<%
														iNoRecord = iNoRecord + 1
														Response.Write("No Record(s) Found")
											%>
													</td>
													</tr>
											<%
													End If
													drSet.Close
										end if'If ((sOrgUnit <> "") And (sSelProd <> "select"))Then
								%>
									<input type=hidden name="hRowCount" value="<%=ictr%>">
											</td>
											</tr>
                                       </table>
								</td>
								<td align="center">
								</td>
                            </tr>
                            <tr>
								<td align="center" colspan="3" class="MiddlePack">
								</td>
                            </tr>
                            <tr>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
								<td valign="top">
									<table border="0" cellpadding="0" cellspacing="0" width="100%">
									<tr>
										<td valign="middle" class="ActionCell">
										<%If iNoRecord = 0 Then%>
											<input type="button" value="Update" class="ActionButton" onClick="updateMe()">
										<%Else%>
											<input type="button" value="Close" class="ActionButton" onClick="closeMe()">
										<%End If%>
										</td>
									</tr>
								</table>
								</td>
								<td align="center">
									<img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
								</td>
                            </tr>
							<tr>
								<td align="center" colspan="3" class="BottomPack">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<input type="hidden" name="hdProductType" value="<%=sProdType%>">
</form>
</BODY>
</HTML>

<%
function GetPeriodInterval(sDate,sIntervalType)
'response.write "A " & sIntervalType
dim iMonth,iYear

iMonth=cint(mid(sDate,4,2))
iYear=mid(sDate,7,4)

	select Case sIntervalType
		Case "M"
				if iMonth < 10 then
					iMonth = "0"&iMonth
				end if
		Case "Q"
				if iMonth>=4 and iMonth<=6 then
					iMonth="01"
				end if
				if iMonth>=7 and iMonth<=9 then
					iMonth="02"
				end if
				if iMonth>=10 and iMonth<=12 then
					iMonth="03"
				end if
				if iMonth>=1 and iMonth<=3 then
					iMonth="04"
				end if
		Case "Y"
'				if iMonth<=12 then
'					iYear=cint(iYear)+1
'				end if
'				iMonth="03"
' CHANGED BY SRIDHARAN FOR THE MONTH/PERIOD PROBLEM
				if iMonth >= 4 and iMonth <= 12 then
					iYear = Cint(iYear) + 1
				else
					iYear = Cint(iYear)
				end if
				iMonth="03"
	end select
				GetPeriodInterval=cstr(iYear)&cstr(iMonth)
End function
'--------------------------------------------------------
Function populateModule(sModule)
sSql = "Select ApplicationCode,ApplicationName from Ms_Applications"
	with drSet
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = sSql
		.Open
	end with
	if not drset.EOF then
		do while not drset.EOF
			if sModule = drset(0) then
				Response.Write "<option value="&drset(0)&" selected>"& drset(1) &"</option>"
			else
				Response.Write "<option value="&drset(0)&">"& drset(1) &"</option>"
			end if
			drset.MoveNext
		loop
	end if
	drset.Close
End Function
'----------------------------------------------------------------------
Function popNoSeries()
	with drSet
		.CursorLocation = 3
		.CursorType = 3
		.ActiveConnection = con
		.Source = "Select SeriesNo,CounterType,Description from Ms_NumberSeries"
		.Open
	end with
	if not drset.EOF then
		do while not drset.EOF
			Response.Write "<option value="&drset(0)&">"& drset(2) &"</option>"
			drset.MoveNext
		loop
	end if
	drset.Close
End Function
%>