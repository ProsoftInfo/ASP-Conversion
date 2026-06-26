<%
	'Program Name				:	DynamicNoSelection.asp
	'Module Name				:	To List all Reference Details
	'Author Name				:	RAGAVENDRAN R
	'Created On					:	Sep 27,2010
	'Tables Used				:
	'Temporary Tables			:
	'Temporary Files			:
	'Input Parameter			:
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


<%Option Explicit%>
<!--#INCLUDE FILE="../include/clsDatabase.asp"-->
<!--#INCLUDE FILE="../include/clsRefNoListMultiple.asp"-->
<!--#include file="../include/populate.asp"-->
<!--#include file="../include/CommonFunctions.asp"-->
<HTML><HEAD><TITLE>Item</TITLE></HEAD>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<BODY>
<%

Dim oDatabase,oDatalist,sTable
dim sRefType,sOrgID,sFilter,sSearchBy,sSelectMode,sFlag,sIType,sPartyCode

Dim sRefCode,sRefDate,sRefName,sOthRef,sRemarks,sRefNo,sQuery,sRefTable,sRefWhereClause,sORGCodeFieldName
Dim rsTemp,sOthRefNoDate,sFilterRuleQuery,sOthRefNo,sViewPath,sRcptIssType,sRefCodeNo
set rsTemp = Server.CreateObject("ADODB.Recordset")
Response.Write "<font color=red>"

'Response.Write "<p><p>Request.QueryString = " & Request.QueryString

sOrgID = trim(Request.QueryString("orgID"))
sRefType = trim(Request.QueryString("RefType"))
sPartyCode = trim(Request.QueryString("ParCode"))
sFilter = trim(Request.QueryString("Query"))
sSearchBy = trim(Request.QueryString("SearchBy"))
sSelectMode=ucase(trim(Request.QueryString("hSelectMode")))
sFlag = trim(Request.QueryString("Flag"))
sOthRefNo = trim(Request.QueryString("OthRefNo"))
sRcptIssType = trim(Request.Querystring("RcptIssType"))
if Trim(sOrgID)="" or IsNull(sOrgID) then
    sOrgID = Session("organizationcode")
end if
sRefCodeNo = Trim(Request.QueryString("RefCodeNo"))

'Response.Write sRefType
 'Response.Write sOrgID
if trim(sSelectMode) = "" then sSelectMode = "M"
'sSelectMode
'Response.Write "<p>Para = " & sSelectMode &"-----"&sFilter
Set oDatabase = New clsDatabase
Set oDatalist = New clsDatalist


oDataList.PageSize = 15

oDataList.PrimaryKey = "ReferenceNo"

oDataList.AddDisplayField "Reference No"

oDataList.AddDisplayField "Reference Date"

oDataList.AddDisplayField "Reference Type"

oDataList.AddDisplayField "Other Reference"

oDataList.AddDisplayField "Remarks"


oDataList.AddEnableField sFlag
oDataList.AddOptDispField sSelectMode

oDataList.AddReturnedField "0" 'Companyitemcode
oDataList.AddReturnedField "1" ' item name
oDataList.AddReturnedField "2"
oDataList.AddReturnedField "3"
oDataList.AddReturnedField "4"
oDataList.AddReturnedField "5"
oDataList.AddReturnedField "6"

oDataList.SearchForDesc "Search For"
oDataList.AddSearchField "Reference No","RN"
oDataList.AddSearchField "Reference Date","RD"

oDataList.LinkField "" '7

'	sQuery = " Select isNull(ReferenceCodeColumnName,ReferenceNoColumnName),isNull(ReferenceDateColumnName,''),isNull(ReferenceName,''),"&_
'			 " isNull(OtherReferenceColumnName,''),isNull(RemarksColumnName,''),isNull(ReferenceNoColumnName,''),"&_
'			 " ReferenceSourceTableName,isNull(WhereClauseText,'') from vw_ReferenceTypes where ReferenceEntryNo = "& sRefType &" and ActivityReferenceStatus='A'"
    sQuery = " Select isNull(ReferenceCodeColumnName,ReferenceNoColumnName),isNull(ReferenceDateColumnName,''),isNull(ReferenceName,''),"&_
			 " isNull(OtherReferenceColumnName,''),isNull(RemarksColumnName,''),isNull(ReferenceNoColumnName,''),"&_
			 " ReferenceSourceTableName,isNull(WhereClauseText,''),isNull(OtherRefNoColumnName,''),isNull(ORGCodeFieldName,''),isNull(FilterRuleQuery,''),IsNull(ViewPath,'') ViewPath from vw_ReferenceTypes where ReferenceEntryNo = "& sRefType &" and ActivityReferenceStatus='A'"

	if trim(sRcptIssType)<>"" then
        sQuery = sQuery & " and TYPE = '"& sRcptIssType &"'"
    end if
    if Trim(sRefCodeNo)<>"" then
        sQuery = sQuery &" and RefCodeNo = "& sRefCodeNo
    end if
	'Response.Write "<textarea>"& sQuery&"</textarea>"

	rsTemp.Open sQuery,con
	if not rsTemp.EOF then
		sRefCode = trim(rsTemp(0))
		sRefDate = trim(rsTemp(1))
		sRefName = trim(rsTemp(2))
		sOthRef  = trim(rsTemp(3))
		sRemarks = trim(rsTemp(4))
		sRefNo   = trim(rsTemp(5))
		sRefTable= trim(rsTemp(6))
		sRefWhereClause = trim(rsTemp(7))
		sOthRefNoDate = trim(rsTemp(8))
		sORGCodeFieldName = Trim(rsTemp(9))
		sFilterRuleQuery = trim(rsTemp(10))
		sViewPath = trim(rsTemp(11))
	end if
	rsTemp.Close
	if trim(sRefCode)="" then sRefCode = Chr(39)&Chr(39)
	if trim(sRefDate)="" then sRefDate = Chr(39)&Chr(39)
	if trim(sRefName)="" then sRefName = Chr(39)&Chr(39)
	if trim(sOthRef) ="" then sOthRef  = Chr(39)&Chr(39)
	if trim(sRemarks)="" then sRemarks = Chr(39)&Chr(39)
	if trim(sRefNo)  ="" then sRefNo   = Chr(39)&Chr(39)
	if trim(sOthRefNoDate)="" then sOthRefNoDate = Chr(39)&Chr(39)
	'if trim(sRefWhereClause)="" then sRefWhereClause = Chr(39)&Chr(39)
	'if Trim(sORGCodeFieldName)="" then sORGCodeFieldName = Chr(39)&Chr(39)
	if trim(sViewPath)="" or IsNull(sViewPath) then sViewPath = Chr(39)&Chr(39)

sRemarks = "''"

	if trim(sRcptIssType)<>"" then
        sRefName = sRefName & " - "& GetRcptIssName(sRcptIssType)
    end if


	if sFilter = "" then
	     if trim(sRefWhereClause)<>"" then
	        sQuery = " Select isNull("& sRefCode &","& sRefNo &") as ReferenceCode,Convert(varchar,"& sRefDate &",103) as ReferenceDate,'"& sRefName &"' as ReferenceType,"&_
						     ""& sOthRef & " as OtherReference,"& sRemarks &" as Remarks,"& sRefNo &" as ReferenceNo, "& sOthRefNoDate  &" as OthRefNoDate,"& sViewPath &" as ViewPath from "&sRefTable &" where "& sRefWhereClause
	     else
	        sQuery = " Select isNull("& sRefCode &","& sRefNo &") as ReferenceCode,Convert(varchar,"& sRefDate &",103) as ReferenceDate,'"& sRefName &"' as ReferenceType,"&_
						 ""& sOthRef & " as OtherReference,"& sRemarks &" as Remarks,"& sRefNo &" as ReferenceNo, "& sOthRefNoDate  &" as OthRefNoDate,"& sViewPath &" as ViewPath from "&sRefTable
	     end if
	'	Response.Write oDatalist.sSql
	elseif sFilter <> "" then
		if Trim(sSearchBy) = "RN" then
			sFilter = "%" & sFilter & "%"
			if trim(sRefWhereClause)<>"" then
		           sQuery = " Select isNull("& sRefCode &","& sRefNo &") as ReferenceCode,Convert(varchar,"& sRefDate &",103) as ReferenceDate,'"& sRefName &"' as ReferenceType,"&_
					             ""& sOthRef & " as OtherReference,"& sRemarks &" as Remarks,"& sRefNo &" as ReferenceNo, "& sOthRefNoDate  &" as OthRefNoDate,"& sViewPath &" as ViewPath from "&sRefTable &" where "&sRefCode&" like '"& sFilter&"' and "& sRefWhereClause
	        else
	               sQuery = " Select isNull("& sRefCode &","& sRefNo &") as ReferenceCode,Convert(varchar,"& sRefDate &",103) as ReferenceDate,'"& sRefName &"' as ReferenceType,"&_
					     ""& sOthRef & " as OtherReference,"& sRemarks &" as Remarks,"& sRefNo &" as ReferenceNo, "& sOthRefNoDate  &" as OthRefNoDate,"& sViewPath &" as ViewPath from "&sRefTable &" where "&sRefCode&" like '"& sFilter&"'"
	        end if

		elseif Trim(sSearchBy) = "RD" then
			sFilter = sFilter
			if trim(sRefWhereClause)<>"" then
            		sQuery = " Select isNull("& sRefCode &","& sRefNo &") as ReferenceCode,Convert(varchar,"& sRefDate &",103) as ReferenceDate,'"& sRefName &"' as ReferenceType,"&_
			                		 ""& sOthRef & " as OtherReference,"& sRemarks &" as Remarks,"& sRefNo &" as ReferenceNo, "& sOthRefNoDate  &" as OthRefNoDate,"& sViewPath &" as ViewPath from "&sRefTable &" where Convert(varchar,"& sRefDate &",103) =  Convert(varchar,'"& sFilter &"',103) and "& sRefWhereClause
			else
					sQuery = " Select isNull("& sRefCode &","& sRefNo &") as ReferenceCode,Convert(varchar,"& sRefDate &",103) as ReferenceDate,'"& sRefName &"' as ReferenceType,"&_
					                 ""& sOthRef & " as OtherReference,"& sRemarks &" as Remarks,"& sRefNo &" as ReferenceNo, "& sOthRefNoDate  &" as OthRefNoDate,"& sViewPath &" as ViewPath from "&sRefTable &" where Convert(varchar,"& sRefDate &",103) =  Convert(varchar,'"& sFilter &"',103)"
			end if
		end if
	end if

if trim(sRefType)="11" then
    if trim(sPartyCode)<>"" then
        if Trim(sFilter)<>"" or Trim(sRefWhereClause)<>"" then
            sQuery = sQuery &" and IssToCode in ('"& sPartyCode &"') and IssToType in ('Party')"
        else
            sQuery = sQuery &" Where IssToCode in ('"& sPartyCode &"') and IssToType in ('Party')"
        end if
    end if
elseif trim(sRefType)="12" then
    if trim(sPartyCode)<>"" then
        if Trim(sFilter)<>"" or Trim(sRefWhereClause)<>"" then
            sQuery = sQuery &" and IssuedToCode in ('"& sPartyCode &"') and IssuedToType in ('Party')"
        else
            sQuery = sQuery &" Where IssuedToCode in ('"& sPartyCode &"') and IssuedToType in ('Party')"
        end if
    end if
elseif (trim(sRefType)="2") then
    if trim(sPartyCode)<>"" then
        if Trim(sFilter)<>"" or Trim(sRefWhereClause)<>"" then
            sQuery = sQuery &" and RFQSentTo in ('"& sPartyCode &"')"
        else
            sQuery = sQuery &" Where RFQSentTo in ('"& sPartyCode &"')"
        end if
    end if
elseif trim(sRefType)<>"1" and trim(sRefType)<>"17" then
    if trim(sPartyCode)<>"" then
        if Trim(sFilter)<>"" or Trim(sRefWhereClause)<>"" then
            sQuery = sQuery &" and PartyCode in ("& sPartyCode &")"
        else
            sQuery = sQuery &" Where PartyCode in ("& sPartyCode &")"
        end if
    end if
end if ' if trim(sRefType)<>"1" then
'Response.Write "sFilter = "& sFilter
'Response.Write "<p>sPartyCode = "& sPartyCode
'Response.Write "<p>sRefWhereClause="& sRefWhereClause
if Trim(sORGCodeFieldName)<>"" and Trim(sPartyCode)="" then
    if Trim(sFilter)<>"" or Trim(sRefWhereClause)<>"" then
        sQuery = sQuery & " and "& sORGCodeFieldName &" = '"& sOrgID  &"'"
    else
        sQuery = sQuery & " where "& sORGCodeFieldName &" = '"& sOrgID &"'"
    end if
elseif Trim(sORGCodeFieldName)<>"" and Trim(sPartyCode)<>"" then
    if trim(sRefType)<>"1" and trim(sRefType)<>"17" then
        sQuery = sQuery & " and "& sORGCodeFieldName &" = '"& sOrgID  &"'"
    else
        if Trim(sFilter)<>"" or Trim(sRefWhereClause)<>""  then
            sQuery = sQuery & " and "& sORGCodeFieldName &" = '"& sOrgID  &"'"
        else
            sQuery = sQuery & " where "& sORGCodeFieldName &" = '"& sOrgID &"'"
        end if
    end if
end if



if trim(sFilterRuleQuery)<>"" then
    sQuery = sQuery & " and "& sRefNo &" in ("& sFilterRuleQuery &")"
end if

if trim(sOthRefNo)<>"" then
    sQuery = sQuery &" and "& sOthRefNoDate &" in ('"& sOthRefNo &"')"
end if

'Response.Write " <textarea>" & sQuery&"</textarea>"
oDataList.sSQL = sQuery

sTable = oDatalist.GetTable(oDatabase)

Set oDatalist = Nothing
Set oDatabase = Nothing

Response.Write sTable
sFilter = ""
%>
</BODY>
</HTML>

