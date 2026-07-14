<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	ItemStockCloseSetupInsert.asp
	'Author Name				:	Ragavendran R
	'Created On					:	April 03,2012
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<!--#include virtual="/include/populate.asp"-->
<%
Dim xmldoc,ndRoot,ndCategory,ndSubCategory,ndClassification
Dim sCategoryCode,sSubCategoryCode,sClassCode,sQuery
Dim iCategoryConsider,iSubCategoryConsider,iClassCodeConsider
con.beginTrans
set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
xmldoc.async = false
xmldoc.load(Request)
xmldoc.save Server.MapPath("../temp/Setup_"& Session.SessionID &".xml")
set ndRoot = xmldoc.documentElement
if ndRoot.hasChildNodes() then
    for each ndCategory in ndRoot.childNodes
        if ndCategory.nodeName="Category" then
            iCategoryConsider = ndCategory.getAttribute("Consider")
            sCategoryCode = ndCategory.getAttribute("Code")
            sQuery = "Update INV_M_ClassificationCategory set ConsiderForYearEndClosing ="& iCategoryConsider &" where CategoryCode = '"& sCategoryCode &"'"
            Response.Write sQuery + vbCrLf
            con.execute sQuery
            if ndCategory.hasChildNodes() then
                for each ndSubCategory in ndCategory.childNodes
                    if ndSubCategory.nodeName="SubCategory" then
                        iSubCategoryConsider = ndSubCategory.getAttribute("Consider")
                        sSubCategoryCode = ndSubCategory.getAttribute("Code")
                        sQuery = "Update INV_M_Classification set ConsiderForYearEndClosing ="& iSubCategoryConsider &" where GroupCode = "& sSubCategoryCode
                        Response.Write sQuery + vbCrLf
                        con.execute sQuery
                        if ndSubCategory.hasChildNodes() then
                            for each ndClassification in ndSubCategory.childNodes
                                if ndClassification.nodeName="Classification" then
                                    iClassCodeConsider = ndClassification.getAttribute("Consider")
                                    sClassCode = ndClassification.getAttribute("Code")
                                    sQuery = "Update INV_M_Classification set ConsiderForYearEndClosing ="& iClassCodeConsider &" where GroupCode = "& sClassCode
                                    Response.Write sQuery + vbCrLf
                                    con.execute sQuery
                                end if
                            next
                        end if 'if ndSubCategory.hasChildNodes() then
                    end if 'if ndSubCategory.nodeName="SubCategory" then
                next
            end if' if ndCategory.hasChildNodes() then
        end if'if ndCategory.nodeName="Catetgory" then
    next
end if


if con.Errors.count <> 0 then
	dim iCounter
	con.RollbackTrans
	for iCounter=0 to con.Errors.count
		Response.Write con.Errors(iCounter) & "<BR>"
	next
	'Redirect to Error Handling System
else
'	con.rollbacktrans
 '   Response.End 
    Response.Clear 
	con.CommitTrans
end if

con.close
set con = nothing
%>
