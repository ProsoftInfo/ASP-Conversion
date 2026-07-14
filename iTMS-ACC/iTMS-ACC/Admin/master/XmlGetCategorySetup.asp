<%@ Language=VBScript %>
<%	option explicit	%>
<%
	'Program Name				:	XmlGetCategorySetup.asp
%>
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
    Dim objDOM,rsObj,rsSubCat,rsTemp,sQuery,ndRoot,ndCategory,ndSubCategory,ndClassification
    Dim sCategoryCode,sCategoryName,sSubCategoryName,sSubCategoryCode,sClassName,sClassCode
    Dim iYearEndClosingCate,iYearEndClosingSubCate,iYearEneClosingClass
    
    set objDOM =Server.CreateObject("Microsoft.XMLDOM")
    set rsObj = Server.CreateObject("ADODB.Recordset")
    set rsTemp = Server.CreateObject("ADODB.Recordset")
    set rsSubCat = Server.CreateObject("ADODB.Recordset")
    
    set ndRoot = objDOM.createElement("Root")
        objDOM.appendChild ndRoot
    
    sQuery = "Select CategoryCode,CategoryName,ConsiderForYearEndClosing from INV_M_ClassificationCategory"
    rsObj.Open sQuery,con
    if not rsObj.EOF then
        do while not rsObj.EOF 
            sCategoryCode = rsObj(0)
            sCategoryName = rsObj(1)
            iYearEndClosingCate =  rsObj(2)
            set ndCategory = objDOM.createElement("Category")
                ndCategory.setAttribute "Code",sCategoryCode 
                ndCategory.setAttribute "Name",sCategoryName 
                ndCategory.setAttribute "Consider",iYearEndClosingCate 
                ndRoot.appendChild ndCategory
                if iYearEndClosingCate = "0" then
                    sQuery = "Select GroupCode,GroupName,ConsiderForYearEndClosing from INV_M_Classification where GroupCategory = '"& sCategoryCode &"'"
                    rsSubCat.Open sQuery,con
                    if not rsSubCat.EOF then
                        do while not rsSubCat.EOF 
                            sSubCategoryCode = rsSubCat(0)
                            sSubCategoryName = rsSubCat(1)
                            iYearEndClosingSubCate = rsSubCat(2)
                            set ndSubCategory = objDOM.createElement("SubCategory")
                                ndSubCategory.setAttribute "Code",sSubCategoryCode 
                                ndSubCategory.setAttribute "Name",sSubCategoryName 
                                ndSubCategory.setAttribute "Consider",iYearEndClosingSubCate 
                                ndCategory.appendChild ndSubCategory
                                if iYearEndClosingSubCate = "0" then
                                    sQuery = "Select GroupCode,GroupName,ConsiderForYearEndClosing from INV_M_Classification where GroupCode<>ParentGroup and ParentGroup = "& sSubCategoryCode &" and ConsiderForYearEndClosing = 1"
                                    rsTemp.Open sQuery,con
                                    if not rsTemp.EOF then
                                        do while not rsTemp.EOF
                                            sClassCode = rsTemp(0)
                                            sClassName = rsTemp(1)
                                            iYearEneClosingClass = rsTemp(2)
                                            set ndClassification = objDOM.createElement("Classification")
                                                ndClassification.setAttribute "Code",sClassCode 
                                                ndClassification.setAttribute "Name",sClassName 
                                                ndClassification.setAttribute "Consider",iYearEneClosingClass 
                                                ndSubCategory.appendChild ndClassification
                                            rsTemp.MoveNext 
                                        loop
                                    end if
                                    rsTemp.Close 
                                end if 'if iYearEndClosingSubCate = "0" then
                            rsSubCat.MoveNext 
                        loop
                    end if
                    rsSubCat.Close 
                end if ' if iYearEndClosingCate = "0" then
            rsObj.MoveNext 
        loop
    end if
    rsObj.Close 
    
    Response.Clear
    Response.ContentType = "text/xml"
    Response.Write objDOM.xml
%>