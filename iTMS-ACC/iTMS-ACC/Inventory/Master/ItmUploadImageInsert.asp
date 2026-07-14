<%@ EnableSessionState=true%> 
<!--#include virtual="/include/DatabaseConnection.asp"-->
<%
Response.Expires = -10000
Server.ScriptTimeOut = 300

Set theForm = Server.CreateObject("ABCUpload.XForm")
theForm.Overwrite = True
theForm.MaxUploadSize = 8000000

theForm.ID = Request.QueryString("ID")
Set theFieldThumb = theForm("imgThumb")(1)
Set theFieldBlow = theForm("imgBlow")(1)
iItemCode = theForm("hItemCode")

    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "SELECT ItemThumbNailPic,ItemBlowUpPic from INV_M_ItemMaster where ItemCode = "&iItemCode, con, 1, 3
    If theFieldThumb.FileExists Then 
          rs("ItemThumbNailPic").Value = theFieldThumb.Data
    end if
    If theFieldBlow.FileExists Then 
          rs("ItemBlowUpPic").Value = theFieldBlow.Data
    end if
    rs.Update
    rs.Close


   ' sQuery = "Update INV_M_ITemMaster set "
   ' Response.write sQuery
   ' If theFieldThumb.FileExists Then
   '     'Response.write theFieldThumb.FileName
'	    'theFieldThumb.Save theFieldThumb.FileName
'	    response.write lbound(theFieldThumb.Data)
'	    sQuery = sQuery &" ItemThumbNailPic ='"& theFieldThumb.Data &"',"
 '   End If
  '  if theFieldBlow.fileExists then
 '       'Response.write theFieldBlow.FileName
  '      'theFieldBlow.Save theFieldBlow.FileName
  '      sQuery = sQuery &" ItemBlowUpPic='"& theFieldBlow.Data &"'"
  '  else
  '      sQuery = mid(sQuery,1,len(sQuery)-1)
  '  end if

'    sQuery = sQuery &" where ItemCode = "& iItemCode
'con.execute sQuery
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>Home</TITLE>
<META http-equiv=Content-Type content="text/html; charset=ISO-8859-1">
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
<LINK REL="STYLESHEET" HREF="../../assets/styles/StandardBody.css" TYPE="text/css">
<SCRIPT>
function FunClose() {
	window.close();
}
</SCRIPT>
</HEAD>
<BODY leftMargin=0 topMargin=0 MARGINHEIGHT="0" MARGINWIDTH="0">
<form method="POST" name="formname" action="">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
    <tr>
		<td align="center" class=PageTitle height="20"><p align="center">Upload Image
		</td>
    </tr>
	<tr>
		<td align="center" class="TopPack">
		</td>
	</tr>
    <tr>
	    <td class="TabBodyWithTopLine">
		    <table border="0" cellpadding="0" cellspacing="0" width="100%">
		        <tr>
				    <td align="center" class="toppack" colspan="3">
				    </td>
                </tr>
                <tr>
		            <td align="center"  class="FieldCell" colspan="3"> Upload Complete...
                </tr>
                <tr>
				    <td align="center" class="bottompack" colspan="3">
				    </td>
                </tr>
                <tr>
				    <td align="center" class="ClearPixel">
					    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
				    </td>
    				
				    <td valign="top">
								    <table border="0" cellpadding="0" cellspacing="0" width="100%">
									    <tr>
										    <td valign="middle" class="ActionCell">
											    <p align="center">
                                    		    <input type="button" value="Close" onClick="FunClose()" name="B13" class="ActionButton" tabindex="3" >
										    </td>
									    </tr>
								    </table>
				    </td>
				    <td align="center" class="ClearPixel">
					    <img border="0" src="../../assets/images/clearpixel.gif" width="5" height="5">
				    </td>
			    </tr>
                 <tr>
				    <td align="center" class="bottompack" colspan="3">
				    </td>
                </tr>
		    </table>
	    </td>
    </tr>
</table>
</form>
</BODY>
</HTML>
