<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<Title>iTMS Welcome</Title>
<LINK REL="STYLESHEET" HREF="assets/iTMSGarments.css" TYPE="text/css">
<script language="javascript">
function winClose()
{
	document.forms[0].closeWin.Click();
}
function checkWin()
{
	hdiff=window.screenTop;
	if(hdiff < 0)
		smallscreen()
	else
		fullscreen()
}

function fullscreen(){
	var hdiff;
	top.window.moveTo(-4,-4);
	hdiff=window.screenTop;
	top.window.moveTo(-6,-hdiff-7);
	top.window.resizeTo(screen.width+13,screen.height+hdiff+33);
	document.forms[0].imgTile.alt = "Restore Down"
	document.forms[0].imgTile.src = "assets/WindowTile.gif"
}
function smallscreen(){
	var hdiff;
	top.window.moveTo(0,0);
	hdiff=window.screenTop;
	top.window.resizeTo(screen.width,screen.height);
	document.forms[0].imgTile.alt = "Maximize"
	document.forms[0].imgTile.src = "assets/WindowMaxi.gif"
}
function minScreen() {
	//top.window.resizeTo(0,0);
}
</script>

</HEAD>
<BODY TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 class="MainBack" scroll=No>
<form>
<table border="0" cellpadding="0" width="100%" cellspacing="0" class="MainTable">
  <tr>
    <td width="100%" valign="top">
      <table border="0" cellpadding="0" cellspacing="1" width="100%" >
        <tr>
          <td width="100%" colspan="2">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
              <tr>
                <td><img border="0" src="assets/iTMS.gif" width="55" height="40"></td>
                <td><img border="0" src="assets/iNTEGRATED.gif" width="351" height="40"></td>
                <td background="assets/TopBlank.gif" width="100%">&nbsp;</td>
                <td><img border="0" src="assets/TopEngravedR.gif" width="23" height="40"></td>
                <td><img border="0" src="assets/WindowMin.gif"  width="34" height="40" onClick="minScreen()" alt="Minimize" style="cursor:hand"></td>
                <td><img border="0" src="assets/WindowTile.gif"  width="23" height="40" onClick="checkWin()" alt="Restore Down" style="cursor:hand" name="imgTile"></td>
                <td><img border="0" src="assets/WindowClose.gif"  width="22" height="40" onClick="winClose()" alt="Close" style="cursor:hand"></td>
                <td><img border="0" src="assets/TopEnd.gif" width="4" height="40"></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="50%" class="MainMenuCell">
            <table border="0" cellpadding="0" cellspacing="0" height="16">
              <tr>
              <td width="40" class="MenuCell">
                <OBJECT id="menu" classid=clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11 width=19 height=13 type=application/x-oleobject VIEWASTEXT class="MenuCell">
                    <param name="Width" value="503">
                    <param name="Height" value="344">
                    <param name="Font" value="Verdana, 8,,#000000, ">
                    <param name="Command" value="Related Topics, menu">
                    <param name="text" value="text:File">
					    <PARAM name="Item1" value="Setup;../accounts/Index_accounts.asp">
				  </OBJECT>
				  </td>
				  <td width="70" class="MenuCell">
                <OBJECT id="menu" classid=clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11 width=27 height=13 type=application/x-oleobject VIEWASTEXT>
					<PARAM NAME="Width" VALUE="1000">
					<PARAM NAME="Height" VALUE="500">
					<PARAM name="Font" value="Verdana, 8,,#000000, ">
					<PARAM NAME="Command" VALUE="Related Topics, menu">
					<PARAM NAME="text" VALUE="text:Goto">
					    <PARAM name="Item1" value="Master;../accounts/Index_accounts.asp">
		   				 <PARAM name="Item2" value="Transaction;../purchase/Index_Purchase.asp">
					    <PARAM name="Item3" value="Reports;../sales/index_sales.asp">
				  </OBJECT>
				 </td>

              </tr>
            </table>
          </td>
          <td width="50%" class="MainMenuCell" align="right">
            <table border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td align="right"><select size="1" name="D1" class="FormElemSmall" >
                            <option>Year 2005 - 2006</option>
                  </select></td>
                <td align="right"><select size="1" name="D2" class="FormElemSmall">
                            <option>Unit I</option>
                    <option>Unit II</option>
                  </select></td>
              </tr>
            </table>
          </td>
        </tr>

      </table>
            <IFRAME NAME="bodyFrame" ID=IFrame2 FRAMEBORDER=0 SCROLLING=YES SRC="body_home.htm" NORESIZE="RESIZE" width="100%" HEIGHT="95%"></IFRAME>
    </td>
  </tr>
</table>
            	<object id=closeWin type="application/x-oleobject" classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11">
<param name="Command" value="Close">
</object>

</form>
</BODY>
</HTML>