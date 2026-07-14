<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<Title>iTMS Welcome</Title>
<LINK REL="STYLESHEET" HREF="assets/styles/Standard.css" TYPE="text/css">
<script>
function getTileButton()
{
	var form = document.forms[0] || null;
	return (form && form.imgTile) || document.getElementsByName("imgTile")[0] || null;
}
function setTileButton(altText, imagePath)
{
	var button = getTileButton();
	if (button) {
		button.alt = altText;
		button.src = imagePath;
	}
}
function runWindowAction(targetWindow, action, firstValue, secondValue)
{
	try {
		if (targetWindow && typeof targetWindow[action] == "function") {
			targetWindow[action](firstValue, secondValue);
		}
	}
	catch (ignore) {
	}
}
function getWindowTop()
{
	if (typeof window.screenTop == "number") {
		return window.screenTop;
	}
	if (typeof window.screenY == "number") {
		return window.screenY;
	}
	return 0;
}
function navigateBodyFrame(selectBox)
{
	var frame;
	if (!selectBox.value) {
		return;
	}
	frame = document.getElementById("IFrame2");
	if (frame && frame.contentWindow) {
		frame.contentWindow.location.href = selectBox.value;
	}
	else if (window.frames && window.frames.bodyFrame) {
		window.frames.bodyFrame.location.href = selectBox.value;
	}
	selectBox.selectedIndex = 0;
}
function winClose()
{
	window.close();
}
function checkWin()
{
	var hdiff = getWindowTop();
	if(hdiff < 0) {
		smallscreen();
	}
	else {
		fullscreen();
	}
}

function fullscreen(){
	var hdiff;
	runWindowAction(top.window, "moveTo", -4, -4);
	hdiff = getWindowTop();
	runWindowAction(top.window, "moveTo", -6, -hdiff - 7);
	runWindowAction(top.window, "resizeTo", screen.width + 13, screen.height + hdiff + 33);
	setTileButton("Restore Down", "assets/images/ExpandButton.gif");
}
function smallscreen(){
	var hdiff;
	runWindowAction(top.window, "moveTo", 0, 0);
	hdiff = getWindowTop();
	runWindowAction(top.window, "resizeTo", screen.width, screen.height);
	setTileButton("Maximize", "assets/images/CollapseButton.gif");
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
                <td><img border="0" src="assets/images/home.gif" width="24" height="24" alt="iTMS"></td>
                <td class="MainTitle" style="padding-left:8px;">iTMS</td>
                <td background="assets/images/clearpixel.gif" width="100%">&nbsp;</td>
                <td><img border="0" src="assets/images/Top.gif" width="23" height="40" alt=""></td>
                <td><img border="0" src="assets/images/minus.gif"  width="16" height="16" onClick="minScreen()" alt="Minimize" style="cursor: pointer"></td>
                <td><img border="0" src="assets/images/ExpandButton.gif"  width="17" height="14" onClick="checkWin()" alt="Restore Down" style="cursor: pointer" name="imgTile"></td>
                <td><img border="0" src="assets/images/CollapseButton.gif"  width="17" height="14" onClick="winClose()" alt="Close" style="cursor: pointer"></td>
                <td><img border="0" src="assets/images/clearpixel.gif" width="4" height="40" alt=""></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="50%" class="MainMenuCell">
            <table border="0" cellpadding="0" cellspacing="0" height="16">
              <tr>
              <td width="40" class="MenuCell">
                <a href="Accounts/Index_accounts.asp" target="bodyFrame" class="MenuCell">File</a>
				  </td>
				  <td width="70" class="MenuCell">
                <select class="FormElemSmall" onchange="navigateBodyFrame(this)">
					<option value="">Goto</option>
					<option value="Accounts/Index_accounts.asp">Master</option>
					<option value="Purchase/Index_Purchase.asp">Transaction</option>
					<option value="Sales/index_sales.asp">Reports</option>
				  </select>
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
            <IFRAME NAME="bodyFrame" ID=IFrame2 FRAMEBORDER=0 SCROLLING=YES SRC="welcome_welcome.asp" NORESIZE="RESIZE" width="100%" HEIGHT="95%"></IFRAME>
    </td>
  </tr>
</table>
</form>
</BODY>
</HTML>
