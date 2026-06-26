<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<TITLE></TITLE>
<script language="javascript">
	function option() {
		var option = "width=800,height=525,toolbar=no,titlebar=no,location=no,directories=no,status=yes,personalbar=no,menubar=no,scrollbars=No,resizable=yes";
		closeWin.Click();
		var win = open ( "login.asp", "Welcome", option ); //./asp/main.asp
		win.moveTo(0,0);
	}
</script>

</HEAD>
<BODY BGCOLOR="#FFFFFF" LINK="#0000FF" VLINK="#800080" TEXT="#000000" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 onLoad="option()">
<object id=closeWin type="application/x-oleobject" classid="clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11">
<param name="Command" value="Close">
</object>
</BODY>
</HTML>