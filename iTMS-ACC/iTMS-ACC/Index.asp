<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 FINAL//EN">
<HTML>
<HEAD>
<TITLE></TITLE>
<script>
	function openLogin() {
		var option = "width=800,height=525,toolbar=no,titlebar=no,location=no,directories=no,status=yes,personalbar=no,menubar=no,scrollbars=No,resizable=yes";
		var win = window.open("login.asp", "Welcome", option); //./asp/main.asp
		if (win && typeof win.moveTo == "function") {
			try {
				win.moveTo(0,0);
			}
			catch (ignore) {
			}
		}
	}
</script>

</HEAD>
<BODY BGCOLOR="#FFFFFF" LINK="#0000FF" VLINK="#800080" TEXT="#000000" TOPMARGIN=0 LEFTMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 onLoad="openLogin()">
</BODY>
</HTML>
