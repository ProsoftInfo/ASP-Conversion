<%@ Language=VBScript %>
<% Option Explicit %>
<%
Response.Expires = -10
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<title>Upload Progress</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<link rel="STYLESHEET" href="../assets/styles/StandardBody.css" type="text/css">
<style type="text/css">
body {
	margin: 0;
	background: #f7f9fc;
}
.upload-progress {
	padding: 18px;
	text-align: center;
	font-family: Arial, sans-serif;
	color: #1f2937;
}
.upload-progress-bar {
	width: 92%;
	height: 14px;
	margin: 12px auto;
	border: 1px solid #8aa3c2;
	background: #ffffff;
	overflow: hidden;
}
.upload-progress-fill {
	width: 45%;
	height: 14px;
	background: #4f83cc;
}
</style>
<script type="text/javascript">
function checkUploadWindow() {
	if (!window.opener || window.opener.closed) {
		window.close();
		return;
	}
	window.setTimeout(checkUploadWindow, 1000);
}
</script>
</head>
<body onload="checkUploadWindow()">
<div class="upload-progress">
	<div>Please wait while the image is uploaded...</div>
	<div class="upload-progress-bar"><div class="upload-progress-fill"></div></div>
	<div>Do not close the upload window.</div>
</div>
</body>
</html>
