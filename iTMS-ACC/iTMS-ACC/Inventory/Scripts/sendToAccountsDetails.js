(function (window, document) {
	"use strict";

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function openDialog(url, features) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", features || "", function () {});
			return null;
		}
		return window.open(url, "_blank", "width=650,height=370,resizable=no,status=no,scrollbars=yes");
	}

	window.CheckSubmit = function () {
		var frm = form();
		if (frm) {
			frm.action = "sendToAccountsInsert.asp";
			frm.submit();
		}
		return false;
	};

	window.CheckSetup = function () {
		openDialog("SetupInvBooks.asp", "dialogWidth:500px;dialogHeight:250px;Status:No;");
		return false;
	};

	window.OpenSendToAccountsPopup = function (accountHead, unit, closingDate) {
		var url = "SendToAccountsPopup.asp?AccountHead=" + encodeURIComponent(accountHead || "") +
			"&Unit=" + encodeURIComponent(unit || "") +
			"&Date=" + encodeURIComponent(closingDate || "");
		openDialog(url, "dialogHeight:370px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No");
		return false;
	};
}(window, document));
