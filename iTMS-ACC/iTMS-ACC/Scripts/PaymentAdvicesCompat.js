(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function openDialog(url, features) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", features, function () {});
			return;
		}
		window.open(url, "_blank", "width=750,height=600,resizable=yes,scrollbars=yes,status=no");
	}

	window.ShowVouch = function (transNo) {
		openDialog("PaymentAdviceDis.asp?Value=" + encodeURIComponent(transNo), "dialogHeight:600px;dialogWidth:750px;center:Yes;help:No;resizable:No;status:No");
		return false;
	};
})(window, document);
