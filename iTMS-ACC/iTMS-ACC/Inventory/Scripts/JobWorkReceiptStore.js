(function (window) {
	"use strict";

	window.ITMSReceiptStore.install({
		lotPopup: "JobWorkReceiptLotSerPop.asp",
		lotDetailsUrl: "XMLGetJobWorkLotDetails.asp",
		formatPopupQuantity: false
	});
}(window));
