(function (window) {
	"use strict";

	window.ITMSReceiptStore.install({
		lotPopup: "receiptLotSerPop.asp",
		lotDetailsUrl: "XMLGetLotDetails.asp",
		formatPopupQuantity: true
	});
}(window));
