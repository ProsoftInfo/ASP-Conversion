(function () {
	if (!window.ITMSVoucherEntry) {
		throw new Error("VoucherEntryCore.js must be loaded before CashVoucher.js");
	}
	window.ITMSVoucherEntry.install({
		bookCode: "01",
		moduleCode: "CA",
		bank: false,
		journal: false,
		payRecPage: "PayRecSelectionWithAllAdj.asp"
	});
}());
