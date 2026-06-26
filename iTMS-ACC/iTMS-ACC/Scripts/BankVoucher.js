(function () {
	if (!window.ITMSVoucherEntry) {
		throw new Error("VoucherEntryCore.js must be loaded before BankVoucher.js");
	}
	window.ITMSVoucherEntry.install({
		bookCode: "02",
		moduleCode: "BA",
		bank: true,
		journal: false,
		payRecPage: "PayRecSelectionWithAllAdj.asp"
	});
}());
