(function () {
	if (!window.ITMSVoucherEntry) {
		throw new Error("VoucherEntryCore.js must be loaded before JournalVoucher.js");
	}
	window.ITMSVoucherEntry.install({
		bookCode: "08",
		moduleCode: "GJ",
		bank: false,
		journal: true,
		payRecPage: "/Accounts/Transaction/PayRecSelection.asp"
	});
}());
