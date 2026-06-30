(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function controls(name) {
		var frm = form();
		var item = frm && (frm.elements[name] || frm[name]);
		if (!item) {
			return [];
		}
		if (item.length && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		return [item];
	}

	function field(name) {
		return controls(name)[0] || null;
	}

	function selectField(name) {
		var list = controls(name);
		for (var i = 0; i < list.length; i += 1) {
			if (list[i].tagName && list[i].tagName.toLowerCase() === "select") {
				return list[i];
			}
		}
		return list[0] || null;
	}

	function selectedRadioValue(name) {
		var list = controls(name);
		var i;
		for (i = 0; i < list.length; i += 1) {
			if (list[i].checked) {
				return list[i].value;
			}
		}
		return "";
	}

	function setValue(name, value) {
		controls(name).forEach(function (item) {
			item.value = value == null ? "" : String(value);
		});
	}

	function textById(id) {
		var element = document.getElementById(id);
		return trim(element && (element.textContent || element.innerText));
	}

	function setText(id, value) {
		var element = document.getElementById(id);
		if (element) {
			element.textContent = value || "\u00a0";
		}
	}

	function nextYearDate(value) {
		var parts = String(value || "").split("/");
		var year = parseInt(parts[2], 10);
		if (parts.length < 3 || isNaN(year)) {
			return value || "";
		}
		return parts[0] + "/" + parts[1] + "/" + (year + 1);
	}

	function encodeParam(value) {
		return encodeURIComponent(value == null ? "" : String(value));
	}

	function openDialog(url, features, callback) {
		var opener;
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			opener = window.ITMSModernCompat.openModalDialog(url, "", features, callback || function () {});
			return !!opener;
		}
		opener = window.open(url, "_blank", "width=600,height=600,resizable=yes,scrollbars=yes");
		if (!opener) {
			alert("Popup was blocked. Please allow popups for this site and try again.");
			return false;
		}
		if (typeof callback === "function") {
			callback();
		}
		return true;
	}

	function submitForm() {
		var frm = form();
		if (frm) {
			frm.submit();
		}
	}

	function selectedOptionText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function setDates(select) {
		var selected;
		var dates;
		if (!select || select.selectedIndex === 0) {
			setText("idPFinEndDate", "");
			setText("idCFinStartDate", "");
			setText("idCFinEndDate", "");
			return false;
		}
		selected = select.value || "";
		dates = selected.split(":");
		setText("idPFinEndDate", dates[1] || "");
		setText("idCFinStartDate", nextYearDate(dates[0]));
		setText("idCFinEndDate", nextYearDate(dates[1]));
		return true;
	}

	function noSeriesSubmit() {
		var frm = form();
		var selected = selectedRadioValue("radModule");
		var parts;
		var passData;
		function finish() {
			frm.action = "CloseEntry.asp?Frm=NS";
			submitForm();
		}
		if (!selected) {
			alert("Select Module");
			return false;
		}
		parts = selected.split("#");
		passData = [parts[4], parts[3], parts[7], parts[8], parts[5], parts[6], parts[2]].join("||");
		if (!confirm("Do U Want to Proceed?")) {
			return false;
		}
		if (parts[1] === "1") {
			return openDialog("TransferClosingPopUp.ASP?Para=" + encodeParam(passData), "dialogHeight:450px;dialogWidth:300px;center:Yes;status:no", finish);
		}
		finish();
		return true;
	}

	function accSubmit() {
		var selected = selectedRadioValue("radFinPeriod");
		var parts;
		var period;
		var currFrom;
		var currTo;
		var frm = form();
		if (!selected) {
			alert("Select Financial Period");
			return false;
		}
		parts = selected.split("#");
		period = (parts[0] || "").split(":");
		currFrom = nextYearDate(period[0]);
		currTo = nextYearDate(period[1]);
		if (!confirm("Do U Want to Proceed?")) {
			return false;
		}
		setValue("selPFinStartDate", parts[0]);
		frm.action = "AccountsClosingNew.asp?UnitCode=" + encodeParam(parts[1]) + "&UnitName=" + encodeParam(parts[2]) + "&PrevFromDate=" + encodeParam(period[0]) + "&PrevToDate=" + encodeParam(period[1]) + "&CurrFromDate=" + encodeParam(currFrom) + "&CurrToDate=" + encodeParam(currTo);
		setValue("hOrgName", parts[2]);
		setValue("hCFinStartDate", currFrom);
		setValue("hCFinEndDate", currTo);
		submitForm();
		return true;
	}

	function stockSubmit() {
		var selected = selectedRadioValue("radIType");
		var parts;
		var period;
		var currFrom;
		var currTo;
		var frm = form();
		if (!selected) {
			alert("Select Item Type");
			return false;
		}
		parts = selected.split("#");
		period = (parts[0] || "").split(":");
		currFrom = nextYearDate(period[0]);
		currTo = nextYearDate(period[1]);
		if (!confirm("Do U Want to Proceed?")) {
			return false;
		}
		setValue("selPFinStartDate", parts[0]);
		frm.action = "InvClosingStockDetailsEntry.asp?CategoryCodes=" + encodeParam(parts[1]) + "&PrevFromDate=" + encodeParam(period[0]) + "&PrevToDate=" + encodeParam(period[1]) + "&CurrFromDate=" + encodeParam(currFrom) + "&CurrToDate=" + encodeParam(currTo) + "&UnitCode=" + encodeParam(parts[3]) + "&UnitName=" + encodeParam(parts[2]);
		submitForm();
		return true;
	}

	function checkSubmit() {
		var frm = form();
		var unit = field("selUnit");
		var finStart = selectField("selPFinStartDate");
		var target;
		if (unit && unit.selectedIndex === 0) {
			alert("Select Unit");
			unit.focus();
			return false;
		}
		if (finStart && finStart.selectedIndex === 0) {
			alert("Previous Financial Year Start Date");
			finStart.focus();
			return false;
		}
		if (!confirm("Do U Want to Proceed?")) {
			return false;
		}
		if (field("hFor").value === "NS") {
			target = "TransferClosingDetailsEntryNew.asp?Frm=NS";
		} else if (field("hFor").value === "IS") {
			target = "TransferClosingDetailsEntryNew.asp?Frm=IS";
		} else if (field("hFor").value === "PROLOSS") {
			target = "GLPartyUpdate.asp";
		} else {
			target = "AccountsClosingNew.asp";
		}
		frm.action = target;
		setValue("hOrgName", selectedOptionText(unit));
		setValue("hCFinStartDate", textById("idCFinStartDate"));
		setValue("hCFinEndDate", textById("idCFinEndDate"));
		submitForm();
		return true;
	}

	function itemStockSetup() {
		return openDialog("ItemStockCloseSetup.asp", "dialogWidth:600px;dialogHeight:600px;Status:No", submitForm);
	}

	function transfer(sWho, sWhat, iApp) {
		var frm = form();
		var passData;
		setValue("hApplication", iApp);
		passData = [
			field("hUnit").value,
			field("hOrgName").value,
			field("hCFinStartDate").value,
			field("hCFinEndDate").value,
			field("hPFinStartDate").value,
			field("hPFinEndDate").value,
			iApp
		].join("||");
		if (sWhat === "1") {
			openDialog("TransferClosingPopUp.ASP?Para=" + encodeParam(passData), "dialogHeight:450px;dialogWidth:300px;center:Yes;status:no");
		}
		if (sWho === "IN") {
			if (sWhat === "2") {
				frm.action = "InvClosingStockDetailsEntry.asp?ItemTypeID=" + encodeParam(field("hItemTypeId").value);
				submitForm();
			} else if (sWhat === "3") {
				frm.action = "MRPendingDetailsEntry.asp";
				submitForm();
			}
		}
		return false;
	}

	function installCloseEntry() {
		window.SetDates = setDates;
		window.NoSeriesSubmit = noSeriesSubmit;
		window.AccSubmit = accSubmit;
		window.StockSubmit = stockSubmit;
		window.CheckSubmit = checkSubmit;
		window.ItemStockSetup = itemStockSetup;
	}

	function installTransferDetails() {
		window.Transfer = transfer;
	}

	window.ITMSAdminTransferClosingCompat = {
		installCloseEntry: installCloseEntry,
		installTransferDetails: installTransferDetails
	};
}(window, document));
