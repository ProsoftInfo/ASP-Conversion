(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && (frm.elements[name] || frm[name]) || document.getElementById(name) || null;
	}

	function selectByText(select, value) {
		var target = trim(value).toLowerCase();
		if (!select || !target) {
			return false;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].text).toLowerCase() === target) {
				select.selectedIndex = i;
				return true;
			}
		}
		return false;
	}

	function selectByValue(select, value) {
		var target = trim(value).toLowerCase();
		if (!select || !target) {
			return false;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (trim(select.options[i].value).toLowerCase() === target) {
				select.selectedIndex = i;
				return true;
			}
		}
		return false;
	}

	function setDate(name, value) {
		var picker = field(name);
		if (!picker) {
			return;
		}
		if (typeof picker.SetDate === "function") {
			picker.SetDate(value);
		} else if (typeof picker.setDate === "function") {
			picker.setDate(value);
		} else {
			picker.value = value || "";
		}
	}

	function getDate(name) {
		var picker = field(name);
		if (!picker) {
			return "";
		}
		if (typeof picker.GetDate === "function") {
			return picker.GetDate();
		}
		if (typeof picker.getDate === "function") {
			return picker.getDate();
		}
		return picker.value || "";
	}

	function init(dateValue, countryText, currencyValue) {
		setDate("OrgCSTRCDate", dateValue);
		selectByText(field("selOrgCountry"), countryText);
		selectByValue(field("selOrgCurrency"), currencyValue);
	}

	function setHiddenDate() {
		var country = field("selOrgCountry");
		var currency = field("selOrgCurrency");
		var hiddenDate = field("txtOrgCSTRCDate");
		var hiddenCountry = field("hcountryValue");
		var hiddenCurrency = field("hcurrency");
		if (hiddenDate) {
			hiddenDate.value = getDate("OrgCSTRCDate");
		}
		if (hiddenCountry && country) {
			hiddenCountry.value = country.value;
		}
		if (hiddenCurrency && currency) {
			hiddenCurrency.value = currency.value;
		}
	}

	window.Init = init;
	window.sethiddenDate = setHiddenDate;
}(window, document));
