(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] : null;
	}

	function byId(id) {
		return document.getElementById(id) || window[id] || null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function checkNumbers(value) {
		return /^[0-9.]+$/.test(String(value || ""));
	}

	function xmlRoot(value) {
		ensureCompat();
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || value;
	}

	function elementChildren(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		var children = node && node.childNodes || [];
		for (var i = 0; i < children.length; i += 1) {
			if (children[i].nodeType === 1 && (!wanted || String(children[i].nodeName).toLowerCase() === wanted)) {
				result.push(children[i]);
			}
		}
		return result;
	}

	function attr(node, name) {
		return node && node.getAttribute ? trim(node.getAttribute(name)) : "";
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args, features, callback);
		}
		return window.open(url, "_blank");
	}

	function dataIsland() {
		ensureCompat();
		return window.Data || document.Data || byId("Data");
	}

	window.GetLot = function (obj) {
		openDialog("newreceiptLotSerPop.asp?sTemp=" + encodeURIComponent(obj && obj.name || ""), dataIsland(), "dialogHeight:320px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No");
	};

	window.GetRate = function (obj) {
		openDialog("newReceiptPackingDetailsPop.asp?sTemp=" + encodeURIComponent(obj && obj.name || ""), dataIsland(), "dialogHeight:420px;dialogWidth:380px;center:Yes;help:No;resizable:No;status:No", function () {
			var root = xmlRoot(dataIsland());
			var headers = elementChildren(root);
			var rateIndex = 1;
			var pages;
			var value;
			var qty;
			var rateField;
			for (var h = 0; h < headers.length; h += 1) {
				pages = elementChildren(headers[h], "STAGE");
				for (var p = 0; p < pages.length; p += 1) {
					value = toNumber(attr(pages[p], "IVALUE"));
					qty = toNumber(attr(pages[p], "IQTY"));
					rateField = field("txtRate" + rateIndex);
					if (rateField && qty) {
						rateField.value = value / qty;
					}
					rateIndex += 1;
				}
			}
		});
	};

	window.DisplayItem = function (obj) {
		var orgName = field("hOrgName");
		openDialog("itmDetailsPop.asp?sTemp=" + encodeURIComponent(obj || ""), orgName ? orgName.value : "", "dialogHeight:400px;dialogWidth:650px;center:Yes;help:No;resizable:No;status:No");
	};

	window.DisableTxt = function (obj) {
		var arrTemp = String(obj && obj.name || "").split("n");
		var tare = field("txtTare" + (arrTemp[1] || ""));
		if (!tare) {
			return;
		}
		tare.value = "";
		tare.disabled = obj && obj.value === "I";
	};

	window.GetLott = function (obj, iQty) {
		var arrTemp = String(obj && obj.name || "").split(":");
		var iRec = arrTemp[1] || "";
		var sClass = arrTemp[2] || "";
		var sItem = arrTemp[3] || "";
		var sOrgID = arrTemp[4] || "";
		var sVar = "btn:" + iRec + ":" + sClass + ":" + sItem + ":" + sOrgID;
		var elements = form() && form().elements || [];
		var iValue = "";
		var item;
		for (var i = 0; i < elements.length; i += 1) {
			item = elements[i];
			if (String(item.type || "").toLowerCase() === "button" && String(item.name || "").indexOf(sVar) > -1) {
				if (elements[i - 2] && elements[i - 2].selectedIndex === 0) {
					alert("Select Quantity in");
					elements[i - 2].focus();
					return;
				}
				if (elements[i - 2] && elements[i - 2].selectedIndex === 1 && elements[i - 1] && elements[i - 1].selectedIndex === 0) {
					alert("Select Tare Weight");
					elements[i - 1].focus();
					return;
				}
				if (elements[i - 1] && elements[i - 1].selectedIndex === 1 && elements[i + 2] && trim(elements[i + 2].value) === "") {
					alert("Enter Tare Weight");
					elements[i + 2].select();
					return;
				}
				if (elements[i - 1] && elements[i - 1].selectedIndex === 1 && elements[i + 2] && trim(elements[i + 2].value) !== "" && !checkNumbers(trim(elements[i + 2].value))) {
					alert("Enter only Numerals");
					elements[i + 2].select();
					return;
				}
				iValue = elements[i + 2] ? elements[i + 2].value : "";
			}
		}
		if (iValue === "") {
			iValue = "N";
		}
		openDialog("newreceiptStorePop.asp?sTemp=" + encodeURIComponent((obj && obj.name || "") + ":" + iValue + ":" + iQty), dataIsland(), "dialogHeight:320px;dialogWidth:400px;center:Yes;help:No;resizable:No;status:No");
	};

	window.Init = function () {
		var spanDate = byId("spanDate");
		var hDate = field("hDate");
		if (spanDate && hDate) {
			spanDate.innerText = hDate.value;
			spanDate.textContent = hDate.value;
		}
	};

	window.CheckSubmit = function () {
		var frm = form();
		var hRecNo = field("hRecNo");
		if (frm) {
			frm.action = "newInternalReceiptAcc.asp?RecNo=" + encodeURIComponent(hRecNo ? hRecNo.value : "");
			frm.submit();
		}
	};
}(window, document));
