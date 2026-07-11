(function (window, document) {
	"use strict";

	var config = window.__itmsCndnNewPageConfig || {};

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function fieldValue(name, fallback) {
		var item = field(name);
		return item ? item.value : fallback || "";
	}

	function setFieldValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name) || null;
	}

	function xmlDocument(name) {
		var object = typeof name === "string" ? xmlObject(name) : name;
		return object && object.XMLDocument || object && object._doc || object || null;
	}

	function xmlRoot(name) {
		var object = xmlObject(name);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || null;
	}

	function loadXml(name, xmlText) {
		var object = xmlObject(name);
		var doc;
		if (object && typeof object.loadXML === "function") {
			object.loadXML(xmlText || "<Root/>");
			return object;
		}
		doc = new DOMParser().parseFromString(xmlText || "<Root/>", "text/xml");
		if (object) {
			object._doc = doc;
		}
		return object || doc;
	}

	function serializeXml(nodeOrDoc) {
		if (!nodeOrDoc) {
			return "";
		}
		if (typeof nodeOrDoc.xml === "string") {
			return nodeOrDoc.xml;
		}
		return new XMLSerializer().serializeToString(nodeOrDoc);
	}

	function childElements(node) {
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1;
		});
	}

	function attr(node, nameOrIndex) {
		var item;
		if (!node || !node.attributes) {
			return "";
		}
		if (typeof nameOrIndex === "number") {
			item = node.attributes.item(nameOrIndex);
			return item ? item.nodeValue : "";
		}
		return node.getAttribute(nameOrIndex) || "";
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
			return;
		}
		window.open(url, "_blank", "width=670,height=410,resizable=yes,status=no");
	}

	function dialogFeatures(width, height) {
		return "dialogHeight:" + height + "px;dialogWidth:" + width + "px;center:Yes;help:No;resizable:No;status:No";
	}

	function popupSize(type, fallbackProgram, fallbackHeight, fallbackWidth) {
		var parts;
		if (window.GetWindowSizeForPopup) {
			parts = String(window.GetWindowSizeForPopup(type)).split(":");
			if (parts.length >= 3) {
				return { program: parts[0], height: parts[1], width: parts[2] };
			}
		}
		return { program: fallbackProgram, height: fallbackHeight, width: fallbackWidth };
	}

	function unitValue() {
		return fieldValue(config.unitField || "selUnitId") || fieldValue("hUnitID") || fieldValue("hOrgId");
	}

	function selectedText(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function getDateControl(name) {
		var control = field(name) || byId(name);
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate) {
			return window.ITMSModernCompat.toDisplayDate(control.value);
		}
		return control.value || "";
	}

	function setDateControl(name, value) {
		var control = field(name) || byId(name);
		if (!control || !value) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else {
			control.value = window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate ? window.ITMSModernCompat.toIsoDate(value) : value;
		}
	}

	function selectedCriteria() {
		var count = Number(fieldValue("hCnt", 0));
		var invoiceType = trim(fieldValue("hInvVal"));
		var selected = { count: 0, value: "", details: "" };
		var item;
		var detail;
		for (var i = 1; i <= count; i += 1) {
			item = field("OptCriteria" + i);
			if (item && item.checked) {
				selected.count += 1;
				if (config.mode === "credit" && invoiceType === "SC") {
					selected.value += item.value + ":";
					detail = field("hOptCriteria" + i);
					selected.details += (detail ? detail.value : "") + ",";
				} else {
					selected.value = item.value;
				}
			}
		}
		return selected;
	}

	function bookFromXml(bookCode, callback) {
		var xhr = syncGet("XMLGetOrgBookCountGJ.asp?BkCode=" + encodeURIComponent(bookCode) + "&orgID=" + encodeURIComponent(unitValue()));
		var root;
		var children;
		var xml = xhr.responseXML && xhr.responseXML.documentElement ? serializeXml(xhr.responseXML) : xhr.responseText;
		if (!trim(xml)) {
			callback({ number: fieldValue("hBookNo"), name: "" });
			return;
		}
		loadXml(config.unitBookXml || "UnitBookData", xml);
		root = xmlRoot(config.unitBookXml || "UnitBookData");
		children = childElements(root);
		if (root && attr(root, "Count") === "1" && children.length) {
			callback({ number: attr(children[0], "BookNumber"), name: attr(children[0], "BookName") });
			return;
		}
		openBookPopup(callback);
	}

	function openBookPopup(callback) {
		openDialog("BookPopUp.asp?Unit=" + encodeURIComponent(unitValue()) + "&VouType=" + encodeURIComponent(fieldValue("hVouType")), "", dialogFeatures(370, 200), function (value) {
			var parts = String(value || "0--").split("--");
			callback({ number: parts[0] || "0", name: parts[1] || "" });
		});
	}

	function chooseBook(voucherType, callback) {
		setFieldValue("hVouType", voucherType);
		if (config.mode === "debitNew") {
			if (trim(voucherType) === "GJ") {
				callback({ number: "3", name: "JOURNAL BOOK" });
			} else {
				openBookPopup(callback);
			}
			return;
		}
		if (config.mode === "credit" || config.mode === "creditCreate") {
			if (trim(voucherType) === "GJ") {
				setFieldValue("hChkVal", "Y");
				bookFromXml("08", callback);
			} else if (config.mode === "creditCreate") {
				setFieldValue("hChkVal", "N");
				bookFromXml(config.nonGjBookCode || "07", callback);
			} else {
				setFieldValue("hChkVal", "N");
				openBookPopup(callback);
			}
			return;
		}
		bookFromXml(trim(voucherType) === "GJ" ? "08" : "06", callback);
	}

	function submitVoucher(voucherType, selected, book) {
		var invoiceType = trim(fieldValue("hInvVal"));
		if (!book || trim(book.number) === "0") {
			return;
		}
		setFieldValue("selBook", book.number);
		setFieldValue("selInvoiceNo", selected.value);
		setFieldValue("hBookName", book.name);
		if (field("hVouDetails")) {
			setFieldValue("hVouDetails", config.mode === "credit" && invoiceType === "SC" ? selected.details : selected.value);
		}
		if (config.mode === "credit" || config.mode === "creditCreate") {
			if (invoiceType === "SI" || invoiceType === "SR") {
				if (config.mode === "creditCreate" && invoiceType === "SR") {
					form().action = "VouCNSalesReturnEntry.asp?hCallFrom=" + encodeURIComponent(voucherType);
				} else {
					form().action = "VouCNOtherInvEntry.asp?hCallFrom=" + encodeURIComponent(voucherType);
				}
			} else if (invoiceType === "SC") {
				form().action = "VouCNSalCommission.asp";
			} else if (invoiceType === "PI") {
				form().action = "VouCNPurInvEntry.asp?hCallFrom=" + encodeURIComponent(voucherType);
			} else if (invoiceType === "OT") {
				form().action = "VouCNOthersEntry.asp" + (config.mode === "creditCreate" ? "?CallFrom=" + encodeURIComponent(voucherType) : "");
			} else if (invoiceType === "MI") {
				form().action = "VouCNThrMiscRec.asp?CallFrom=" + encodeURIComponent(voucherType);
			}
		} else {
			if (invoiceType === "SI") {
				form().action = "VouDNSalInvEntry.asp" + (config.mode === "debitCreate" ? "?CallFrom=" + encodeURIComponent(voucherType) : "");
			} else if (invoiceType === "PI") {
				form().action = "VouDNOtherInvEntry.asp" + (config.mode === "debitCreate" ? "?CallFrom=" + encodeURIComponent(voucherType) : "");
			} else if (invoiceType === "OT") {
				form().action = "VouDNOthersEntry.asp" + (config.mode === "debitCreate" ? "?CallFrom=" + encodeURIComponent(voucherType) : "");
			} else if (invoiceType === "MI") {
				form().action = "VouDNThrMiscPay.asp?CallFrom=" + encodeURIComponent(voucherType);
			}
		}
		form().submit();
	}

	function continuePartySelection(url, nextUrlPrefix, features, callback) {
		openDialog(url, "", features, function (value) {
			var parts = String(value || "").split(":");
			if (parts.length <= 1 && trim(value)) {
				continuePartySelection(nextUrlPrefix + value, nextUrlPrefix, features, callback);
				return;
			}
			callback(value || "");
		});
	}

	function applyPartySelection(value, partyTypeValue, select) {
		var parts = String(value || "").split(":");
		var root;
		var nodes;
		var selectedNode;
		if (parts.length <= 2) {
			if (select) {
				select.selectedIndex = 0;
				select.focus();
			}
			return;
		}
		setFieldValue("hParType", parts[4] || "");
		setFieldValue("hSubParType", parts[3] || "");
		setFieldValue("hParCode", parts[1] || "");
		setFieldValue("hParName", parts[0] || "");
		root = xmlRoot("AccHeadData");
		clearChildren(root);
		if (window.GetPartyHeadXml) {
			window.GetPartyHeadXml(parts[1] || "", parts[0] || "", "0:0:0");
		}
		nodes = childElements(xmlRoot("AccHeadData"));
		selectedNode = nodes[0];
		if (selectedNode) {
			setFieldValue("hPartyCode", partyTypeValue + "?" + attr(selectedNode, 0));
			setFieldValue("txtPartyName", attr(selectedNode, 3));
		} else if (select) {
			select.selectedIndex = 0;
		}
	}

	window.FnInit = function () {
		setFieldValue("txtPartyName", fieldValue("hParName"));
		if (field("hFromDate")) {
			setDateControl("ctlFromDate", fieldValue("hFromDate"));
		}
		if (field("hToDate")) {
			setDateControl("ctlToDate", fieldValue("hToDate"));
		}
	};

	window.Validate = function () {
		var parts;
		if (config.validateUnitFromSelect !== false && field("selUnitId") && field("hUnitID")) {
			setFieldValue("hUnitID", fieldValue("selUnitId"));
		}
		parts = String(fieldValue("hPartyCode")).split("?");
		setFieldValue("hParType", parts[0] || "");
		setFieldValue("hSubParType", parts[1] || "");
		setFieldValue("hParName", fieldValue("txtPartyName"));
		setFieldValue("hParCode", parts[3] || "");
		if (field("selPartyType")) {
			setFieldValue("selPartyType", fieldValue("hPartyCode"));
		}
		if (field("hFromDate")) {
			setFieldValue("hFromDate", getDateControl("ctlFromDate"));
		}
		if (field("hToDate")) {
			setFieldValue("hToDate", getDateControl("ctlToDate"));
		}
		form().submit();
	};

	window.Voucher = function (voucherType) {
		var selected = selectedCriteria();
		if (selected.count === 0) {
			if (trim(fieldValue("selPartyType")) === "S" || trim(fieldValue("txtPartyName")) === "") {
				alert("Select 'Party Type' from Expand band to create Other type voucher.");
				return false;
			}
			if (!confirm("This will Create 'Other' Type Voucher for the Party Selected. Do you want to Continue?")) {
				alert("Select Any One Invoice Number");
				return false;
			}
			setFieldValue("hInvVal", "OT");
		}
		chooseBook(voucherType, function (book) {
			submitVoucher(voucherType, selected, book);
		});
		return false;
	};

	window.SetVal = function (obj) {
		var count = Number(fieldValue("hCnt", 0));
		if (config.mode === "credit" && trim(fieldValue("hInvVal")) === "SC") {
			return;
		}
		for (var i = 1; i <= count; i += 1) {
			var item = field("OptCriteria" + i);
			if (item && item !== obj && item.checked && trim(item.value) !== trim(obj.value)) {
				item.checked = false;
			}
		}
	};

	window.popPartyType = function () {
		var select = field("selPartyType");
		var xhr;
		var root;
		if (!select) {
			return;
		}
		while (select.options.length > 1) {
			select.remove(1);
		}
		xhr = syncGet("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(unitValue()));
		if (xhr.responseXML && xhr.responseXML.documentElement) {
			loadXml("OutData", serializeXml(xhr.responseXML));
		} else if (trim(xhr.responseText)) {
			loadXml("OutData", xhr.responseText);
		}
		root = xmlRoot("OutData");
		childElements(root).forEach(function (node) {
			select.add(new Option(trim(node.textContent || node.text || ""), attr(node, 0)));
		});
	};

	window.selParty = function (select) {
		var orgId = unitValue();
		var partyTypeValue = select.value + "?" + selectedText(select);
		var size;
		var url;
		var nextUrl;
		setFieldValue("txtPartyName", "");
		setFieldValue("hPartyCode", "");
		if (!select || select.selectedIndex <= 0) {
			return false;
		}
		if (config.mode === "debitCreate") {
			size = popupSize("12", "PartySelection.asp", "500", "420");
			url = "../../Common/" + size.program + "?orgID=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyTypeValue);
			nextUrl = "../../Common/" + size.program + "?";
			continuePartySelection(url, nextUrl, "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (value) {
				applyPartySelection(value, partyTypeValue, select);
			});
		} else {
			url = "PartySelection.asp?orgId=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyTypeValue);
			continuePartySelection(url, "PartySelection.asp?", dialogFeatures(420, 500), function (value) {
				applyPartySelection(value, partyTypeValue, select);
			});
		}
		return false;
	};

	window.setInvoiceNo = function (obj) {
		var items = document.getElementsByName("voutype");
		for (var i = 0; i < items.length; i += 1) {
			items[i].checked = trim(items[i].value) === trim(obj.value);
			if (items[i].checked) {
				setFieldValue("hInvVal", items[i].value);
			}
		}
		form().submit();
	};

	window.ShowVouch = function (transNo) {
		var page = trim(fieldValue("hInvVal")) !== "PI" ? "SalesVouchView_San.asp" : "PurchaseVouchView_San.asp";
		openDialog(page + "?TransNo=" + encodeURIComponent(transNo), "", "dialogHeight:410px;dialogWidth:670px;center:Yes;help:No;resizable:No;status:No");
		return false;
	};
}(window, document));
