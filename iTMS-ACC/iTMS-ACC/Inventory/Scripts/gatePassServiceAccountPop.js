(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function form() {
		return document.forms.formname || document.forms[0];
	}

	function field(name) {
		var frm = form();
		if (!frm) {
			return null;
		}
		return frm.elements[name] || frm.elements[name.toLowerCase()] || frm.elements[name.toUpperCase()] || null;
	}

	function numeric(value) {
		var number = Number(trim(value));
		return isNaN(number) ? 0 : number;
	}

	function getXmlText(xmlObject) {
		var target = xmlObject && (xmlObject.XMLDocument || xmlObject._doc || xmlObject);
		if (!target) {
			return "";
		}
		if (typeof target.xml === "string") {
			return target.xml;
		}
		return new XMLSerializer().serializeToString(target);
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function childElements(node) {
		return Array.prototype.filter.call(node && node.childNodes || [], function (child) {
			return child.nodeType === 1;
		});
	}

	function getCheckedRadioValue(name) {
		var control = field(name);
		var i;
		if (!control) {
			return "";
		}
		if (typeof control.length === "number" && !control.tagName) {
			for (i = 0; i < control.length; i += 1) {
				if (control[i].checked) {
					return control[i].value;
				}
			}
			return "";
		}
		return control.checked ? control.value : "";
	}

	function setRadioDisabled(name, disabled) {
		var control = field(name);
		var i;
		if (!control) {
			return;
		}
		if (typeof control.length === "number" && !control.tagName) {
			for (i = 0; i < control.length; i += 1) {
				control[i].disabled = !!disabled;
			}
			return;
		}
		control.disabled = !!disabled;
	}

	function publishReturnValue() {
		var root = window.InvData && window.InvData.documentElement;
		if (!root) {
			return;
		}
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	function closeWithReturnValue() {
		publishReturnValue();
		window.close();
	}

	function isValidReceivedDate(value, today) {
		var isValid = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.test(trim(value));
		if (!isValid) {
			return false;
		}
		if (typeof window.checkValidDate === "function") {
			return window.checkValidDate(value, today, 0);
		}
		return true;
	}

	window.CheckSubmit = function () {
		var itemCount = parseInt(field("hItemCtr") && field("hItemCtr").value, 10) || 0;
		var root = window.OutData && window.OutData.documentElement;
		var invoiceRoot = window.InvData && window.InvData.documentElement;
		var today = field("hToDaysDate") && field("hToDaysDate").value;
		var dataCount = 0;
		var i;
		var entryNo;
		var receivedOn;
		var itemNode;
		var invoiceNode;
		var xhr;

		if (!itemCount || !root) {
			return;
		}

		clearChildren(root);
		root.setAttribute("GPNO", trim(field("hGPNo") && field("hGPNo").value));

		for (i = 1; i <= itemCount; i += 1) {
			entryNo = field("hEntryNo" + i);
			receivedOn = field("RecdON" + i);
			if (!receivedOn || !trim(receivedOn.value)) {
				continue;
			}
			if (!isValidReceivedDate(receivedOn.value, today)) {
				alert("Invalid Date");
				receivedOn.focus();
				return;
			}
			dataCount += 1;
			itemNode = window.OutData.createElement("Item");
			itemNode.setAttribute("EntryNo", entryNo ? entryNo.value : "");
			itemNode.setAttribute("ReceivedOn", trim(receivedOn.value));
			root.appendChild(itemNode);
		}

		if (invoiceRoot) {
			clearChildren(invoiceRoot);
			if (trim(field("hMaterialRcvd") && field("hMaterialRcvd").value) === "N" && field("chkInvoice") && field("chkInvoice").checked) {
				invoiceNode = window.InvData.createElement("WithInv");
				invoiceNode.setAttribute("Rate", getCheckedRadioValue("radRate"));
				invoiceRoot.appendChild(invoiceNode);
			}
		}

		if (itemCount !== dataCount) {
			closeWithReturnValue();
			return;
		}

		xhr = new XMLHttpRequest();
		xhr.open("POST", "GatePassServiceAccountInsert.asp", true);
		xhr.setRequestHeader("Content-Type", "text/xml; charset=UTF-8");
		xhr.onreadystatechange = function () {
			if (xhr.readyState !== 4) {
				return;
			}
			if (xhr.responseText === "") {
				alert("Gate Pass for Service has been accounted.");
				closeWithReturnValue();
			} else {
				alert(xhr.responseText);
			}
		};
		xhr.send(getXmlText(window.OutData));
	};

	window.chkInvoice_onClick = function () {
		setRadioDisabled("radRate", !(field("chkInvoice") && field("chkInvoice").checked));
	};

	window.CalculateTotal = function () {
		var itemCount = parseInt(field("hItemCtr") && field("hItemCtr").value, 10) || 0;
		var total = 0;
		var i;
		var rate;
		var quantity;
		var value;
		var lineTotal;

		for (i = 1; i <= itemCount; i += 1) {
			rate = field("txtRate" + i);
			if (!rate) {
				continue;
			}
			if (!trim(rate.value)) {
				alert("Enter the Rate");
				rate.value = "";
				rate.focus();
				return;
			}
			if (trim(rate.value) === "0") {
				alert("Item Rate must be greater that zero");
				rate.value = "";
				rate.focus();
				return;
			}
			if (isNaN(Number(rate.value))) {
				alert("Enter Numeric Value");
				rate.value = "";
				rate.focus();
				return;
			}
		}

		for (i = 1; i <= itemCount; i += 1) {
			quantity = field("txtQuantity" + i);
			rate = field("txtRate" + i);
			value = field("txtValue" + i);
			if (!quantity || !rate || !value) {
				continue;
			}
			lineTotal = numeric(quantity.value) * numeric(rate.value);
			total += lineTotal;
			value.value = lineTotal;
		}
		if (field("txtTotalValue")) {
			field("txtTotalValue").value = total;
		}
	};

	window.addEventListener("beforeunload", publishReturnValue);
})(window, document);
