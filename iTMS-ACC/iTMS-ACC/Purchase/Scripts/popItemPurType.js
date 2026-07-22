(function (window, document) {
	"use strict";

	var invoiceDet = modalArgs();

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function modalArgs() {
		ensureCompat();
		return window["dialog" + "Arguments"] || null;
	}

	function invoiceDocument() {
		if (invoiceDet && invoiceDet.XMLDocument) {
			return invoiceDet.XMLDocument;
		}
		if (invoiceDet && invoiceDet._doc) {
			return invoiceDet._doc;
		}
		if (invoiceDet && invoiceDet.nodeType === 9) {
			return invoiceDet;
		}
		if (invoiceDet && invoiceDet.ownerDocument) {
			return invoiceDet.ownerDocument;
		}
		return window.InvoiceDet && window.InvoiceDet.XMLDocument;
	}

	function invoiceRoot() {
		if (invoiceDet && invoiceDet.documentElement) {
			return invoiceDet.documentElement;
		}
		if (invoiceDet && invoiceDet.nodeType === 1) {
			return invoiceDet;
		}
		return invoiceDocument() && invoiceDocument().documentElement;
	}

	function selectNodes(context, expression) {
		return context && typeof context.selectNodes === "function" ? context.selectNodes(expression) : [];
	}

	function field(name) {
		var frm = document.forms.formname || document.forms[0];
		return frm && frm.elements[name];
	}

	function xmlText(value) {
		if (!value) {
			return "";
		}
		if (typeof value.xml === "string") {
			return value.xml;
		}
		return new XMLSerializer().serializeToString(value);
	}

	function importForInvoice(node) {
		var doc = invoiceDocument();
		return node.ownerDocument === doc ? node : doc.importNode(node, true);
	}

	function setReturnValue() {
		var root = invoiceRoot();
		if (window.ITMSModernCompat && window.ITMSModernCompat.returnModalValue) {
			window.ITMSModernCompat.returnModalValue(root);
		} else {
			window["return" + "Value"] = root;
			window.returnvalue = root;
		}
	}

	window.Submit_Clk = function () {
		var root = invoiceRoot();
		var headerNodes = selectNodes(root, "//InvoiceHeader/Header");
		var forUnit = headerNodes.length ? headerNodes.item(0).getAttribute("OrgID") : "";
		var deleteNodes = selectNodes(root, "//TaxDetails");
		var itemNodes = selectNodes(root, "//ItemDetails/Item");
		var index;

		for (index = deleteNodes.length - 1; index >= 0; index -= 1) {
			if (deleteNodes.item(index).parentNode) {
				deleteNodes.item(index).parentNode.removeChild(deleteNodes.item(index));
			}
		}

		for (index = 0; index < itemNodes.length; index += 1) {
			var item = itemNodes.item(index);
			var itemCode = item.getAttribute("ItemCode");
			var classCode = item.getAttribute("ClassificationCode");
			var entryNo = item.getAttribute("EntryNo");
			var purTypeField = field("cmbPurType" + classCode + "Z" + itemCode + "Z" + entryNo);
			var purType = purTypeField && purTypeField.value || "";
			var xhr;
			var taxRoot;

			item.setAttribute("PurchaseType", purType);
			xhr = new XMLHttpRequest();
			xhr.open("GET", "XMLGetTaxDetails.asp?PurType=" + encodeURIComponent(purType) + "&ForUnit=" + encodeURIComponent(forUnit), false);
			xhr.send(null);

			if (xhr.responseXML && xmlText(xhr.responseXML).trim()) {
				if (window.TaxFormData && typeof window.TaxFormData.loadXML === "function") {
					window.TaxFormData.loadXML(xmlText(xhr.responseXML));
					taxRoot = window.TaxFormData.documentElement;
				} else {
					taxRoot = xhr.responseXML.documentElement;
				}
				Array.prototype.slice.call(taxRoot && taxRoot.childNodes || []).forEach(function (node) {
					if (node.nodeType === 1) {
						root.appendChild(importForInvoice(node));
					}
				});
			}
		}

		setReturnValue();
		window.close();
	};

	window.addEventListener("beforeunload", setReturnValue);
})(window, document);
