(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value) {
		return toNumber(value).toFixed(2);
	}

	function round2(value) {
		return Math.round(toNumber(value) * 100) / 100;
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && (frm.elements[name] || frm[name]) || null;
	}

	function fieldAt(name, index) {
		var item = field(name);
		if (item && item.length && !item.tagName) {
			return item[index] || item[0] || null;
		}
		return item || null;
	}

	function setField(name, value, index) {
		var item = fieldAt(name, index || 0);
		if (item) {
			item.value = value;
		}
	}

	function upgradeXml() {
		if (window.ITMSModernCompat) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
	}

	function xmlObject(name) {
		var element;
		upgradeXml();
		element = document.getElementById(name);
		return window[name] || document[name] || element && element._itmsXmlIsland || element || null;
	}

	function xmlDocument(name) {
		var data = xmlObject(name);
		return data && (data.XMLDocument || data._doc || data) || null;
	}

	function xmlRoot(name) {
		var data = xmlObject(name);
		var doc = xmlDocument(name);
		return data && data.documentElement || doc && doc.documentElement || null;
	}

	function selectNodes(context, expression) {
		var found;
		var nodes = [];
		if (!context) {
			return nodes;
		}
		if (typeof context.selectNodes === "function") {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		found = (context.nodeType === 9 ? context : context.ownerDocument).evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (var i = 0; i < found.snapshotLength; i += 1) {
			nodes.push(found.snapshotItem(i));
		}
		return nodes;
	}

	function childElements(node) {
		var result = [];
		for (var i = 0; node && i < node.childNodes.length; i += 1) {
			if (node.childNodes[i].nodeType === 1) {
				result.push(node.childNodes[i]);
			}
		}
		return result;
	}

	function attr(node, name) {
		var item = node && node.attributes && node.attributes.getNamedItem(name);
		return item ? item.value : "";
	}

	function setAttr(node, name, value) {
		if (node) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function attrAt(node, index) {
		var item = node && node.attributes && node.attributes.item(index);
		return item ? item.value : "";
	}

	function setAttrAt(node, index, value) {
		var item = node && node.attributes && node.attributes.item(index);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function request(method, url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open(method, url, false);
		xhr.send(body || null);
		return xhr;
	}

	function updateMoneyAttr(node, name, rate) {
		setAttr(node, name, formatNumber(round2(toNumber(attr(node, name)) * rate)));
	}

	function updateMoneyAttrAt(node, index, rate) {
		setAttrAt(node, index, formatNumber(round2(toNumber(attrAt(node, index)) * rate)));
	}

	function showInvoice(sPassInvNo) {
		var frm = form();
		field("hInvno").value = sPassInvNo;
		frm.action = "../../sales/TRANSACTION/AmndSaltrInvoice.asp";
		frm.submit();
	}

	function checkSubmit() {
		var frm = form();
		var book = field("selBook");
		if (book && book.selectedIndex < 1) {
			alert("Select a Book");
			book.focus();
			return false;
		}
		if (String(field("hExpInv").value) === "Y") {
			request("POST", "XMLSalSave.asp?TransNo=" + field("hTransNo").value, xmlDocument("SaleData"));
		}
		field("hBookName").value = book.options[book.selectedIndex].text;
		field("btnAction").disabled = true;
		frm.submit();
		return true;
	}

	function finalCancel() {
		var frm = form();
		frm.action = "AppOtherAppSelection.asp";
		frm.submit();
	}

	function viewInv() {
		var url = "SalesInvView.asp?TransNo=" + field("hTransNo").value;
		var features = "dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No";
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "A", features, function () {});
		} else {
			window.open(url, "SalesInvView", "width=800,height=600,resizable=yes,scrollbars=yes");
		}
	}

	function dispBook(source) {
		var result;
		var doc;
		var root;
		var book;
		if (!source || source.selectedIndex === 0) {
			return;
		}
		result = request("GET", "XMLGetOrgBook.asp?BkCode=05&orgID=" + source.value);
		if (trim(result.responseText) === "") {
			return;
		}
		doc = xmlDocument("UnitBookData");
		if (doc && typeof doc.loadXML === "function") {
			doc.loadXML(result.responseText);
		}
		root = xmlRoot("UnitBookData");
		book = field("selBook");
		if (!root || !book) {
			return;
		}
		book.length = 1;
		childElements(root).forEach(function (node) {
			book.options[book.options.length] = new Option(attrAt(node, 1), attrAt(node, 0));
		});
	}

	function calNewTax() {
		var rateField = field("txtConRate");
		var rate = toNumber(rateField && rateField.value);
		var root = xmlRoot("SaleData");
		var entries;
		if (!root) {
			return false;
		}
		entries = selectNodes(root, "//Entry");
		entries.forEach(function (node) {
			updateMoneyAttr(node, "Amount", rate);
			updateMoneyAttr(node, "TransBasicamt", rate);
			updateMoneyAttr(node, "DisAmount", rate);
		});
		selectNodes(root, "//Details").forEach(function (node) {
			updateMoneyAttrAt(node, 0, rate);
			updateMoneyAttrAt(node, 1, rate);
			updateMoneyAttrAt(node, 2, rate);
		});
		selectNodes(root, "//TaxDetails").forEach(function (node) {
			updateMoneyAttrAt(node, 0, rate);
			updateMoneyAttrAt(node, 1, rate);
			updateMoneyAttrAt(node, 2, rate);
		});
		selectNodes(root, "//Tax").forEach(function (node) {
			updateMoneyAttrAt(node, 5, rate);
		});
		entries.forEach(function (node, index) {
			setField("txtBasVal", attr(node, "TransBasicamt"), index);
			setField("txtDisVal", attr(node, "DisAmount"), index);
			setField("txtNetVal", attr(node, "Amount"), index);
		});
		selectNodes(root, "//TaxDetails").forEach(function (node) {
			setField("txtInvVal", attrAt(node, 0));
			setField("txtTotNetVal", attrAt(node, 1));
			setField("txtTotBasVal", attrAt(node, 1));
			setField("txtTotDisVal", "0.00");
		});
		selectNodes(root, "//Tax").forEach(function (node) {
			setField("txtTaxVal" + attrAt(node, 0) + "Z" + attrAt(node, 1), attrAt(node, 5));
		});
		return true;
	}

	function install() {
		upgradeXml();
		window.ShowInvoice = showInvoice;
		window.checkSubmit = checkSubmit;
		window.finalCancel = finalCancel;
		window.ViewInv = viewInv;
		window.DispBook = dispBook;
		window.CalNewTax = calNewTax;
	}

	window.ITMSAppOtherSALViewCompat = {
		install: install
	};
}(window, document));
