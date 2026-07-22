(function (window, document) {
	"use strict";

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function form() {
		return document.forms.formname || document.forms[0] || null;
	}

	function field(name) {
		var frm = form();
		return frm && frm.elements ? frm.elements[name] || null : null;
	}

	function fieldAny() {
		for (var i = 0; i < arguments.length; i += 1) {
			if (field(arguments[i])) {
				return field(arguments[i]);
			}
		}
		return null;
	}

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function xmlIsland(name) {
		ensureCompat();
		return window[name] || document[name] || document.getElementById(name);
	}

	function xmlDocument(value) {
		ensureCompat();
		if (!value) {
			return null;
		}
		if (value.XMLDocument) {
			return value.XMLDocument;
		}
		if (value._doc) {
			return value._doc;
		}
		if (value.nodeType === 9) {
			return value;
		}
		return value.ownerDocument || null;
	}

	function xmlRoot(value) {
		if (!value) {
			return null;
		}
		return value.documentElement || value.XMLDocument && value.XMLDocument.documentElement || value._doc && value._doc.documentElement || (value.nodeType === 1 ? value : null);
	}

	function serializeXml(value) {
		var doc = xmlDocument(value);
		var root = xmlRoot(value);
		if (value && typeof value.xml === "string") {
			return value.xml;
		}
		if (doc && typeof doc.xml === "string") {
			return doc.xml;
		}
		return new XMLSerializer().serializeToString(doc || root);
	}

	function createElement(doc, name) {
		return (doc || document).createElement(name);
	}

	function setAttr(node, name, value) {
		if (node) {
			node.setAttribute(name, String(value == null ? "" : value));
		}
	}

	function removeChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
	}

	function syncPost(url, xmlText) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		xhr.setRequestHeader("Content-Type", "text/xml");
		xhr.send(xmlText || "");
		return xhr;
	}

	function submitForm(action) {
		if (action) {
			form().action = action;
		}
		form().submit();
	}

	function itemCheckbox(index) {
		return fieldAny("ChkBox" + index, "Chkbox" + index);
	}

	window.Search = function () {
		var values = [];
		[["ChkPur", "P"], ["ChkSales", "S"], ["ChkInv", "I"], ["ChkManu", "M"]].forEach(function (entry) {
			var checkbox = field(entry[0]);
			if (checkbox && checkbox.checked) {
				values.push(checkbox.value || entry[1]);
			}
		});
		if (field("hEligibleFor")) {
			field("hEligibleFor").value = values.join(",");
		}
		submitForm();
		return false;
	};

	window.Sort = function (fieldNo, orderByField, order) {
		var target = field("hField" + trim(fieldNo));
		if (target) {
			target.value = trim(orderByField) + ":" + trim(order);
		}
		if (field("hFieldSelected")) {
			field("hFieldSelected").value = fieldNo;
		}
		submitForm();
		return false;
	};

	window.CalculateTotalStock = function (itemCtr, rowCtr, locBinCtr) {
		var issueQty = field("hIssQtyZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr);
		var openRate = field("hOpRateZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr);
		var stockQty = field("txtStkQtyZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr);
		var stockValue = field("txtStkValueZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr);
		var totalQty = field("txtTotQtyZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr);
		if (stockValue && trim(openRate && openRate.value) !== "0") {
			stockValue.value = toNumber(stockQty && stockQty.value) * toNumber(openRate && openRate.value);
		}
		if (totalQty) {
			totalQty.value = toNumber(issueQty && issueQty.value) + toNumber(stockQty && stockQty.value);
		}
		return false;
	};

	window.EnableStock = function (itemCtr) {
		var checkbox = itemCheckbox(itemCtr);
		var total = toNumber(field("hTotalRow" + itemCtr) && field("hTotalRow" + itemCtr).value);
		for (var i = 1; i <= total; i += 1) {
			var locBinCtr = field("hLocBinCtrZ" + itemCtr + "Z" + i) ? field("hLocBinCtrZ" + itemCtr + "Z" + i).value : "";
			var qty = field("txtStkQtyZ" + itemCtr + "Z" + i + "Z" + locBinCtr);
			var value = field("txtStkValueZ" + itemCtr + "Z" + i + "Z" + locBinCtr);
			var className = checkbox && checkbox.checked ? "FormElem" : "FormElemRead";
			if (qty) {
				qty.className = className;
			}
			if (value) {
				value.className = className;
			}
		}
		return false;
	};

	window.CheckSubmit = function () {
		var stockData = xmlIsland("StockData");
		var doc = xmlDocument(stockData);
		var root = xmlRoot(stockData);
		var totalItems = toNumber(field("hTotItem") && field("hTotItem").value);
		var selected = 0;
		removeChildren(root);
		for (var itemCtr = 1; itemCtr <= totalItems; itemCtr += 1) {
			var checkbox = itemCheckbox(itemCtr);
			var itemNode;
			var parts;
			var totalRows;
			if (!checkbox || !checkbox.checked) {
				continue;
			}
			selected += 1;
			parts = String(checkbox.value || "").split(":");
			itemNode = createElement(doc, "Item");
			setAttr(itemNode, "ItemCode", parts[0]);
			setAttr(itemNode, "ClassCode", parts[1]);
			setAttr(itemNode, "OrgCode", parts[2]);
			setAttr(itemNode, "CompItemCode", parts[3]);
			setAttr(itemNode, "Desc", parts[4]);
			root.appendChild(itemNode);
			totalRows = toNumber(field("hTotalRow" + itemCtr) && field("hTotalRow" + itemCtr).value);
			for (var rowCtr = 1; rowCtr <= totalRows; rowCtr += 1) {
				var locBinCtr = field("hLocBinCtrZ" + itemCtr + "Z" + rowCtr) ? field("hLocBinCtrZ" + itemCtr + "Z" + rowCtr).value : "";
				var locBinVal = field("hLocBinValZ" + itemCtr + "Z" + rowCtr) ? field("hLocBinValZ" + itemCtr + "Z" + rowCtr).value : "";
				var locBin = String(locBinVal).split(":");
				var locNode = createElement(doc, "Loc");
				setAttr(locNode, "Loc", locBin[0]);
				setAttr(locNode, "Bin", locBin[1]);
				setAttr(locNode, "TotChangeQty", field("txtTotQtyZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr) ? field("txtTotQtyZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr).value : "");
				setAttr(locNode, "Rate", field("hOpRateZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr) ? field("hOpRateZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr).value : "");
				setAttr(locNode, "StkChange", field("txtStkQtyZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr) ? field("txtStkQtyZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr).value : "");
				setAttr(locNode, "StkValue", field("txtStkValueZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr) ? field("txtStkValueZ" + itemCtr + "Z" + rowCtr + "Z" + locBinCtr).value : "");
				itemNode.appendChild(locNode);
			}
		}
		if (selected > 0) {
			syncPost("XMLSave.asp?SessionFlag=true&Name=ItemOpenStockChange_", serializeXml(stockData));
			submitForm("ItemStockOpenInsert.asp");
		}
		return false;
	};
}(window, document));
