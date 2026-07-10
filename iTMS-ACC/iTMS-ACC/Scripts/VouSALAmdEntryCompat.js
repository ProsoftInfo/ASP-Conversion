(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value, decimals) {
		return toNumber(value).toFixed(decimals == null ? 2 : decimals);
	}

	function form() {
		return document.formname || document.forms.formname || document.forms[0] || {};
	}

	function field(name) {
		var frm = form();
		var target = String(name).toLowerCase();
		var index;
		if (!frm || !frm.elements) {
			return null;
		}
		if (frm.elements[name]) {
			return frm.elements[name];
		}
		for (index = 0; index < frm.elements.length; index += 1) {
			if (String(frm.elements[index].name || "").toLowerCase() === target) {
				return frm.elements[index];
			}
		}
		return document.getElementsByName(name)[0] || document.getElementById(name) || null;
	}

	function valueOf(name, fallback) {
		var item = field(name);
		return item ? item.value : fallback || "";
	}

	function setValue(name, value) {
		var item = field(name);
		if (item) {
			item.value = value == null ? "" : String(value);
		}
	}

	function byId(id) {
		return document.getElementById(id) || document.getElementsByName(id)[0] || window[id] || null;
	}

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.innerHTML = value == null ? "" : String(value);
		}
	}

	function selectedText(select) {
		return select && select.selectedIndex >= 0 ? select.options[select.selectedIndex].text : "";
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		}
	}

	function xmlObject(nameOrObject) {
		if (typeof nameOrObject !== "string") {
			return nameOrObject;
		}
		ensureCompat();
		return window[nameOrObject] || document[nameOrObject] || byId(nameOrObject) || null;
	}

	function xmlDocument(nameOrObject) {
		var object = xmlObject(nameOrObject);
		return object && object.XMLDocument || object && object._doc || object && object.nodeType === 9 && object || null;
	}

	function xmlRoot(nameOrObject) {
		var object = xmlObject(nameOrObject);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || object && object.nodeType === 1 && object || null;
	}

	function serializeXml(nameOrObject) {
		var doc = xmlDocument(nameOrObject);
		var root = xmlRoot(nameOrObject);
		if (doc) {
			return new XMLSerializer().serializeToString(doc);
		}
		return root ? new XMLSerializer().serializeToString(root) : "";
	}

	function childElements(node, nodeName) {
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
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

	function setAttr(node, name, value) {
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var result = [];
		var i;
		if (!context) {
			return result;
		}
		if (typeof context.selectNodes === "function") {
			return Array.prototype.slice.call(context.selectNodes(expression));
		}
		doc = context.nodeType === 9 ? context : context.ownerDocument;
		if (!doc || !doc.evaluate) {
			return result;
		}
		found = doc.evaluate(expression, context, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
		for (i = 0; i < found.snapshotLength; i += 1) {
			result.push(found.snapshotItem(i));
		}
		return result;
	}

	function syncPost(url, body) {
		var xhr = new XMLHttpRequest();
		xhr.open("POST", url, false);
		try {
			xhr.setRequestHeader("Content-Type", "text/xml");
		} catch (ignore) {}
		xhr.send(body);
		return xhr;
	}

	function clearTable(tableName, startIndex, keepCount) {
		var table = byId(tableName);
		var start = Number(startIndex) || 0;
		var keep = Number(keepCount) || 0;
		if (!table || !table.rows) {
			return null;
		}
		while (table.rows.length > start + keep) {
			table.deleteRow(start);
		}
		return table;
	}

	function insertCell(row, html, className, align, valign, colspan) {
		if (typeof window.InsertCell === "function") {
			return window.InsertCell(row, 1, "", html, className || "ExcelDisplayCell", align || "left", valign || "top", 0, 0, colspan || 0, 0, "");
		}
		var cell = row.insertCell();
		cell.innerHTML = html == null ? "" : String(html);
		cell.className = className || "ExcelDisplayCell";
		if (align) {
			cell.align = align;
		}
		if (valign) {
			cell.vAlign = valign;
		}
		if (colspan) {
			cell.colSpan = colspan;
		}
		return cell;
	}

	function escapeHtml(value) {
		return String(value == null ? "" : value)
			.replace(/&/g, "&amp;")
			.replace(/</g, "&lt;")
			.replace(/>/g, "&gt;")
			.replace(/"/g, "&quot;");
	}

	function setEntryButtons(editing) {
		var add = field("btnAdd");
		var next = field("btnNext");
		var update = field("btnUpdate");
		var del = field("btnDel");
		if (add) {
			add.disabled = !!editing;
		}
		if (next) {
			next.disabled = !!editing;
		}
		if (update) {
			update.disabled = !editing;
		}
		if (del) {
			del.disabled = !editing;
		}
	}

	function updateDetailsTotals(detailsRoot) {
		var total = 0;
		if (!detailsRoot) {
			return;
		}
		childElements(detailsRoot, "Entry").forEach(function (entry) {
			total += toNumber(attr(entry, "Amount"));
		});
		setAttr(detailsRoot, "BasicValue", formatNumber(total, 2));
		setAttr(detailsRoot, "ActualValue", formatNumber(total, 2));
	}

	function syncInvoiceHeader(root) {
		var saleInvoice = selectNodes(root, "//SaleInvoice")[0];
		if (saleInvoice) {
			setAttr(saleInvoice, "InvDate", valueOf("hInvDate"));
			setAttr(saleInvoice, "InvNo", valueOf("hInvNo"));
			setAttr(saleInvoice, "RefNo", valueOf("hRefNo"));
		}
	}

	function fromOtherApplication() {
		return trim(valueOf("hFrmOthApp", "0")) !== "0";
	}

	function displayAmount(entry) {
		var ratePer = toNumber(attr(entry, "RatePer")) || 1;
		if (fromOtherApplication()) {
			return toNumber(attr(entry, "Amount"));
		}
		return toNumber(attr(entry, "Qty")) * (toNumber(attr(entry, "Rate")) / ratePer) -
			toNumber(attr(entry, "DisAmount")) + toNumber(attr(entry, "RndOff"));
	}

	window.validate = function () {
		return true;
	};

	window.ValidateAmount = function (amount, name, from, to) {
		var text = trim(amount);
		var number = toNumber(amount);
		if (text === "") {
			alert(name + " Cannot be blank");
			return false;
		}
		if (isNaN(Number(String(amount).replace(/,/g, "")))) {
			alert("Enter Numeric values for " + name);
			return false;
		}
		if (number < Number(from) || number > Number(to)) {
			alert(name + " should be >" + from + " and < " + to);
			return false;
		}
		return true;
	};

	window.checkFileds = function () {
		if (!window.ValidateAmount(valueOf("txtQty"), "Quantity", 1, 9999999.999)) {
			if (field("txtQty")) {
				field("txtQty").select();
			}
			return false;
		}
		if (!window.ValidateAmount(valueOf("txtRate"), "Rate", 1, 9999999999.99)) {
			if (field("txtRate")) {
				field("txtRate").select();
			}
			return false;
		}
		if (!window.ValidateAmount(valueOf("txtDisAmount"), "Discount", 0, 9999999999.99)) {
			if (field("txtDisAmount")) {
				field("txtDisAmount").select();
			}
			return false;
		}
		return true;
	};

	window.calculateField = function (flag) {
		var qty = toNumber(valueOf("txtQty"));
		var rate = toNumber(valueOf("txtRate"));
		var ratePer = toNumber(valueOf("txtRatePer"));
		var value;
		var discountPercent;
		var discountAmount;
		if (!window.ValidateAmount(valueOf("txtQty"), "Quantity", 0, 9999999.999)) {
			if (field("txtQty")) {
				field("txtQty").select();
			}
			return false;
		}
		if (!window.ValidateAmount(valueOf("txtRate"), "Rate", 0, 9999999999.99)) {
			if (field("txtRate")) {
				field("txtRate").select();
			}
			return false;
		}
		if (!window.ValidateAmount(valueOf("txtRatePer"), "Rate", 0.000001, 9999999999.99)) {
			if (field("txtRatePer")) {
				field("txtRatePer").select();
			}
			return false;
		}
		value = rate / ratePer * qty;
		setValue("txtValue", formatNumber(value, 2));
		if (Number(flag) === 2) {
			discountPercent = toNumber(valueOf("txtDisPercentage"));
			if (discountPercent > 100) {
				alert("DisCount Percentage Should be less than 100");
				if (field("txtDisPercentage")) {
					field("txtDisPercentage").select();
				}
				return false;
			}
			discountAmount = value * (discountPercent / 100);
			setValue("txtDisAmount", formatNumber(discountAmount, 2));
		} else if (Number(flag) === 3) {
			discountAmount = toNumber(valueOf("txtDisAmount"));
			if (discountAmount > value) {
				alert("DisCount Value Should be less than actual Value");
				if (field("txtDisAmount")) {
					field("txtDisAmount").select();
				}
				return false;
			}
			setValue("txtDisPercentage", value ? formatNumber(discountAmount / value * 100, 2) : "0.00");
		} else {
			discountPercent = toNumber(valueOf("txtDisPercentage"));
			discountAmount = discountPercent > 0 ? value * (discountPercent / 100) : toNumber(valueOf("txtDisAmount"));
			setValue("txtDisAmount", formatNumber(discountAmount, 2));
		}
		setValue("txtAmount", formatNumber(value - toNumber(valueOf("txtDisAmount")), 2));
		if (typeof window.popAddAmount1 === "function") {
			window.popAddAmount1();
		}
		return true;
	};

	window.DisplayVoucher = function () {
		var details = selectNodes(xmlRoot("VoucherData"), "//Details")[0] || xmlRoot("VoucherData");
		var table = clearTable("tblVoucher", 1, 1);
		var total = 0;
		var display = byId("DisVoucher");
		var row;
		if (display && display.style) {
			display.style.height = "200px";
			display.style.visibility = "visible";
		}
		if (!table) {
			return;
		}
		window.VouRoot = details;
		childElements(details, "Entry").forEach(function (entry, index) {
			var amount = displayAmount(entry);
			var accountNode = childElements(entry, "AccHead")[0];
			var description = attr(entry, "PayTo");
			var entryNo = index + 1;
			setAttr(entry, "No", entryNo);
			if (accountNode && trim(attr(accountNode, "Name")) !== "") {
				description = attr(accountNode, "Name") + " - " + description;
			}
			total += amount;
			row = table.insertRow(table.rows.length);
			insertCell(row, entryNo, "ExcelSerial", "Center", "top");
			insertCell(row, fromOtherApplication() ? "&nbsp;" : '<a href="javascript:EditEntry(\'' + entryNo + '\')" class="ExcelDisplayCell"><b>Edit</b></a>', "ExcelDisplayCell", "Center", "top");
			insertCell(row, escapeHtml(description), "ExcelDisplayCell", "left", "top");
			insertCell(row, escapeHtml(formatNumber(attr(entry, "Qty"), 3) + " " + attr(entry, "UOMValue")), "ExcelDisplayCell", "left", "top");
			insertCell(row, formatNumber(attr(entry, "Rate"), 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, formatNumber(attr(entry, "ActValue"), 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, formatNumber(attr(entry, "DisAmount"), 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, formatNumber(amount, 2), "ExcelDisplayCell", "right", "top");
		});
		row = table.insertRow(table.rows.length);
		insertCell(row, "<b>Total</b>", "ExcelSerial", "right", "top", 7);
		insertCell(row, formatNumber(total, 2), "ExcelDisplayCell", "right", "top");
	};

	window.SaveXML = function () {
		var root = xmlRoot("VoucherData");
		var details = selectNodes(root, "//Details")[0] || root;
		var xhr;
		if (!root) {
			alert("Voucher XML is not loaded.");
			return false;
		}
		syncInvoiceHeader(root);
		updateDetailsTotals(details);
		xhr = syncPost("XMLSave.asp?Name=Voucher AMD&Mod=SAL", serializeXml("VoucherData"));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return false;
		}
		if (field("btnNext")) {
			field("btnNext").disabled = true;
		}
		form().submit();
		return true;
	};

	window.ChDisp = function (select) {
		if (!select) {
			return;
		}
		if (String(select.name || "").toUpperCase() === "SELUOM") {
			setText("spUOM", selectedText(select));
		} else {
			setText("spPack", selectedText(select));
		}
	};

	window.CancelAction = function (page) {
		form().action = page;
		form().submit();
	};

	window.InitVouSALAmdEntry = function () {
		var details;
		window.SalesVoucherEntryMode = "edit";
		ensureCompat();
		details = selectNodes(xmlRoot("VoucherData"), "//Details")[0] || xmlRoot("VoucherData");
		window.VouRoot = details;
		window.EntryRoot = xmlRoot("EntryData");
		window.iEntryNo = childElements(details, "Entry").length + 1;
		setEntryButtons(false);
		window.DisplayVoucher();
		if (field("selUOM")) {
			setText("spUOM", selectedText(field("selUOM")));
		}
		if (field("selPack")) {
			setText("spPack", selectedText(field("selPack")));
		}
	};
}(window, document));
