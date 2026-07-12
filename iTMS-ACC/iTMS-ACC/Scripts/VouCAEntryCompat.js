(function (window, document) {
	"use strict";

	var baseAddEntry = window.AddEntry;
	var baseEditEntry = window.EditEntry;
	var baseDelEntry = window.DelEntry;
	var baseInit = window.Init;
	var baseSelAccountHead = window.selAccountHead;
	var baseCheckFileds = window.checkFileds;
	var baseCheckFinDate = window.CheckFinDate;

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
		return document.formname || document.forms[0] || {};
	}

	function field(name) {
		var frm = form();
		return frm.elements && frm.elements[name] || frm[name] || document.getElementsByName(name)[0] || null;
	}

	function valueOf(name, fallback) {
		var item = field(name);
		return item ? item.value : fallback;
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
		var element = byId(id);
		if (element) {
			element.innerHTML = value == null ? "" : String(value);
		}
	}

	function textOf(id) {
		var element = byId(id);
		return element ? trim(element.innerText || element.textContent || element.innerHTML || "") : "";
	}

	function selectedOption(select) {
		return select && select.options && select.selectedIndex >= 0 ? select.options[select.selectedIndex] : null;
	}

	function selectedValue(select) {
		var option = selectedOption(select);
		return option ? option.value : select && select.value || "";
	}

	function selectedText(select) {
		var option = selectedOption(select);
		return option ? option.text : "";
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

	function xmlRoot(name) {
		var object = xmlObject(name);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		return object && object.XMLDocument || object && object._doc || object || null;
	}

	function xmlString(nodeOrDoc) {
		if (!nodeOrDoc) {
			return "";
		}
		if (typeof nodeOrDoc.xml === "string") {
			return nodeOrDoc.xml;
		}
		return new XMLSerializer().serializeToString(nodeOrDoc);
	}

	function attr(node, nameOrIndex) {
		var attribute;
		if (!node || !node.attributes) {
			return "";
		}
		if (typeof nameOrIndex === "number") {
			attribute = node.attributes.item(nameOrIndex);
			return attribute ? attribute.nodeValue : "";
		}
		return node.getAttribute(nameOrIndex) || "";
	}

	function setAttr(node, name, value) {
		if (node) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function childElements(node, name) {
		var result = [];
		var wanted = name && String(name).toLowerCase();
		if (!node || !node.childNodes) {
			return result;
		}
		Array.prototype.forEach.call(node.childNodes, function (child) {
			if (child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted)) {
				result.push(child);
			}
		});
		return result;
	}

	function selectNodes(context, expression) {
		var doc;
		var found;
		var result = [];
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
		for (var i = 0; i < found.snapshotLength; i += 1) {
			result.push(found.snapshotItem(i));
		}
		return result;
	}

	function createHttp() {
		return new XMLHttpRequest();
	}

	function getText(url) {
		var xhr = createHttp();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function currentDateValue() {
		var control = field("ctlDate") || byId("ctlDate");
		if (!control) {
			return "";
		}
		if (typeof control.GetDate === "function") {
			return control.GetDate();
		}
		if (typeof control.getDate === "function") {
			return control.getDate();
		}
		return control.value || "";
	}

	function setDateValue(value) {
		var control = field("ctlDate") || byId("ctlDate");
		if (!control || !value) {
			return;
		}
		if (typeof control.setDate === "function") {
			control.setDate(value);
		} else if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else {
			control.value = value;
		}
	}

	function parseLegacyDate(value) {
		var months = {
			jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
			jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11
		};
		var text = trim(value);
		var match;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
		if (match) {
			return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
		}
		match = text.match(/^(\d{1,2})[\/.-]([A-Za-z]{3})[\/.-](\d{2,4})$/);
		if (match) {
			return new Date(Number(match[3]), months[match[2].toLowerCase()], Number(match[1]));
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
		}
		return null;
	}

	function currentAccountingUnit() {
		var select = field("selAccUnitId");
		var selected = trim(selectedValue(select));
		if (String(valueOf("hOtherUnitFlag", "0")) === "1" && select && select.selectedIndex > 0 && selected && selected !== "A") {
			return {
				id: selected,
				name: selectedText(select)
			};
		}
		return {
			id: valueOf("hOrgId", ""),
			name: valueOf("hOrgName", "")
		};
	}

	function checkedCRDR() {
		var crdr = field("selCRDR");
		if (crdr && crdr.length != null) {
			for (var i = 0; i < crdr.length; i += 1) {
				if (crdr[i].checked) {
					return crdr[i].value;
				}
			}
		}
		return valueOf("hVouCRDR", "");
	}

	function bookAccHead() {
		return valueOf("hBookAccHead", attr(xmlRoot("VoucherData"), "BookAcchead"));
	}

	function formatBalance(value) {
		var number = toNumber(value);
		return formatNumber(Math.abs(number), 2) + (number >= 0 ? " Dr " : " Cr ");
	}

	function clearVoucherTable() {
		var table = byId("tblVoucher");
		if (!table || !table.rows) {
			return null;
		}
		if (typeof window.ClearTable === "function") {
			window.ClearTable("tblVoucher", 1, 1);
		} else {
			while (table.rows.length > 1) {
				table.deleteRow(1);
			}
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

	function entryAccountText(entry) {
		var account = "";
		childElements(entry).forEach(function (child) {
			if (child.nodeName === "AccHead") {
				account = attr(child, "Type") === "P" ? attr(child, "Name") : attr(child, "No") + "-" + attr(child, "Name");
			}
		});
		return account;
	}

	function entryNarration(entry) {
		var narration = "";
		childElements(entry, "Narration").forEach(function (node) {
			narration = node.textContent || node.text || "";
		});
		return narration;
	}

	function entryAdditionalText(entry) {
		var lines = [];
		childElements(entry).forEach(function (header) {
			if (header.nodeName === "CostCenter" || header.nodeName === "Analytical") {
				childElements(header).forEach(function (node) {
					lines.push(escapeHtml(attr(node, "ShortName") || attr(node, "Name") || attr(node, 2)) + "-" +
						escapeHtml(attr(node, "Ratio") || attr(node, 3)) + "%&nbsp;" +
						escapeHtml(attr(node, "Amount") || attr(node, 4)));
				});
			}
			if (header.nodeName === "PayRec") {
				childElements(header).forEach(function (node) {
					lines.push(escapeHtml(attr(node, "InvNo") || attr(node, 1)) + ":" +
						escapeHtml(attr(node, "InvDate") || attr(node, 2)) + "-&nbsp;" +
						escapeHtml(attr(node, "AmtToAdjust") || attr(node, 5)));
				});
			}
		});
		return lines.join("<br>");
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

	function payRecInput(node, index) {
		var code = attr(node, "No");
		var payNo = trim(attr(node, "PayableNo"));
		return field("txtDocAmount" + code + "Z" + payNo + "Z" + (index + 1)) || field("txtDocAmount" + code);
	}

	function selectedOptionIsAccountingUnitPrompt() {
		var select = field("selAccUnitId");
		return String(valueOf("hOtherUnitFlag", "0")) === "1" && select && select.selectedIndex <= 0;
	}

	window.DisplayBalamt = function () {
		var xhr;
		var values;
		var accountHead = bookAccHead();
		if (!accountHead) {
			return;
		}
		xhr = getText("GetDayOpenByDate.asp?sValue=" + encodeURIComponent(valueOf("hOrgId", "") + ":" + accountHead + ":" + currentDateValue()));
		values = String(xhr.responseText || "").split("*");
		if (values.length === 2) {
			setText("spBookBal", formatBalance(values[0]));
			setText("spCurrBal", formatBalance(values[1]));
		}
	};

	window.popAccHead = function () {
		var select = field("selAccHead");
		var unit = currentAccountingUnit();
		var xhr;
		var root;
		if (!select) {
			return;
		}
		select.selectedIndex = 0;
		while (select.options.length > 2) {
			select.remove(2);
		}
		setValue("hHeadCount", "0");
		if (!unit.id || unit.id === "A") {
			return;
		}
		xhr = getText("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(unit.id));
		if (xhr.responseText) {
			if (xmlObject("OutData") && typeof xmlObject("OutData").loadXML === "function") {
				xmlObject("OutData").loadXML(xhr.responseText);
			}
			root = xmlRoot("OutData");
			childElements(root).forEach(function (node) {
				select.options[select.options.length] = new Option(node.textContent || node.text || "", attr(node, "ParType") || attr(node, 0));
			});
		}
	};

	window.selAccountHead = function (objAcc) {
		var unit = field("selAccUnitId");
		if (selectedOptionIsAccountingUnitPrompt()) {
			if (objAcc) {
				objAcc.selectedIndex = 0;
			}
			if (unit) {
				unit.focus();
			}
			return;
		}
		window.DisplayBalamt();
		if (typeof baseSelAccountHead === "function") {
			baseSelAccountHead(objAcc);
		}
	};

	window.DisplayVoucher = function () {
		var root = xmlRoot("VoucherData");
		var table = clearVoucherTable();
		var entries = childElements(root, "Entry");
		var total = 0;
		var tdsTotal = 0;
		var voucherDate = attr(root, "VouDate");
		var display = byId("DisVoucher");
		var row;
		if (!table) {
			return;
		}
		if (voucherDate) {
			setDateValue(voucherDate);
		}
		if (display && display.style) {
			display.style.height = "200px";
			display.style.visibility = "visible";
		}
		entries.forEach(function (entry, index) {
			var entryNo = index + 1;
			var crdr = attr(entry, "CRDR");
			var amount = toNumber(attr(entry, "Amount"));
			var tdsAmount = toNumber(attr(entry, "TdsAmount"));
			setAttr(entry, "No", entryNo);
			if (crdr === "C") {
				total -= amount;
			} else {
				total += amount;
			}
			tdsTotal += tdsAmount;
			setValue("hPayTo", attr(entry, "Payto"));
			row = table.insertRow(table.rows.length);
			insertCell(row, entryNo, "ExcelSerial", "Center", "top");
			insertCell(row, '<a href="#" onclick="EditEntry(\'' + entryNo + '\'); return false;" class="ExcelDisplayCell"><b>Edit</b></a>', "ExcelDisplayCell", "Center", "top");
			insertCell(row, escapeHtml(attr(entry, "AccName")), "ExcelDisplayCell", "left", "top");
			insertCell(row, escapeHtml(entryAccountText(entry)), "ExcelDisplayCell", "left", "top");
			insertCell(row, entryAdditionalText(entry), "ExcelDisplayCell", "left", "top");
			insertCell(row, escapeHtml(entryNarration(entry)), "ExcelDisplayCell", "left", "top");
			insertCell(row, formatNumber(amount, 2) + "&nbsp;" + escapeHtml(crdr), "ExcelDisplayCell", "right", "top");
			insertCell(row, formatNumber(tdsAmount, 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, formatNumber(attr(entry, "TdsPercentage"), 2), "ExcelDisplayCell", "right", "top");
		});
		row = table.insertRow(table.rows.length);
		insertCell(row, "<b>Total</b>", "ExcelDisplayCell", "right", "top", 6);
		insertCell(row, '<input type="text" name="txtTotalAmt" value="' + formatNumber(total, 2) + '" size="13" class="Formelemread" style="text-align:right">', "ExcelDisplayCell", "right", "top");
		insertCell(row, formatNumber(tdsTotal, 2), "ExcelDisplayCell", "right", "top");
		insertCell(row, "", "ExcelDisplayCell", "right", "top");
		window.iEntryNo = entries.length;
		setValue("hEntryNo", String(entries.length + 1));
		setText("spAccHead", "");
		setText("spEntryNo", String(entries.length + 1));
	};

	window.UpdateXML = function () {
		var root = xmlRoot("VoucherData");
		var approver = field("selUserId");
		if (!root) {
			return;
		}
		setAttr(root, "UnitNo", valueOf("hOrgId", ""));
		setAttr(root, "UnitName", valueOf("hOrgName", ""));
		setAttr(root, "BookNo", valueOf("hBookcode", attr(root, "BookNo")));
		setAttr(root, "BookName", valueOf("hBookName", attr(root, "BookName")));
		setAttr(root, "CRDR", valueOf("hVouCRDR", checkedCRDR()));
		setAttr(root, "VouDate", currentDateValue());
		setAttr(root, "BookAcchead", bookAccHead());
		if (approver) {
			setAttr(root, "Approver", selectedValue(approver));
		}
	};

	window.CheckVouAmount = function () {
		var root = xmlRoot("VoucherData");
		var receiptTotal = 0;
		var paymentTotal = 0;
		childElements(root, "Entry").forEach(function (entry) {
			if (attr(entry, "CRDR") === "C") {
				receiptTotal += toNumber(attr(entry, "Amount"));
			} else {
				paymentTotal += toNumber(attr(entry, "Amount"));
			}
		});
		return String(valueOf("hVouCRDR", "")) === "D" ? receiptTotal - paymentTotal : paymentTotal - receiptTotal;
	};

	window.CheckAdjVal = function (total) {
		var docs = selectNodes(xmlRoot("VoucherData"), "//PayRec/Doc");
		var adjusted = 0;
		if (!docs.length) {
			return true;
		}
		docs.forEach(function (node) {
			adjusted += toNumber(attr(node, "AmtToAdjust"));
		});
		return toNumber(total) <= adjusted;
	};

	window.CheckVouStat = function () {
		var total = toNumber(window.CheckVouAmount());
		var selectedPayRec = toNumber(valueOf("hSelPayRecCount", "0"));
		var totalPayRec = toNumber(valueOf("hPayRecCount", "0"));
		var voucherDate = parseLegacyDate(currentDateValue());
		var currentDate = parseLegacyDate(valueOf("hCurrDate", ""));
		var balanceParts;
		if (voucherDate && currentDate && voucherDate > currentDate) {
			alert("Voucher Date Should be Less than the System Date ");
			return false;
		}
		if (total < 0) {
			alert(String(valueOf("hVouCRDR", "")) === "C" ? "Total Voucher Amount is more than the Payment Amount" : "Total Voucher Amount is more than the Receipt Amount");
			return false;
		}
		if (total === 0) {
			alert("Total Voucher amount should be More than Zero ");
			return false;
		}
		if (totalPayRec !== 0 && selectedPayRec === 0 && !confirm("Adjustment is Not made for the Party!!, Continue Without Adjustments? ")) {
			return false;
		}
		if (!window.CheckAdjVal(total) && !confirm("Payment Amount is made more than the bill value!!, Continue?  ")) {
			return false;
		}
		if (String(valueOf("hVouCRDR", "")) === "C") {
			balanceParts = textOf("spCurrBal").split(/\s+/);
			if (total > toNumber(balanceParts[0])) {
				alert("Voucher Amount is Greater than the Current Balance ");
				return false;
			}
		}
		return true;
	};

	window.checkFileds = function () {
		var limit = toNumber(valueOf("hTransLimit", "0"));
		var amount = toNumber(valueOf("txtAmount", "0"));
		var entryRoot = window.EntryRoot && window.EntryRoot.nodeType ? window.EntryRoot : xmlRoot("EntryData");
		if (typeof baseCheckFileds === "function" && !baseCheckFileds()) {
			return false;
		}
		if (limit && amount > limit) {
			if (String(window.sTransFlag || "A") === "W") {
				alert("Amount is greater than the amount limit");
			}
			if (String(window.sTransFlag || "A") === "R") {
				alert("Amount should be less than " + limit);
				return false;
			}
		}
		return childElements(entryRoot, "PayRec").every(function (payRec) {
			var totalToAdjust = 0;
			return childElements(payRec).every(function (node, index) {
				var input = payRecInput(node, index);
				var value = toNumber(input && input.value);
				var max = toNumber(attr(node, "TransAmount")) - toNumber(attr(node, "AmtAdjusted"));
				if (String(attr(node, "AdjType")) === "I") {
					max -= toNumber(attr(node, "AmtToAccount"));
				}
				if (value > max) {
					alert('"To Adjust Amount" should be less than "Document Amount-(Adjusted +To Account)"');
					if (input) {
						input.focus();
					}
					return false;
				}
				totalToAdjust += value;
				if (totalToAdjust > amount) {
					alert('Total of "To Adjust Amount" should be less than "Voucher Amount"');
					return false;
				}
				return true;
			});
		});
	};

	window.CheckFinDate = function () {
		if (!field("hFinFrm") || !field("hFinTo")) {
			return true;
		}
		return typeof baseCheckFinDate === "function" ? baseCheckFinDate() : true;
	};

	window.showNarration = function (bookCode) {
		var bookNo = (bookCode || "01") + "?" + valueOf("hBookcode", "");
		var url = "NarrationSelection.asp?orgId=" + encodeURIComponent(valueOf("hOrgId", "")) + "&BookCode=" + encodeURIComponent(bookNo);
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			window.ITMSModernCompat.openModalDialog(url, "", "dialogHeight:300px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No", function (narration) {
				if (narration) {
					setValue("txtNarration", narration);
				}
			});
		}
	};

	window.AddEntry = function (flag) {
		if (typeof baseAddEntry === "function") {
			baseAddEntry(flag);
		}
		if (flag !== "S") {
			setEntryButtons(false);
		}
	};

	window.EditEntry = function (entryNo, editType) {
		if (typeof baseEditEntry === "function") {
			baseEditEntry(entryNo, editType);
		}
		setEntryButtons(editType !== "D");
	};

	window.DelEntry = function () {
		if (typeof baseDelEntry === "function") {
			baseDelEntry();
		}
		setEntryButtons(false);
	};

	window.Init = function () {
		ensureCompat();
		if (typeof baseInit === "function") {
			baseInit();
		}
		window.bVouFlag = false;
		window.bSavFlag = false;
		window.bEditFlag = true;
		window.sTransFlag = window.sTransFlag || "A";
		window.UpdateXML();
		setEntryButtons(false);
	};

	function wireDateBlur() {
		var control = field("ctlDate") || byId("ctlDate");
		if (control && control.addEventListener) {
			control.addEventListener("blur", window.DisplayBalamt);
		}
	}

	if (document.readyState === "loading") {
		document.addEventListener("DOMContentLoaded", wireDateBlur);
	} else {
		wireDateBlur();
	}

	window.ITMSVouCAEntryCompat = {
		currentAccountingUnit: currentAccountingUnit,
		xml: function () {
			return xmlString(xmlDocument("VoucherData"));
		}
	};
}(window, document));
