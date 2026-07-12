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

	function pageConfig() {
		var options = window.CNCommisionEntryConfig || {};
		return {
			dataIsland: options.dataIsland || "DetData",
			entryIsland: options.entryIsland || "EntryData",
			saveIsland: options.saveIsland || options.dataIsland || "DetData",
			saveMod: options.saveMod || "CN",
			saveName: options.saveName || "Voucher Entry",
			checkApp: options.checkApp !== false,
			checkFinancialDate: options.checkFinancialDate !== false,
			forceMultiEntryFlow: options.forceMultiEntryFlow === true,
			includeAccUnit: options.includeAccUnit === true,
			setDateFromVoucher: options.setDateFromVoucher === true,
			setFinancialDateLimits: options.setFinancialDateLimits === true,
			alwaysDisplayVoucher: options.alwaysDisplayVoucher === true,
			showTotalRow: options.showTotalRow !== false,
			updateActionWhenFlagTrue: options.updateActionWhenFlagTrue === true,
			flagField: options.flagField || "hFlag",
			flagAction: options.flagAction || "AmdAccCrNtGenerate.asp"
		};
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

	function fields(name) {
		var item = field(name);
		if (item && item.length != null && !item.tagName) {
			return Array.prototype.slice.call(item);
		}
		return item ? [item] : [];
	}

	function valueOf(name, fallback) {
		var item = field(name);
		return item && item.value != null ? item.value : fallback || "";
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

	function createNode(xmlName, nodeName) {
		var object = xmlObject(xmlName);
		if (object && typeof object.createElement === "function") {
			return object.createElement(nodeName);
		}
		if (object && object.XMLDocument) {
			return object.XMLDocument.createElement(nodeName);
		}
		return document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function childElements(node, nodeName) {
		var wanted = nodeName ? String(nodeName).toLowerCase() : "";
		return Array.prototype.slice.call(node && node.childNodes || []).filter(function (child) {
			return child.nodeType === 1 && (!wanted || String(child.nodeName).toLowerCase() === wanted);
		});
	}

	function clearChildren(node) {
		while (node && node.firstChild) {
			node.removeChild(node.firstChild);
		}
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

	function attrAny(node) {
		var i;
		var value;
		for (i = 1; i < arguments.length; i += 1) {
			value = attr(node, arguments[i]);
			if (value !== "") {
				return value;
			}
		}
		return "";
	}

	function setExistingAttr(node, preferredName, aliases, value) {
		var names = [preferredName].concat(aliases || []);
		var i;
		if (!node || !node.attributes) {
			return;
		}
		for (i = 0; i < names.length; i += 1) {
			if (node.hasAttribute && node.hasAttribute(names[i])) {
				node.setAttribute(names[i], value == null ? "" : String(value));
				return;
			}
		}
		setAttr(node, preferredName, value);
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

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return null;
		}
		return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
	}

	function rootFromDialog(value) {
		return xmlRoot(value) || value && value.nodeType === 9 && value.documentElement || value && value.nodeType === 1 && value || null;
	}

	function importFor(parent, node) {
		if (!parent || !node) {
			return node;
		}
		if (node.ownerDocument !== parent.ownerDocument && parent.ownerDocument.importNode) {
			return parent.ownerDocument.importNode(node, true);
		}
		return node;
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
		return window.ITMSModernCompat && window.ITMSModernCompat.toDisplayDate ? window.ITMSModernCompat.toDisplayDate(control.value) : control.value || "";
	}

	function setDateControl(name, value) {
		var control = field(name) || byId(name);
		if (!control || trim(value) === "") {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else {
			control.value = value;
		}
	}

	function setDateLimits(name, minDate, maxDate) {
		var control = field(name) || byId(name);
		if (!control) {
			return;
		}
		if (typeof control.SetMinDate === "function") {
			control.SetMinDate(minDate);
		} else if (typeof control.setMinDate === "function") {
			control.setMinDate(minDate);
		}
		if (typeof control.SetMaxDate === "function") {
			control.SetMaxDate(maxDate);
		} else if (typeof control.setMaxDate === "function") {
			control.setMaxDate(maxDate);
		}
	}

	function parseDate(value) {
		var match = trim(value).match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		var year;
		if (!match) {
			return null;
		}
		year = Number(match[3]);
		if (year < 100) {
			year += 2000;
		}
		return new Date(year, Number(match[2]) - 1, Number(match[1]));
	}

	function initState() {
		var config = pageConfig();
		var root = xmlRoot(config.dataIsland) || xmlRoot("DetData");
		window.VouRoot = root;
		window.EntryRoot = window.EntryRoot && window.EntryRoot.nodeType ? window.EntryRoot : xmlRoot(config.entryIsland);
		if (!window.EntryRoot || String(window.EntryRoot.nodeName).toLowerCase() !== "entry") {
			window.clearXML();
		}
		if (!window.iEntryNo) {
			window.iEntryNo = childElements(root, "Entry").length + 1;
		}
		return { vouRoot: root, entryRoot: window.EntryRoot };
	}

	function nextEntryNo(root) {
		var max = 0;
		childElements(root, "Entry").forEach(function (entry) {
			max = Math.max(max, toNumber(attr(entry, "No")));
		});
		return max + 1 || 1;
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

	function selectedCrDr() {
		var selected = fields("OptCRDR").filter(function (item) {
			return item.checked;
		})[0];
		return selected ? selected.value : "D";
	}

	function setRadioValue(name, value) {
		fields(name).forEach(function (item) {
			item.checked = item.value === value;
		});
	}

	function clearGeneratedEntryChildren(entryRoot) {
		childElements(entryRoot, "Narration").forEach(function (node) {
			entryRoot.removeChild(node);
		});
	}

	function updateAnalysisAmounts(entryRoot) {
		childElements(entryRoot).forEach(function (header) {
			childElements(header).forEach(function (node) {
				var code = attr(node, "No") || attr(node, 0);
				var groupCode = attr(node, "GroupCode");
				var ratioField;
				var amountField;
				if (header.nodeName === "CostCenter") {
					ratioField = field("txtCCRatio" + code);
					amountField = field("txtCCAmount" + code);
				}
				if (header.nodeName === "Analytical") {
					ratioField = field("txtANALRatio" + code + "Z" + groupCode) || field("txtANALRatio" + code);
					amountField = field("txtANALAmount" + code + "Z" + groupCode) || field("txtANALAmount" + code);
				}
				if (ratioField) {
					setAttr(node, "Ratio", ratioField.value);
				}
				if (amountField) {
					setAttr(node, "Amount", amountField.value);
				}
			});
		});
	}

	function updateVoucherDates(root) {
		childElements(root, "voucher").forEach(function (voucher) {
			setAttr(voucher, "VouDate", getDateControl("ctlDate"));
		});
	}

	function resetEntryFields() {
		var account = field("selAccHead");
		if (account) {
			account.selectedIndex = 0;
		}
		setValue("txtPayTo", "");
		setValue("txtNarration", "");
		setValue("txtTdsAmount", "0.00");
		setValue("txtTdsper", "0.00");
		setValue("txtAmount", "0.00");
		setValue("hEditEntry", "0");
		setValue("hEntryNo", "0");
		window.setADDDisplay(0);
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

	function additionalDetails(entry) {
		var details = [];
		childElements(entry).forEach(function (header) {
			if (header.nodeName === "CostCenter" || header.nodeName === "Analytical") {
				if (header.nodeName === "Analytical" && childElements(header).length) {
					details.push("---------------------------");
				}
				childElements(header).forEach(function (node) {
					details.push(escapeHtml((attr(node, "ShortName") || attr(node, 2)) + "-" + (attr(node, "Ratio") || attr(node, 3)) + "% " + (attr(node, "Amount") || attr(node, 4))));
				});
			}
			if (header.nodeName === "PayRec") {
				childElements(header).forEach(function (node) {
					details.push(escapeHtml(attr(node, 1) + ":" + attr(node, 2) + "- " + attr(node, 5)));
				});
			}
		});
		return details.join("<br>");
	}

	function runStringDialog(url, nextPrefix, features, done) {
		openDialog(url, "", features, function (value) {
			var root = rootFromDialog(value);
			var text = trim(value);
			var parts;
			if (root && childElements(root).length) {
				done("", [], root);
				return;
			}
			if (!text) {
				return;
			}
			parts = text.split(":");
			if (parts.length <= 1 && nextPrefix) {
				runStringDialog(nextPrefix + text, nextPrefix, features, done);
				return;
			}
			done(text, parts, null);
		});
	}

	function retFromGlRoot(root) {
		var entry = childElements(root, "Entry")[0] || childElements(root)[0];
		if (!entry) {
			return "";
		}
		return [0, 1, 2, 3, 4, 5, 6, 7].map(function (index) {
			return attr(entry, "RetField" + index);
		}).join(":");
	}

	function appendAccHeadFromData() {
		var state = initState();
		childElements(xmlRoot("AccHeadData")).forEach(function (node) {
			state.entryRoot.appendChild(importFor(state.entryRoot, node));
		});
	}

	function applyTdsEligibility(value) {
		var eligible = trim(value) === "1";
		setValue("hTdsElgi", eligible ? "1" : "0");
		if (field("txtTdsAmount")) {
			field("txtTdsAmount").disabled = !eligible;
		}
		if (field("txtTdsper")) {
			field("txtTdsper").disabled = !eligible;
		}
	}

	window.showNarration = function () {
		var url = "NarrationSelection.asp?orgId=" + encodeURIComponent(valueOf("hOrgId")) + "&BookCode=" + encodeURIComponent("07?" + valueOf("hBookcode"));
		openDialog(url, "", "", function (narration) {
			if (trim(narration) !== "") {
				setValue("txtNarration", narration);
			}
		});
	};

	window.selGLHead = function (select) {
		var parts;
		var node;
		var state = initState();
		if (!select || select.selectedIndex <= 0) {
			return false;
		}
		if (select.value === "G") {
			return window.showGLHead(valueOf("hOrgId"));
		}
		parts = String(select.value || "").split("?");
		node = createNode(pageConfig().entryIsland, "AccHead");
		setAttr(node, "No", trim(parts[0]));
		setAttr(node, "CostCenter", trim(parts[1]));
		setAttr(node, "Analytical", trim(parts[2]));
		setAttr(node, "Name", selectedText(select));
		setAttr(node, "Type", "G");
		setAttr(node, "Group", "");
		state.entryRoot.appendChild(node);
		setValue("txtPayTo", selectedText(select));
		if (typeof window.showCCAnal === "function") {
			window.showCCAnal(valueOf("hOrgId"), trim(parts[0]), trim(parts[1]), trim(parts[2]));
		}
		return false;
	};

	window.showGLHead = function () {
		var base = "GLHeadSelection.asp";
		var url = base + "?orgId=" + encodeURIComponent(valueOf("hOrgId")) + "&BookId=07&BookNo=" + encodeURIComponent(valueOf("hBookcode"));
		runStringDialog(url, base + "?", "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (text, parts, root) {
			var ret = root ? retFromGlRoot(root) : text;
			var accNode;
			if (!ret || ret.split(":").length <= 1) {
				return;
			}
			parts = ret.split(":");
			applyTdsEligibility(parts[6]);
			clearChildren(xmlRoot("AccHeadData"));
			if (typeof window.GetGlHeadXml === "function") {
				window.GetGlHeadXml(ret);
			}
			window.clearXML();
			appendAccHeadFromData();
			accNode = childElements(window.EntryRoot, "AccHead")[0];
			if (accNode) {
				setValue("txtPayTo", attr(accNode, "Name") || attr(accNode, 3));
				if (typeof window.showCCAnal === "function") {
					window.showCCAnal(valueOf("hOrgId"), attr(accNode, "No") || attr(accNode, 0), attr(accNode, "CostCenter") || attr(accNode, 1), attr(accNode, "Analytical") || attr(accNode, 2));
				}
			} else {
				setValue("txtPayTo", "");
				window.setADDDisplay(0);
			}
		});
		return false;
	};

	window.checkFileds = function () {
		var voucherDate = parseDate(getDateControl("ctlDate"));
		var invoiceDate = parseDate(valueOf("hInvDate"));
		if (field("selAccHead") && field("selAccHead").selectedIndex === 0) {
			alert("Select Account Head ");
			field("selAccHead").focus();
			return false;
		}
		if (trim(valueOf("txtNarration")) === "") {
			alert("Enter Narration");
			if (field("txtNarration")) {
				field("txtNarration").select();
			}
			return false;
		}
		if (voucherDate && invoiceDate && voucherDate < invoiceDate) {
			alert("Credit Note date should be >= Invoice date");
			if (field("ctlDate")) {
				field("ctlDate").focus();
			}
			return false;
		}
		return true;
	};

	window.AddEntry = function (flag) {
		var config = pageConfig();
		if (config.forceMultiEntryFlow || valueOf("hVouchTy") === "SC") {
			return window.AddEntrySC(flag);
		}
		var state = initState();
		if (!window.checkFileds()) {
			return false;
		}
		setAttr(state.entryRoot, "No", "1");
		setExistingAttr(state.entryRoot, "Payto", ["PayTo"], valueOf("txtPayTo"));
		setAttr(state.entryRoot, "Amount", valueOf("txtAmount"));
		setAttr(state.entryRoot, "CRDR", selectedCrDr());
		if (config.includeAccUnit) {
			setAttr(state.entryRoot, "AccUnit", valueOf("hOrgId"));
			setAttr(state.entryRoot, "AccName", valueOf("hOrgName"));
		}
		setAttr(state.entryRoot, "TdsAmount", valueOf("txtTdsAmount", "0.00"));
		setExistingAttr(state.entryRoot, "TDSElgi", ["TdsElgi"], valueOf("hTdsElgi", "0"));
		setAttr(state.entryRoot, "TdsPercentage", valueOf("txtTdsper", "0.00"));
		updateVoucherDates(state.vouRoot);
		clearGeneratedEntryChildren(state.entryRoot);
		var narration = createNode(config.entryIsland, "Narration");
		narration.textContent = valueOf("txtNarration");
		state.entryRoot.appendChild(narration);
		updateAnalysisAmounts(state.entryRoot);
		state.vouRoot.appendChild(importFor(state.vouRoot, state.entryRoot));
		return window.SaveXML();
	};

	window.AddEntrySC = function (flag) {
		var config = pageConfig();
		var state = initState();
		var entry = state.entryRoot;
		var nextNode;
		var narration;
		if (String(flag) === "S" && field("selAccHead") && field("selAccHead").selectedIndex === 0) {
			return window.SaveXML();
		}
		if (!window.checkFileds()) {
			return false;
		}
		setAttr(entry, "No", String(flag) === "U" ? valueOf("hEditEntry") || valueOf("hEntryNo") : window.iEntryNo);
		setExistingAttr(entry, "Payto", ["PayTo"], valueOf("txtPayTo"));
		setAttr(entry, "Amount", valueOf("txtAmount"));
		setAttr(entry, "CRDR", selectedCrDr());
		if (config.includeAccUnit) {
			setAttr(entry, "AccUnit", valueOf("hOrgId"));
			setAttr(entry, "AccName", valueOf("hOrgName"));
		}
		setAttr(entry, "TdsAmount", valueOf("txtTdsAmount", "0.00"));
		setExistingAttr(entry, "TDSElgi", ["TdsElgi"], valueOf("hTdsElgi", "0"));
		setAttr(entry, "TdsPercentage", valueOf("txtTdsper", "0.00"));
		updateVoucherDates(state.vouRoot);
		clearGeneratedEntryChildren(entry);
		narration = createNode(config.entryIsland, "Narration");
		narration.textContent = valueOf("txtNarration");
		entry.appendChild(narration);
		updateAnalysisAmounts(entry);
		if (String(flag) === "U") {
			nextNode = window._CNCommEditNextNode;
			if (nextNode && nextNode.parentNode === state.vouRoot) {
				state.vouRoot.insertBefore(importFor(state.vouRoot, entry), nextNode);
			} else {
				state.vouRoot.appendChild(importFor(state.vouRoot, entry));
			}
			window._CNCommEditNextNode = null;
		} else {
			state.vouRoot.appendChild(importFor(state.vouRoot, entry));
		}
		if (String(flag) === "S") {
			return window.SaveXML();
		}
		window.DisplayVoucher();
		window.iEntryNo = nextEntryNo(state.vouRoot);
		window.clearXML();
		resetEntryFields();
		setEntryButtons(false);
		return false;
	};

	window.DisplayVoucher = function () {
		var config = pageConfig();
		var root = initState().vouRoot;
		var table = clearTable("tblVoucher", 1, 1);
		var display = byId("DisVoucher");
		var total = 0;
		var row;
		var rowNo = 1;
		if (display) {
			display.style.height = "200px";
			display.style.visibility = "visible";
		}
		if (!table) {
			return;
		}
		if (config.setDateFromVoucher) {
			setDateControl("ctlDate", attr(root, "VouDate") || attr(root, 5));
		}
		childElements(root, "Entry").forEach(function (entry) {
			var amount = toNumber(attr(entry, "Amount"));
			var tdsAmount = formatNumber(attr(entry, "TdsAmount"), 2);
			var tdsPercent = formatNumber(attr(entry, "TdsPercentage"), 2);
			var account = "";
			var narration = "";
			setAttr(entry, "No", rowNo);
			total += attr(entry, "CRDR") === "D" ? amount : -amount;
			childElements(entry).forEach(function (node) {
				if (node.nodeName === "AccHead") {
					account = attr(node, "Name") || attr(node, 3);
				}
				if (node.nodeName === "Narration") {
					narration = node.textContent || "";
				}
			});
			row = table.insertRow(table.rows.length);
			insertCell(row, rowNo, "ExcelSerial", "Center", "top");
			insertCell(row, '<a href="#" onclick="EditEntry(\'' + rowNo + '\'); return false;">Edit</a>', "ExcelDisplayCell", "Center", "top");
			insertCell(row, escapeHtml(valueOf("hOrgName")), "ExcelDisplayCell", "left", "top");
			insertCell(row, escapeHtml(account), "ExcelDisplayCell", "left", "top");
			insertCell(row, escapeHtml(narration), "ExcelDisplayCell", "left", "top");
			insertCell(row, formatNumber(amount, 2), "ExcelDisplayCell", "right", "top");
			insertCell(row, additionalDetails(entry), "ExcelDisplayCell", "left", "top");
			insertCell(row, tdsAmount, "ExcelDisplayCell", "left", "top");
			insertCell(row, tdsPercent, "ExcelDisplayCell", "left", "top");
			rowNo += 1;
		});
		if (config.showTotalRow) {
			row = table.insertRow(table.rows.length);
			insertCell(row, "<b>Total</b>", "ExcelDisplayCell", "right", "top", 5);
			insertCell(row, "Rs. &nbsp;" + formatNumber(Math.abs(total), 2) + (total < 0 ? "&nbsp;Cr" : "&nbsp;Dr"), "ExcelDisplayCell", "right", "top");
			insertCell(row, "", "ExcelDisplayCell", "right", "top", 3);
		}
	};

	window.EditEntry = function (entryNo) {
		var root = initState().vouRoot;
		var entry = childElements(root, "Entry").filter(function (node) {
			return String(attr(node, "No")) === String(entryNo);
		})[0];
		if (!entry) {
			return false;
		}
		window.setADDDisplay(0);
		setValue("hEditEntry", entryNo);
		setValue("hEntryNo", entryNo);
		setValue("txtAmount", attr(entry, "Amount"));
		setRadioValue("OptCRDR", attr(entry, "CRDR") || "D");
		setValue("txtPayTo", attrAny(entry, "Payto", "PayTo"));
		setValue("txtTdsAmount", attr(entry, "TdsAmount") || "0.00");
		setValue("txtTdsper", attr(entry, "TdsPercentage") || "0.00");
		applyTdsEligibility(attrAny(entry, "TDSElgi", "TdsElgi"));
		childElements(entry).forEach(function (node) {
			if (node.nodeName === "AccHead" && field("selAccHead")) {
				field("selAccHead").selectedIndex = 1;
			}
			if (node.nodeName === "Narration") {
				setValue("txtNarration", node.textContent || "");
			}
			if (node.nodeName === "CostCenter" && typeof window.popCostCenter === "function") {
				window.setADDDisplay(1);
				window.popCostCenter(node);
			}
			if (node.nodeName === "Analytical" && typeof window.popAnalytical === "function") {
				window.setADDDisplay(1);
				window.popAnalytical(node);
			}
		});
		window._CNCommEditNextNode = entry.nextSibling;
		window.EntryRoot = root.removeChild(entry);
		setEntryButtons(true);
		window.DisplayVoucher();
		return false;
	};

	window.DelEntry = function () {
		window._CNCommEditNextNode = null;
		window.clearXML();
		window.setADDDisplay(0);
		resetEntryFields();
		setEntryButtons(false);
		window.DisplayVoucher();
		return false;
	};

	window.SaveXML = function () {
		var config = pageConfig();
		var xhr;
		if (config.checkApp && typeof window.CheckApp === "function" && !window.CheckApp()) {
			return false;
		}
		if (config.checkFinancialDate && !window.CheckFinDate()) {
			return false;
		}
		xhr = syncPost("XMLSave.asp?Mod=" + encodeURIComponent(config.saveMod) + "&Name=" + encodeURIComponent(config.saveName), serializeXml(config.saveIsland));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return false;
		}
		if (field("btnNext")) {
			field("btnNext").disabled = true;
		}
		if (field("btnAdd")) {
			field("btnAdd").disabled = true;
		}
		if (config.updateActionWhenFlagTrue && trim(valueOf(config.flagField)) === "True") {
			form().action = config.flagAction;
		}
		form().submit();
		return true;
	};

	window.clearXML = function () {
		var config = pageConfig();
		var entry = createNode(config.entryIsland, "Entry");
		setAttr(entry, "No", window.iEntryNo || 1);
		if (config.includeAccUnit) {
			setAttr(entry, "CRDR", "");
			setAttr(entry, "Payto", "");
			setAttr(entry, "Amount", "");
			setAttr(entry, "AccUnit", "");
			setAttr(entry, "AccName", "");
		} else {
			setAttr(entry, "Payto", "");
			setAttr(entry, "Amount", "");
			setAttr(entry, "CRDR", "");
		}
		setAttr(entry, "TdsAmount", "");
		setAttr(entry, "TDSElgi", "0");
		setAttr(entry, "TdsPercentage", "");
		window.EntryRoot = entry;
		return entry;
	};

	window.setAnalDisplay = function (display, flag) {
		var visible = Number(flag) !== 0;
		var item = byId(String(display) === "A" ? "DisAnal" : "DisCost");
		if (item) {
			item.style.height = visible ? "100px" : "1px";
			item.style.width = visible ? "280px" : "1px";
			item.style.visibility = visible ? "visible" : "hidden";
		}
	};

	window.setADDDisplay = function (flag) {
		var visible = Number(flag) !== 0;
		var additional = byId("Disaddtional");
		var ccanl = byId("DisCCANL");
		if (additional) {
			additional.style.height = visible ? "115px" : "1px";
			additional.style.visibility = visible ? "visible" : "hidden";
		}
		if (ccanl) {
			ccanl.style.height = visible ? "114px" : "1px";
			ccanl.style.visibility = visible ? "visible" : "hidden";
		}
	};

	window.CancelAction = function (page) {
		form().action = page;
		form().submit();
	};

	window.SelMisParty = function () {
		runStringDialog("MisPartySelection.asp", "MisPartySelection.asp?", "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (text, parts) {
			if (text === "AN") {
				window.AddNewParty();
				return;
			}
			if (parts.length > 1) {
				setValue("txtPayTo", parts[0]);
			}
		});
		return false;
	};

	window.AddNewParty = function () {
		openDialog("MisParCreate.asp", "", "dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (value) {
			if (trim(value) !== "") {
				setValue("txtPayTo", value);
			}
		});
		return false;
	};

	window.CheckFinDate = function () {
		var finFrom = toNumber(valueOf("hFinFrm"));
		var finTo = toNumber(valueOf("hFinTo"));
		var parts = String(getDateControl("ctlDate") || "").split("/");
		var current = toNumber((parts[2] || "") + (parts[1] || ""));
		if (current < finFrom || current > finTo) {
			alert("Voucher Date Should Be Between 01/04/" + String(finFrom).substring(0, 4) + " To 31/03/" + String(finTo).substring(0, 4));
			return false;
		}
		return true;
	};

	window.InitVouCNCommisionEntry = function () {
		var config = pageConfig();
		var root;
		ensureCompat();
		if (config.setFinancialDateLimits) {
			setDateLimits("ctlDate", "01/04/" + trim(valueOf("hFromYr")), "31/03/" + trim(valueOf("hToYr")));
		}
		initState();
		applyTdsEligibility(valueOf("hTdsElgi", "0"));
		root = xmlRoot(config.dataIsland);
		if (config.alwaysDisplayVoucher || childElements(root, "Entry").length) {
			window.DisplayVoucher();
		}
	};
}(window, document));
