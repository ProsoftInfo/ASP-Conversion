(function (window, document) {
	"use strict";

	function trim(value) {
		return String(value == null ? "" : value).replace(/^\s+|\s+$/g, "");
	}

	function isAmendMode() {
		return window.VouDNOthersEntryMode === "amend";
	}

	function toNumber(value) {
		var parsed = parseFloat(String(value == null ? "" : value).replace(/,/g, ""));
		return isNaN(parsed) ? 0 : parsed;
	}

	function formatNumber(value) {
		return toNumber(value).toFixed(2);
	}

	function parseDate(value) {
		var text = trim(value);
		var match;
		var parsed;
		if (!text) {
			return null;
		}
		match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
		if (match) {
			return new Date(Number(match[3]) < 100 ? Number(match[3]) + 2000 : Number(match[3]), Number(match[2]) - 1, Number(match[1]));
		}
		parsed = new Date(text);
		return isNaN(parsed.getTime()) ? null : parsed;
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

	function setText(id, value) {
		var item = byId(id);
		if (item) {
			item.innerHTML = value == null ? "" : String(value);
		}
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

	function xmlDocument(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
		return object && object.XMLDocument || object && object._doc || object && object.nodeType === 9 && object || null;
	}

	function xmlRoot(nameOrObject) {
		var object = typeof nameOrObject === "string" ? xmlObject(nameOrObject) : nameOrObject;
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

	function clearXml(name) {
		var doc = xmlDocument(name);
		while (doc && doc.firstChild) {
			doc.removeChild(doc.firstChild);
		}
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
		if (node) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
	}

	function importFor(parent, node) {
		if (node && parent && parent.ownerDocument && parent.ownerDocument.importNode && node.ownerDocument !== parent.ownerDocument) {
			return parent.ownerDocument.importNode(node, true);
		}
		return node;
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

	function syncGet(url) {
		var xhr = new XMLHttpRequest();
		xhr.open("GET", url, false);
		xhr.send();
		return xhr.responseText || "";
	}

	function openDialog(url, args, features, callback) {
		ensureCompat();
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
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

	function setDateLimits() {
		var control = field("ctlDate") || byId("ctlDate");
		var from = trim(valueOf("hFromYr"));
		var to = trim(valueOf("hToYr"));
		if (!control || !from || !to) {
			return;
		}
		if (typeof control.SetMinDate === "function") {
			control.SetMinDate("01/04/" + from);
		}
		if (typeof control.SetMaxDate === "function") {
			control.SetMaxDate("31/03/" + to);
		}
	}

	function initState() {
		var root = xmlRoot("VoucherData");
		window.VouRoot = root;
		window.EntryRoot = window.EntryRoot && window.EntryRoot.nodeType ? window.EntryRoot : xmlRoot("EntryData");
		if (!window.EntryRoot || String(window.EntryRoot.nodeName).toLowerCase() !== "entry") {
			window.clearXML();
		}
		window.iEntryNo = Number(window.iEntryNo || childElements(root, "Entry").length || 0);
		if (!window.bEditFlag && window.bEditFlag !== false) {
			window.bEditFlag = true;
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

	function entryNarration(entry) {
		var narration = "";
		childElements(entry, "Narration").forEach(function (node) {
			narration = node.textContent || "";
		});
		return narration;
	}

	function entryAccount(entry) {
		var account = "";
		childElements(entry, "AccHead").forEach(function (node) {
			account = attr(node, "Type") === "P" ? attr(node, "Name") || attr(node, 3) : (attr(node, "No") || attr(node, 0)) + " - " + (attr(node, "Name") || attr(node, 3));
		});
		return account;
	}

	function additionalDetails(entry) {
		var details = [];
		childElements(entry).forEach(function (header) {
			if (header.nodeName === "CostCenter" || header.nodeName === "Analytical") {
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

	function clearGeneratedEntryChildren(entryRoot) {
		childElements(entryRoot, "Narration").forEach(function (node) {
			entryRoot.removeChild(node);
		});
	}

	function resetEntryFields() {
		if (field("selAccHead")) {
			field("selAccHead").selectedIndex = 0;
		}
		setValue("txtPayTo", "");
		setValue("txtNarration", "");
		setValue("txtAmount", "0.00");
		setValue("hAction", "New");
		window.setADDDisplay(0);
		window.clearXML();
	}

	function appendAccHeadFromData() {
		var state = initState();
		childElements(xmlRoot("AccHeadData")).forEach(function (node) {
			state.entryRoot.appendChild(importFor(state.entryRoot, node));
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

	function selectedBookAccHead() {
		return valueOf("hBookAccHead", "");
	}

	function addNarration(entryRoot) {
		var narration = createNode("EntryData", "Narration");
		narration.textContent = valueOf("txtNarration");
		entryRoot.appendChild(narration);
	}

	function fillEntry(entryRoot, flag) {
		var no = String(flag) === "U" ? valueOf("hEntryNo") : nextEntryNo(xmlRoot("VoucherData"));
		setAttr(entryRoot, "No", no);
		setAttr(entryRoot, "CRDR", "C");
		setAttr(entryRoot, "Payto", valueOf("txtPayTo"));
		setAttr(entryRoot, "Amount", valueOf("txtAmount"));
		setAttr(entryRoot, "AccUnit", valueOf("hOrgId"));
		setAttr(entryRoot, "AccName", valueOf("hOrgName"));
		clearGeneratedEntryChildren(entryRoot);
		addNarration(entryRoot);
		updateAnalysisAmounts(entryRoot);
		return no;
	}

	function appendEntryInOrder(root, entry, entryNo) {
		var next = childElements(root, "Entry").filter(function (node) {
			return toNumber(attr(node, "No")) === toNumber(entryNo) + 1;
		})[0];
		if (next) {
			root.insertBefore(importFor(root, entry), next);
		} else {
			root.appendChild(importFor(root, entry));
		}
	}

	function finishPost(url, xml, action) {
		var xhr = syncPost(url, xml);
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
		form().action = action;
		form().submit();
		return true;
	}

	function saveAmendVoucher() {
		var root = xmlRoot("VoucherData");
		var book = field("selBook");
		var bookAmend = trim(valueOf("hFlag")) === "True";
		var action = bookAmend ? "AmdAccDbNtGenerate.asp" : (form().action || "VouDNOTAmdGenerate.asp");
		if (root && bookAmend) {
			if (book) {
				setAttr(root, "BookNo", book.value);
				setAttr(root, "BookName", book.options[book.selectedIndex] ? book.options[book.selectedIndex].text : "");
			}
			setAttr(root, "VouDate", getDateControl("ctlDate"));
		}
		return finishPost("XMLSave.asp?Mod=DN&Name=Voucher%20AMD", serializeXml("VoucherData"), action);
	}

	function appendGjEntry(doc, root, no, crdr, payTo, amount, unitNo, unitName) {
		var entry = doc.createElement("Entry");
		setAttr(entry, "No", no);
		setAttr(entry, "CRDR", crdr);
		setAttr(entry, "Payto", payTo);
		setAttr(entry, "Amount", amount);
		setAttr(entry, "AccUnit", unitNo);
		setAttr(entry, "AccName", unitName);
		setAttr(entry, "TdsAmount", "0.00");
		setAttr(entry, "TDSElgi", "0");
		setAttr(entry, "TdsPercentage", "0");
		setAttr(entry, "PayRecAmount", "0");
		root.appendChild(entry);
		return entry;
	}

	function buildGeneralJournalVoucher() {
		var sourceRoot = xmlRoot("VoucherData");
		var doc = xmlDocument("GJVoucher");
		var root;
		var partyParts;
		var partyDetails;
		var payRecCount;
		var sourceEntry;
		var firstEntry;
		var secondEntry;
		var accHead;
		var narration;
		var unitNo;
		var unitName;
		var amount;
		if (!sourceRoot || !doc) {
			return null;
		}
		clearXml("GJVoucher");
		root = doc.createElement("voucher");
		doc.appendChild(root);
		unitNo = attr(sourceRoot, "UnitNo");
		unitName = attr(sourceRoot, "UnitName");
		setAttr(root, "UnitNo", unitNo);
		setAttr(root, "UnitName", unitName);
		setAttr(root, "BookNo", attr(sourceRoot, "BookNo"));
		setAttr(root, "BookName", attr(sourceRoot, "BookName"));
		setAttr(root, "CRDR", "");
		setAttr(root, "VouDate", attr(sourceRoot, "VouDate"));
		setAttr(root, "BookAcchead", "0");
		setAttr(root, "Approver", attr(sourceRoot, "Approver"));
		sourceEntry = childElements(sourceRoot, "Entry")[0];
		if (!sourceEntry) {
			return doc;
		}
		amount = attr(sourceEntry, "Amount");
		partyParts = String(attr(sourceRoot, "PartyCode")).split("?");
		partyDetails = unitNo + "&ParSubType=" + (partyParts[1] || "") + "&ParType=" + (partyParts[0] || "") + "&PartyCode=" + (partyParts[3] || "");
		payRecCount = syncGet("XMLGetPayRecCount.asp?orgID=" + partyDetails).split(":");
		firstEntry = appendGjEntry(doc, root, "1", "D", "0", amount, unitNo, unitName);
		accHead = doc.createElement("AccHead");
		setAttr(accHead, "No", attr(sourceRoot, "PartyCode"));
		setAttr(accHead, "Pay", payRecCount[0] || "");
		setAttr(accHead, "Rec", payRecCount[1] || "");
		setAttr(accHead, "Name", attr(sourceRoot, "PartyName"));
		setAttr(accHead, "Type", "P");
		setAttr(accHead, "Adv", payRecCount[2] || "");
		firstEntry.appendChild(accHead);
		accHead = doc.createElement("RecCount");
		setAttr(accHead, "Val", "1");
		firstEntry.appendChild(accHead);
		narration = doc.createElement("Narration");
		narration.textContent = valueOf("txtNarration");
		firstEntry.appendChild(narration);
		secondEntry = appendGjEntry(doc, root, "2", "C", "", amount, unitNo, unitName);
		childElements(sourceEntry, "AccHead").forEach(function (sourceAcc) {
			var node = doc.createElement("AccHead");
			setAttr(node, "No", attr(sourceAcc, "No"));
			setAttr(node, "CostCenter", attr(sourceAcc, "CostCenter"));
			setAttr(node, "Analytical", attr(sourceAcc, "Analytical"));
			setAttr(node, "Name", attr(sourceAcc, "Name"));
			setAttr(node, "Type", "G");
			setAttr(node, "TransFlag", "A");
			secondEntry.appendChild(node);
		});
		return doc;
	}

	window.showNarration = function () {
		var url = "NarrationSelection.asp?orgId=" + encodeURIComponent(valueOf("hOrgId")) + "&BookCode=" + encodeURIComponent("06?" + valueOf("hBookcode"));
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
		node = createNode("EntryData", "AccHead");
		setAttr(node, "No", trim(parts[0]));
		setAttr(node, "CostCenter", trim(parts[1]));
		setAttr(node, "Analytical", trim(parts[2]));
		setAttr(node, "Name", select.options[select.selectedIndex].text);
		setAttr(node, "Type", "G");
		setAttr(node, "Group", "");
		state.entryRoot.appendChild(node);
		setValue("txtPayTo", select.options[select.selectedIndex].text);
		if (typeof window.showCCAnal === "function") {
			window.showCCAnal(valueOf("hOrgId"), trim(parts[0]), trim(parts[1]), trim(parts[2]));
		}
		return false;
	};

	window.showGLHead = function (orgId) {
		var base = "../../Common/GLHeadSelection.asp";
		var url = base + "?orgId=" + encodeURIComponent(orgId || valueOf("hOrgId")) + "&BookId=01&BookNo=" + encodeURIComponent(valueOf("hBookcode")) + "&AccHead=" + encodeURIComponent(selectedBookAccHead());
		function handle(value) {
			var root = xmlRoot(value) || value && value.nodeType === 1 && value || null;
			var ret = root ? retFromGlRoot(root) : String(value || "");
			var action = root ? trim(attr(root, "Action")).toUpperCase() : "";
			var passQuery = root ? attr(root, "PassQuery") : "";
			var accNode;
			if (action === "CLOSE") {
				return;
			}
			if (action && action !== "DONE" && passQuery) {
				openDialog(base + "?" + passQuery, xmlObject("GLHeadData"), "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", handle);
				return;
			}
			if (!ret || ret.split(":").length <= 1) {
				return;
			}
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
		}
		openDialog(url, xmlObject("GLHeadData"), "dialogHeight:480px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", handle);
		return false;
	};

	window.checkFileds = function () {
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
		if (isAmendMode()) {
			var voucherDate = parseDate(getDateControl("ctlDate"));
			var invoiceDate = parseDate(valueOf("hInvDate"));
			if (voucherDate && invoiceDate && voucherDate.getTime() < invoiceDate.getTime()) {
				alert("Credit Note date should be >= Invoice date");
				if (field("ctlDate")) {
					field("ctlDate").focus();
				}
				return false;
			}
		}
		return true;
	};

	window.AddNew = function () {
		return window.AddEntry(trim(valueOf("hAction")) === "Edit" ? "U" : "A");
	};

	window.AddEntry = function (flag) {
		var state = initState();
		var entry = state.entryRoot;
		var entryNo;
		if (String(flag) === "S" && trim(valueOf("txtAmount")) === "0.00") {
			return window.SaveXML();
		}
		if (!window.checkFileds()) {
			return false;
		}
		setAttr(state.vouRoot, "VouDate", getDateControl("ctlDate"));
		setAttr(state.vouRoot, "UnitNo", valueOf("hOrgId"));
		setAttr(state.vouRoot, "UnitName", valueOf("hOrgName"));
		entryNo = fillEntry(entry, flag);
		if (String(flag) === "U") {
			appendEntryInOrder(state.vouRoot, entry, entryNo);
			setValue("hAction", "New");
		} else {
			state.vouRoot.appendChild(importFor(state.vouRoot, entry));
		}
		if (String(flag) === "S") {
			return window.SaveXML();
		}
		window.DisplayVoucher();
		window.iEntryNo = nextEntryNo(state.vouRoot) - 1;
		resetEntryFields();
		return false;
	};

	window.DisplayVoucher = function () {
		var root = initState().vouRoot;
		var table = clearTable("tblVoucher", 1, 1);
		var display = byId("DisVoucher");
		var rowNo = 1;
		var height;
		if (!table) {
			return;
		}
		height = 60 + childElements(root, "Entry").length * 25;
		if (display) {
			display.style.height = height + "px";
			display.style.visibility = "visible";
		}
		childElements(root, "Entry").forEach(function (entry) {
			var row = table.insertRow(table.rows.length);
			var amount = formatNumber(attr(entry, "Amount")) + "&nbsp;" + (attr(entry, "CRDR") || "C");
			setAttr(entry, "No", rowNo);
			insertCell(row, rowNo, "ExcelSerial", "Center", "top");
			insertCell(row, '<img src="../../assets/images/iTMS%20icons/Deleteicon.gif" onclick="EditEntry(\'' + rowNo + '\',\'D\')">', "ExcelDisplayCell", "Center", "top");
			insertCell(row, '<a class="ExcelDisplaylink" href="#" onclick="EditEntry(\'' + rowNo + '\',\'E\'); return false;">Edit</a>', "ExcelDisplayCell", "Center", "top");
			insertCell(row, escapeHtml(entryAccount(entry)), "ExcelDisplayCell", "left", "top");
			insertCell(row, escapeHtml(entryNarration(entry)), "ExcelDisplayCell", "left", "top");
			insertCell(row, amount, "ExcelDisplayCell", "right", "top");
			insertCell(row, additionalDetails(entry), "ExcelDisplayCell", "left", "top");
			rowNo += 1;
		});
		window.iEntryNo = rowNo - 1;
		setText("spEntryNo", String(rowNo));
	};

	window.SaveXML = function () {
		var doc;
		if (isAmendMode()) {
			return saveAmendVoucher();
		}
		if (trim(valueOf("hCallFrom")) === "DR") {
			if (typeof window.CheckApp === "function" && !window.CheckApp()) {
				return false;
			}
			return finishPost("XMLSave.asp?Mod=DN&Name=Voucher%20Entry", serializeXml("VoucherData"), "VouCNGenerate.asp");
		}
		doc = buildGeneralJournalVoucher();
		if (!doc) {
			alert("Voucher XML data is not available.");
			return false;
		}
		return finishPost("XMLSave.asp?Mod=GJ&Name=Voucher%20Entry", new XMLSerializer().serializeToString(doc), "VouGenerate.asp");
	};

	window.clearXML = function () {
		var entry = createNode("EntryData", "Entry");
		setAttr(entry, "No", window.iEntryNo || 1);
		setAttr(entry, "CRDR", "");
		setAttr(entry, "Payto", "");
		setAttr(entry, "Amount", "");
		setAttr(entry, "AccUnit", "");
		setAttr(entry, "AccName", "");
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

	window.EditEntry = function (entryNo, editType) {
		var root = initState().vouRoot;
		var entry = childElements(root, "Entry").filter(function (node) {
			return String(attr(node, "No")) === String(entryNo);
		})[0];
		if (!entry) {
			return false;
		}
		setValue("hEntryNo", entryNo);
		window.setADDDisplay(0);
		setDateControl("ctlDate", attr(root, "VouDate"));
		root.removeChild(entry);
		window.EntryRoot = entry;
		if (String(editType) === "D") {
			window.DelEntry();
			return false;
		}
		setValue("hAction", "Edit");
		setValue("txtAmount", attr(entry, "Amount"));
		setValue("txtPayTo", attr(entry, "Payto"));
		setValue("txtNarration", entryNarration(entry));
		childElements(entry).forEach(function (node) {
			if (node.nodeName === "AccHead" && field("selAccHead")) {
				field("selAccHead").selectedIndex = 1;
			}
			if (node.nodeName === "CostCenter" && typeof window.popCostCenter === "function") {
				window.setADDDisplay(1);
				window.popCostCenter(node);
			}
			if (node.nodeName === "Analytical" && typeof window.popAnalytical === "function") {
				window.setADDDisplay(1);
				window.popAnalytical(node);
			}
			if (node.nodeName === "PayRec" && typeof window.popPayRec === "function") {
				window.popPayRec(node);
			}
		});
		window.DisplayVoucher();
		return false;
	};

	window.DelEntry = function () {
		window.clearXML();
		window.setADDDisplay(0);
		resetEntryFields();
		window.DisplayVoucher();
		window.bVouFlag = false;
		window.bEditFlag = true;
		window.bSavFlag = true;
		return false;
	};

	window.InitVouDNOthersEntry = function () {
		window.VouDNOthersEntryMode = "";
		ensureCompat();
		setDateLimits();
		initState();
	};

	window.InitVouDNOtherAmend = function () {
		window.VouDNOthersEntryMode = "amend";
		ensureCompat();
		setDateLimits();
		initState();
		setDateControl("ctlDate", attr(xmlRoot("VoucherData"), "VouDate") || valueOf("hInvDate"));
		window.DisplayVoucher();
	};
}(window, document));
