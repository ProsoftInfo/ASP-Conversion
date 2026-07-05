(function (window, document) {
	"use strict";

	var config = {
		bookCode: "01",
		moduleCode: "CA",
		bank: false,
		enableTds: false,
		saveName: "Voucher Entry",
		actionUrl: "VouMsiGenerate.asp",
		transLimit: 0
	};

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
		if (!element) {
			return;
		}
		element.textContent = value == null ? "" : String(value);
	}

	function textOf(id) {
		var element = byId(id);
		return element ? trim(element.textContent || element.innerText || "") : "";
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

	function checkedValue(name, fallback) {
		var item = field(name);
		if (item && item.length != null) {
			for (var i = 0; i < item.length; i += 1) {
				if (item[i].checked) {
					return item[i].value;
				}
			}
		}
		return valueOf(name, fallback || "");
	}

	function ensureCompat() {
		if (window.ITMSModernCompat && window.ITMSModernCompat.init) {
			window.ITMSModernCompat.init(document);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.upgradeXmlIslands) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
		}
	}

	function xmlObject(name) {
		ensureCompat();
		return window[name] || document[name] || byId(name) || null;
	}

	function xmlRoot(name) {
		var object = xmlObject(name);
		return object && object.documentElement || object && object.XMLDocument && object.XMLDocument.documentElement || object && object._doc && object._doc.documentElement || null;
	}

	function xmlDocument(name) {
		var object = xmlObject(name);
		return object && object.XMLDocument || object && object._doc || object || null;
	}

	function createNode(xmlName, nodeName) {
		var object = xmlObject(xmlName);
		if (object && typeof object.createElement === "function") {
			return object.createElement(nodeName);
		}
		if (object && object.XMLDocument && object.XMLDocument.createElement) {
			return object.XMLDocument.createElement(nodeName);
		}
		return document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function clearXmlChildren(xmlName) {
		var root = xmlRoot(xmlName);
		if (!root) {
			return;
		}
		while (root.firstChild) {
			root.removeChild(root.firstChild);
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
		if (node && node.setAttribute) {
			node.setAttribute(name, value == null ? "" : String(value));
		}
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

	function xmlString(nodeOrDoc) {
		if (!nodeOrDoc) {
			return "";
		}
		if (typeof nodeOrDoc.xml === "string") {
			return nodeOrDoc.xml;
		}
		return new XMLSerializer().serializeToString(nodeOrDoc);
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

	function createHttp() {
		return window.CreateObject ? window.CreateObject("MSXML2.XMLHTTP") : new XMLHttpRequest();
	}

	function getText(url) {
		var xhr = createHttp();
		xhr.open("GET", url, false);
		xhr.send(null);
		return xhr;
	}

	function postXml(url, xml) {
		var xhr = createHttp();
		xhr.open("POST", url, false);
		xhr.send(xml);
		return xhr;
	}

	function popupSize(type, fallbackProgram, fallbackHeight, fallbackWidth) {
		var value = typeof window.GetWindowSizeForPopup === "function" ? window.GetWindowSizeForPopup(String(type)) : "";
		var parts = String(value || "").split(":");
		return {
			program: parts[0] || fallbackProgram,
			height: parts[1] || fallbackHeight || "500",
			width: parts[2] || fallbackWidth || "500"
		};
	}

	function openDialog(url, args, features, callback) {
		if (window.ITMSModernCompat && window.ITMSModernCompat.openModalDialog) {
			return window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
		}
		alert("Modern browser compatibility script is still loading. Please try again.");
		return null;
	}

	function currentOrgId() {
		return valueOf("hOrgId", valueOf("hOrgID", ""));
	}

	function currentBookParts() {
		var value = selectedValue(field("selBook"));
		return String(value || "").split(value && value.indexOf("?") !== -1 ? "?" : "-");
	}

	function currentDateValue(controlName) {
		var control = field(controlName || "ctlDate") || byId(controlName || "ctlDate");
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

	function setDateValue(controlName, value) {
		var control = field(controlName || "ctlDate") || byId(controlName || "ctlDate");
		if (!control || !value) {
			return;
		}
		if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else if (typeof control.setDate === "function") {
			control.setDate(value);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate) {
			control.value = window.ITMSModernCompat.toIsoDate(value);
		} else {
			control.value = value;
		}
	}

	function setDateLimit(controlName, kind, value) {
		var control = field(controlName || "ctlDate") || byId(controlName || "ctlDate");
		var method;
		if (!control || !value) {
			return;
		}
		method = kind === "min" ? "SetMinDate" : "SetMaxDate";
		if (typeof control[method] === "function") {
			control[method](value);
		} else if (window.ITMSModernCompat && window.ITMSModernCompat.toIsoDate) {
			control[kind] = window.ITMSModernCompat.toIsoDate(value);
		}
	}

	function parseDisplayDate(value) {
		var parts = String(value || "").split("/");
		if (parts.length === 3) {
			return new Date(Number(parts[2]), Number(parts[1]) - 1, Number(parts[0]));
		}
		return new Date(value);
	}

	function initRoots() {
		window.VouRoot = xmlRoot("VoucherData");
		window.EntryRoot = window.EntryRoot && window.EntryRoot.nodeType ? window.EntryRoot : xmlRoot("EntryData");
		window.iEntryNo = Number(window.iEntryNo || 1);
		window.bVouFlag = Boolean(window.bVouFlag);
		window.bSavFlag = Boolean(window.bSavFlag);
		window.bEditFlag = true;
		window.sTransFlag = window.sTransFlag || "A";
		window.dTransLimit = Number(window.dTransLimit || config.transLimit || 0);
		return { vouRoot: window.VouRoot, entryRoot: window.EntryRoot };
	}

	function clearEntryRoot() {
		var root;
		if (typeof window.clearXML === "function") {
			window.clearXML();
			root = window.EntryRoot || xmlRoot("EntryData");
		} else {
			root = createNode("EntryData", "Entry");
			setAttr(root, "No", window.iEntryNo || 1);
			setAttr(root, "CRDR", "0");
			setAttr(root, "Payto", "");
			setAttr(root, "Amount", "");
			setAttr(root, "AccUnit", "");
			setAttr(root, "AccName", "");
			setAttr(root, "TdsAmount", "0");
			setAttr(root, "TDSElgi", "0");
			setAttr(root, "TdsPercentage", "0");
			setAttr(root, "PayRecAmount", "0");
			window.EntryRoot = root;
		}
		return root;
	}

	function setTdsEligibility(value) {
		var disabled = String(value || "0") !== "1";
		setValue("hTdsElgi", value || "0");
		if (field("txtTdsAmount")) {
			field("txtTdsAmount").disabled = disabled;
		}
		if (field("txtTdsper")) {
			field("txtTdsper").disabled = disabled;
		}
	}

	function appendToEntry(node) {
		var entryRoot = initRoots().entryRoot || clearEntryRoot();
		if (entryRoot && node) {
			entryRoot.appendChild(importFor(entryRoot, node));
		}
	}

	function updateVoucherHeader() {
		var state = initRoots();
		var root = state.vouRoot;
		var select = field("selBook");
		var parts = currentBookParts();
		var approve = checkedValue("optApprove", "Y");
		if (!root) {
			return;
		}
		setAttr(root, "UnitNo", currentOrgId());
		setAttr(root, "UnitName", valueOf("hOrgName", ""));
		setAttr(root, "CRDR", valueOf("hVouCRDR", checkedValue("selCRDR", "")));
		setAttr(root, "VouDate", currentDateValue("ctlDate"));
		setAttr(root, "Approver", approve === "Y" ? "Y" : "N");
		if (select && select.selectedIndex > 0) {
			setAttr(root, "BookNo", parts[0] || selectedValue(select));
			setAttr(root, "BookName", selectedText(select));
			setAttr(root, "BookAcchead", parts[1] || "");
		}
		if (config.bank) {
			setAttr(root, "InstNo", valueOf("txtInstNo", ""));
			setAttr(root, "InstDate", currentDateValue("ctlInsDate"));
			setAttr(root, "PayAt", valueOf("txtPayableAt", ""));
			setAttr(root, "DrawnOn", valueOf("txtDrawnOn", ""));
		}
	}

	function fillEntryFromForm(entryRoot) {
		var narration;
		if (!entryRoot) {
			return;
		}
		setAttr(entryRoot, "No", window.iEntryNo || 1);
		setAttr(entryRoot, "CRDR", checkedValue("selCRDR", valueOf("hVouCRDR", "")));
		setAttr(entryRoot, "Payto", valueOf("txtPayTo", valueOf("txtPayto", "")));
		setAttr(entryRoot, "Amount", valueOf("txtAmount", "0"));
		setAttr(entryRoot, "AccUnit", currentOrgId());
		setAttr(entryRoot, "AccName", valueOf("hOrgName", ""));
		setAttr(entryRoot, "TdsAmount", valueOf("txtTdsAmount", "0"));
		setAttr(entryRoot, "TDSElgi", valueOf("hTdsElgi", "0"));
		setAttr(entryRoot, "TdsPercentage", valueOf("txtTdsper", "0"));
		childElements(entryRoot, "Narration").forEach(function (node) {
			entryRoot.removeChild(node);
		});
		updateAdditionalAmounts(entryRoot);
		narration = createNode("EntryData", "Narration");
		narration.textContent = valueOf("txtNarration", "");
		entryRoot.appendChild(narration);
	}

	function documentAmountField(node, index) {
		var code = attr(node, "No");
		var payNo = trim(attr(node, "PayableNo"));
		return field("txtDocAmount" + code) || field("txtDocAmount" + code + "Z" + payNo + "Z" + (index + 1));
	}

	function updateAdditionalAmounts(entryRoot) {
		childElements(entryRoot).forEach(function (header) {
			childElements(header).forEach(function (node, index) {
				var code = attr(node, "No");
				var groupCode = attr(node, "GroupCode");
				var amountField;
				var ratioField;
				if (header.nodeName === "CostCenter") {
					ratioField = field("txtCCRatio" + code);
					amountField = field("txtCCAmount" + code);
				} else if (header.nodeName === "Analytical") {
					ratioField = field("txtANALRatio" + code + "Z" + groupCode);
					amountField = field("txtANALAmount" + code + "Z" + groupCode);
				} else if (header.nodeName === "PayRec") {
					amountField = documentAmountField(node, index);
				}
				if (ratioField) {
					setAttr(node, "Ratio", ratioField.value);
				}
				if (amountField) {
					setAttr(node, header.nodeName === "PayRec" ? "AmtToAdjust" : "Amount", amountField.value);
				}
			});
		});
	}

	function resetEntryForm() {
		setValue("txtAmount", "0.00");
		setValue("txtNarration", "");
		setValue("txtTdsAmount", "0.00");
		setValue("txtTdsper", "0.00");
		setText("spAccHead", "");
		if (field("selAccHead")) {
			field("selAccHead").selectedIndex = 0;
		}
		if (field("txtPayTo") && window.iEntryNo > 1) {
			field("txtPayTo").readOnly = true;
		}
		if (typeof window.setADDDisplay === "function") {
			window.setADDDisplay(0);
		}
		if (typeof window.setPayableDisplay === "function") {
			window.setPayableDisplay(0);
		}
		clearEntryRoot();
		window.bVouFlag = false;
		window.sTransFlag = "A";
	}

	function renderPayRec(headerNode, fillTotals) {
		var nodes = childElements(headerNode);
		var table = byId("tblPayable");
		var totalTrans = 0;
		var totalAccount = 0;
		var totalAdjusted = 0;
		var narration = [];
		if (!nodes.length || !table) {
			if (typeof window.setPayableDisplay === "function") {
				window.setPayableDisplay(0);
			}
			return;
		}
		if (typeof window.setPayableDisplay === "function") {
			window.setPayableDisplay(1);
		}
		if (typeof window.ClearTable === "function") {
			window.ClearTable("tblPayable", 2, 1);
		} else {
			while (table.rows.length > 2) {
				table.deleteRow(2);
			}
		}
		nodes.forEach(function (node, index) {
			var row = table.insertRow(index + 2);
			var docNo = attr(node, "No");
			var invNo = attr(node, "InvNo");
			var invDate = attr(node, "InvDate");
			var transAmount = toNumber(attr(node, "TransAmount"));
			var adjusted = toNumber(attr(node, "AmtAdjusted"));
			var toAccount = toNumber(attr(node, "AmtToAccount"));
			var toAdjust = attr(node, "AmtToAdjust") || "0";
			totalTrans += transAmount;
			totalAccount += toAccount;
			totalAdjusted += adjusted;
			if (invNo) {
				narration.push(invNo);
			}
			window.InsertCell(row, 1, "", index + 1, "ExcelSerial", "Center", "", 0, 0, 0, 0, "");
			window.InsertCell(row, 1, "", invNo, "ExcelDisplayCell", "", "", 0, 0, 0, 0, "");
			window.InsertCell(row, 1, "", invDate, "ExcelDisplayCell", "", "", 0, 0, 0, 0, "");
			window.InsertCell(row, 1, "", formatNumber(transAmount, 2), "ExcelDisplayCell", "Right", "", 0, 0, 0, 0, "");
			window.InsertCell(row, 1, "", formatNumber(adjusted, 2), "ExcelDisplayCell", "Right", "", 0, 0, 0, 0, "");
			window.InsertCell(row, 1, "", formatNumber(toAccount, 2), "ExcelDisplayCell", "Right", "", 0, 0, 0, 0, "");
			window.InsertCell(row, 2, "txtDocAmount" + docNo, toAdjust, "ExcelInputCell", "", "", 12, 10, 0, 0, "style=\"text-align:right\"");
		});
		if (fillTotals) {
			setValue("txtAmount", formatNumber(totalTrans - totalAccount - totalAdjusted, 2));
			if (narration.length) {
				setValue("txtNarration", narration.join(" "));
			}
		}
	}

	function processPayRec(orgId, partyCode, vouType, pay, rec) {
		var shouldShow = (toNumber(rec) >= 1 && vouType === "D") || (toNumber(pay) >= 1 && vouType === "C");
		if (!shouldShow) {
			if (typeof window.setPayableDisplay === "function") {
				window.setPayableDisplay(0);
			}
			return;
		}
		openDialog("PayRecSelection.asp?VouDate=" + encodeURIComponent(currentDateValue("ctlDate")) + "&orgId=" + encodeURIComponent(orgId) + "&ParCode=" + encodeURIComponent(String(partyCode).replace(/&/g, "and")) + "&Type=" + encodeURIComponent(vouType), "", "", function (node) {
			if (!node || attr(node, 0) !== "1") {
				if (typeof window.setPayableDisplay === "function") {
					window.setPayableDisplay(0);
				}
				return;
			}
			childElements(node).forEach(function (child) {
				appendToEntry(child);
				if (childElements(child).length) {
					renderPayRec(child, true);
				}
			});
		});
	}

	function loadPartyAccount(orgId, partyType, vouType, parType, parSubType, parCode, partyName) {
		var xhr;
		var root;
		var fullPartyCode = "";
		var pay = 0;
		var rec = 0;
		if (!parType || !parSubType || !parCode) {
			return false;
		}
		xhr = getText("XMLGetPayRecCount.asp?orgID=" + encodeURIComponent(orgId) + "&ParSubType=" + encodeURIComponent(parSubType) + "&ParType=" + encodeURIComponent(parType) + "&PartyCode=" + encodeURIComponent(parCode));
		if (trim(xhr.responseText) && typeof window.GetPartyHeadXml === "function") {
			clearXmlChildren("AccHeadData");
			window.GetPartyHeadXml(parCode, partyName || "", xhr.responseText);
		}
		root = xmlRoot("AccHeadData");
		if (!childElements(root).length) {
			setText("spAccHead", "");
			if (field("selAccHead")) {
				field("selAccHead").selectedIndex = 0;
			}
			window.bVouFlag = false;
			if (typeof window.setPayableDisplay === "function") {
				window.setPayableDisplay(0);
			}
			return false;
		}
		clearEntryRoot();
		childElements(root).forEach(function (header) {
			fullPartyCode = partyType + "?" + attr(header, "No");
			pay = attr(header, "Pay");
			rec = attr(header, "Rec");
			setAttr(header, "No", fullPartyCode);
			appendToEntry(header);
			setText("spAccHead", attr(header, "Name"));
			if (window.iEntryNo === 1 && !trim(valueOf("txtPayTo", valueOf("txtPayto", "")))) {
				setValue("txtPayTo", attr(header, "Name"));
				setValue("txtPayto", attr(header, "Name"));
			}
			window.bVouFlag = true;
			window.sTransFlag = "A";
		});
		processPayRec(orgId, fullPartyCode, vouType, pay, rec);
		return true;
	}

	function handlePartyDialog(orgId, partyType, vouType, size, value) {
		var action = trim(attr(value, "Action")).toUpperCase();
		var query = trim(attr(value, "PassQuery"));
		var parType = "";
		var parSubType = "";
		var parCode = "";
		var partyName = "";
		var parts;
		if (action && action !== "DONE" && action !== "CLOSE" && query) {
			openDialog("../../Common/" + size.program + "?" + query, xmlObject("PartyData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (nextValue) {
				handlePartyDialog(orgId, partyType, vouType, size, nextValue);
			});
			return;
		}
		if (action === "CLOSE" || !value) {
			return;
		}
		childElements(value, "Entry").some(function (entry) {
			parType = attr(entry, "RetField3");
			parSubType = attr(entry, "RetField4");
			parCode = attr(entry, "RetField1");
			partyName = attr(entry, "RetField0");
			return true;
		});
		if (!parCode && typeof value === "string") {
			parts = String(value).split(":");
			partyName = parts[0] || "";
			parCode = parts[1] || "";
			parSubType = parts[3] || "";
			parType = parts[4] || "";
		}
		loadPartyAccount(orgId, partyType, vouType, parType, parSubType, parCode, partyName);
	}

	function handleGlDialog(orgId, size, value) {
		var action = trim(attr(value, "Action")).toUpperCase();
		var query = trim(attr(value, "PassQuery"));
		var ret = "";
		var tds = "0";
		var root;
		var accNode;
		if (action && action !== "DONE" && action !== "CLOSE" && query) {
			openDialog("../../Common/" + size.program + "?" + query, xmlObject("TempXMLData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (nextValue) {
				handleGlDialog(orgId, size, nextValue);
			});
			return;
		}
		if (action === "CLOSE" || !value) {
			return;
		}
		childElements(value, "Entry").forEach(function (entry) {
			ret = [0, 1, 2, 3, 4, 5, 6, 7].map(function (index) {
				return attr(entry, "RetField" + index);
			}).join(":");
			tds = attr(entry, "RetField6") || "0";
		});
		if (!ret) {
			return;
		}
		setTdsEligibility(tds);
		if (typeof window.GetGlHeadXml === "function") {
			clearXmlChildren("AccHeadData");
			window.GetGlHeadXml(ret);
		}
		root = xmlRoot("AccHeadData");
		accNode = childElements(root, "AccHead")[0];
		if (!accNode) {
			setText("spAccHead", "");
			if (field("selAccHead")) {
				field("selAccHead").selectedIndex = 0;
			}
			window.bVouFlag = false;
			if (typeof window.setADDDisplay === "function") {
				window.setADDDisplay(0);
			}
			return;
		}
		clearEntryRoot();
		appendToEntry(accNode);
		window.bVouFlag = true;
		window.sTransFlag = attr(accNode, "TransFlag") || attr(accNode, "TransFalg") || "A";
		setText("spAccHead", attr(accNode, "Name"));
		if (window.iEntryNo === 1 && !trim(valueOf("txtPayTo", valueOf("txtPayto", "")))) {
			setValue("txtPayTo", attr(accNode, "Name"));
			setValue("txtPayto", attr(accNode, "Name"));
		}
		if (typeof window.showCCAnal === "function") {
			window.showCCAnal(orgId, attr(accNode, "No"), attr(accNode, "CostCenter"), attr(accNode, "Analytical"));
		}
	}

	function additionalText(entry) {
		var text = [];
		childElements(entry).forEach(function (header) {
			childElements(header).forEach(function (node) {
				if (header.nodeName === "CostCenter" || header.nodeName === "Analytical") {
					text.push([attr(node, "ShortName") || attr(node, 2), attr(node, "Ratio") || attr(node, 3) ? (attr(node, "Ratio") || attr(node, 3)) + "%" : "", attr(node, "Amount") || attr(node, 4)].filter(Boolean).join(" - "));
				}
				if (header.nodeName === "PayRec") {
					text.push([attr(node, "InvNo") || attr(node, 1), attr(node, "InvDate") || attr(node, 2), attr(node, "AmtToAdjust") || attr(node, 5)].filter(Boolean).join(" - "));
				}
			});
		});
		return text.join("<br>");
	}

	function entryAccount(entry) {
		var account = "";
		childElements(entry, "AccHead").forEach(function (node) {
			account = attr(node, "Name") || attr(node, 3);
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

	function appendTdsData() {
		var group = field("SelTDSGrp");
		var tdsRoot = xmlRoot("TDSData");
		var groupValue = group ? selectedValue(group) : "";
		var groupName = group ? selectedValue(group) : "";
		childElements(initRoots().vouRoot, "Entry").forEach(function (entry) {
			setAttr(entry, "TDSFlag", "");
			setAttr(entry, "GroupId", groupValue);
			setAttr(entry, "GroupName", groupName);
			if (tdsRoot) {
				childElements(tdsRoot).forEach(function (node) {
					entry.appendChild(importFor(entry, node));
				});
			}
		});
	}

	function install(userConfig) {
		config = Object.assign(config, userConfig || {});

		window.CheckAccHead = function (root, accountHead) {
			return selectNodes(root, "//AccHead[@No='" + accountHead + "']").length > 0;
		};

		window.popMonBalance = function (value) {
			var parts;
			var bookParts;
			if (field("selBook") && field("selBook").selectedIndex === 0) {
				alert("Select Book ");
				field("selBook").focus();
				return;
			}
			parts = String(value || "").split("~");
			bookParts = currentBookParts();
			openDialog("PopMonBalance.asp?orgid=" + encodeURIComponent(parts[0] || "") + "&Acchead=" + encodeURIComponent(bookParts[1] || "") + "&TillDate=" + encodeURIComponent(parts[2] || ""), "", "dialogHeight:390px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No");
		};

		window.popDayBalance = function (value) {
			var parts;
			var bookParts;
			if (field("selBook") && field("selBook").selectedIndex === 0) {
				alert("Select Book ");
				field("selBook").focus();
				return;
			}
			parts = String(value || "").split("~");
			bookParts = currentBookParts();
			openDialog("PopDayBalance.asp?orgid=" + encodeURIComponent(parts[0] || "") + "&Acchead=" + encodeURIComponent(bookParts[1] || "") + "&TillDate=" + encodeURIComponent(parts[2] || ""), "", "dialogHeight:390px;dialogWidth:620px;center:Yes;help:No;resizable:No;status:No");
		};

		window.popAccHead = function () {
			var select = field("selAccHead");
			var headCount = Number(valueOf("hHeadCount", 0));
			var orgId = currentOrgId();
			var bookNo = valueOf("hBookcode", "");
			var xhr;
			if (!select) {
				return;
			}
			while (headCount > 0 && select.options.length > 1) {
				select.remove(1);
				headCount -= 1;
			}
			setValue("hHeadCount", "0");
			xhr = getText("XMLGetOrgFreqHeads.asp?BkCode=" + encodeURIComponent(config.bookCode) + "&BkNo=" + encodeURIComponent(bookNo) + "&orgID=" + encodeURIComponent(orgId));
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				childElements(xhr.responseXML.documentElement).forEach(function (node, index) {
					select.add(new Option(node.textContent || node.text || "", attr(node, "optValue")), index + 1);
				});
				setValue("hHeadCount", String(childElements(xhr.responseXML.documentElement).length));
			}
		};

		window.selAccountHead = function (select) {
			var option = selectedOption(select || field("selAccHead"));
			var orgId = currentOrgId();
			if (!option || !select || select.selectedIndex <= 0) {
				return;
			}
			if (option.value === "G" || select.selectedIndex === 1) {
				window.showGLHead(orgId);
			} else {
				window.showPartyHead(orgId, option.value + "?" + option.text, valueOf("hVouCRDR", checkedValue("selCRDR", "")));
			}
			if (field("txtNarration")) {
				field("txtNarration").focus();
			}
		};

		window.showPartyHead = function (orgId, partyType, vouType) {
			var size = popupSize("2", "PartySelectionAcc.asp", "500", "500");
			var url = "../../Common/" + size.program + "?orgid=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyType);
			openDialog(url, xmlObject("PartyData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (value) {
				handlePartyDialog(orgId, partyType, vouType, size, value);
			});
		};

		window.showGLHead = function (orgId) {
			var size = popupSize("5", "GLHeadSelection.asp", "500", "500");
			var bookParts = currentBookParts();
			var url = "../../Common/" + size.program + "?orgID=" + encodeURIComponent(orgId) + "&BookId=" + encodeURIComponent(config.bookCode) + "&BookNo=" + encodeURIComponent(bookParts[0] || valueOf("hBookcode", ""));
			openDialog(url, xmlObject("TempXMLData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", function (value) {
				handleGlDialog(orgId, size, value);
			});
		};

		window.checkFileds = function () {
			var amount = valueOf("txtAmount", "");
			var totalAdjust = 0;
			if (!trim(valueOf("txtNarration", ""))) {
				alert("Enter Narration");
				if (field("txtNarration")) {
					field("txtNarration").select();
				}
				return false;
			}
			if (config.bank) {
				if (!trim(valueOf("txtInstNo", ""))) {
					alert("Enter Instrument No");
					field("txtInstNo").focus();
					return false;
				}
				if (!trim(valueOf("txtPayableAt", ""))) {
					alert("Enter PayableAt");
					field("txtPayableAt").focus();
					return false;
				}
				if (!trim(valueOf("txtDrawnOn", ""))) {
					alert("Enter DrawnOn");
					field("txtDrawnOn").focus();
					return false;
				}
			}
			if (typeof window.ValidateAmount === "function" && !window.ValidateAmount(amount)) {
				if (field("txtAmount")) {
					field("txtAmount").select();
				}
				return false;
			}
			if (config.transLimit && toNumber(amount) > toNumber(config.transLimit)) {
				if (window.sTransFlag === "W") {
					alert("Amount is greater than the amount limit");
				} else if (window.sTransFlag === "R") {
					alert("Amount should be less than " + config.transLimit);
					return false;
				}
			}
			childElements(initRoots().entryRoot).forEach(function (header) {
				if (header.nodeName !== "PayRec") {
					return;
				}
				totalAdjust = 0;
				childElements(header).forEach(function (node, index) {
					var amountField = documentAmountField(node, index);
					var adjustLimit = toNumber(attr(node, "TransAmount")) - toNumber(attr(node, "AmtAdjusted")) - toNumber(attr(node, "AmtToAccount"));
					var adjustValue = toNumber(amountField && amountField.value);
					if (adjustValue > adjustLimit) {
						alert("\"To Adjust Amount\" should be less than \"Document Amount-(Adjusted +To Account)\"");
						if (amountField) {
							amountField.focus();
						}
						totalAdjust = Number.MAX_VALUE;
					} else {
						totalAdjust += adjustValue;
					}
				});
			});
			if (totalAdjust === Number.MAX_VALUE) {
				return false;
			}
			if (totalAdjust > toNumber(amount)) {
				alert("Total of \"To Adjust Amount\" should be less than \"Voucher Amount\"");
				return false;
			}
			return true;
		};

		window.AddEntry = function (flag) {
			var state = initRoots();
			var entryRoot = state.entryRoot || clearEntryRoot();
			var entries = childElements(state.vouRoot, "Entry");
			updateVoucherHeader();
			if (flag === "S" && !window.bVouFlag && entries.length) {
				window.SaveXML();
				return;
			}
			if (field("selBook") && field("selBook").selectedIndex === 0) {
				alert("Select Book");
				field("selBook").focus();
				return;
			}
			if (!window.bVouFlag) {
				return;
			}
			if (!window.checkFileds()) {
				return;
			}
			window.bSavFlag = true;
			fillEntryFromForm(entryRoot);
			state.vouRoot.appendChild(importFor(state.vouRoot, entryRoot));
			if (flag === "A") {
				window.DisplayVoucher();
				window.iEntryNo += 1;
				setText("spEntryNo", String(window.iEntryNo));
				resetEntryForm();
			} else {
				window.SaveXML();
			}
		};

		window.DisplayVoucher = function () {
			var state = initRoots();
			var table = byId("tblVoucher");
			var total = 0;
			var row;
			var serial = 0;
			if (!table) {
				return;
			}
			if (byId("DisVoucher")) {
				byId("DisVoucher").style.height = "200px";
				byId("DisVoucher").style.visibility = "visible";
			}
			if (typeof window.ClearTable === "function") {
				window.ClearTable("tblVoucher", 1, 1);
			} else {
				while (table.rows.length > 1) {
					table.deleteRow(1);
				}
			}
			childElements(state.vouRoot, "Entry").forEach(function (entry, index) {
				var amount = toNumber(attr(entry, "Amount"));
				var crdr = attr(entry, "CRDR");
				serial = index + 1;
				setAttr(entry, "No", serial);
				total += crdr === "C" ? -amount : amount;
				row = table.insertRow(table.rows.length);
				window.InsertCell(row, 1, "", serial, "ExcelSerial", "Center", "top", 0, 0, 0, 0, "");
				window.InsertCell(row, 1, "", attr(entry, "AccName"), "ExcelDisplayCell", "left", "top", 0, 0, 0, 0, "");
				window.InsertCell(row, 1, "", entryAccount(entry), "ExcelDisplayCell", "left", "top", 0, 0, 0, 0, "");
				window.InsertCell(row, 1, "", entryNarration(entry), "ExcelDisplayCell", "left", "top", 0, 0, 0, 0, "");
				window.InsertCell(row, 1, "", formatNumber(amount, 2) + "&nbsp;" + crdr + "r", "ExcelDisplayCell", "right", "top", 0, 0, 0, 0, "");
				window.InsertCell(row, 1, "", additionalText(entry), "ExcelDisplayCell", "left", "top", 0, 0, 0, 0, "");
				window.InsertCell(row, 1, "", formatNumber(attr(entry, "TdsAmount"), 2), "ExcelDisplayCell", "right", "top", 0, 0, 0, 0, "");
				window.InsertCell(row, 1, "", formatNumber(attr(entry, "TdsPercentage"), 2), "ExcelDisplayCell", "right", "top", 0, 0, 0, 0, "");
			});
			row = table.insertRow(table.rows.length);
			window.InsertCell(row, 1, "", "<b>Total</b>", "ExcelDisplayCell", "right", "top", 0, 0, 4, 0, "");
			window.InsertCell(row, 1, "", "Rs. &nbsp;" + formatNumber(Math.abs(total), 2) + (total < 0 ? "&nbsp;Cr" : "&nbsp;Dr"), "ExcelDisplayCell", "right", "top", 0, 0, 0, 0, "");
			window.InsertCell(row, 1, "", "", "ExcelDisplayCell", "right", "top", 0, 0, 0, 0, "");
			window.InsertCell(row, 1, "", "", "ExcelDisplayCell", "right", "top", 0, 0, 0, 0, "");
			window.InsertCell(row, 1, "", "", "ExcelDisplayCell", "right", "top", 0, 0, 0, 0, "");
		};

		window.SaveXML = function () {
			var xhr;
			var moduleCode = valueOf("hVouName", config.moduleCode);
			updateVoucherHeader();
			if (config.enableTds) {
				appendTdsData();
			}
			if (!window.bSavFlag) {
				return;
			}
			xhr = postXml("XMLSave.asp?Name=" + encodeURIComponent(config.saveName) + "&Mod=" + encodeURIComponent(moduleCode), xmlString(xmlDocument("VoucherData")));
			if (trim(xhr.responseText)) {
				alert(xhr.responseText);
				return;
			}
			if (field("B12")) {
				field("B12").disabled = true;
			}
			form().submit();
		};

		window.CancelAction = function (page) {
			form().action = page;
			form().submit();
		};

		window.SelMisParty = function () {
			var size = popupSize("4", "MisPartySelection.asp", "480", "420");
			function handle(value) {
				var parts = String(value || "").split(":");
				if (String(value) === "AN") {
					window.AddNewParty();
					return;
				}
				if (parts.length <= 1 && trim(value)) {
					openDialog(size.program + "?" + value, "", "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;center:Yes;help:No;resizable:No;status:No", handle);
					return;
				}
				if (parts.length > 1) {
					setValue("txtPayTo", parts[0]);
				}
			}
			if (field("txtPayTo") && field("txtPayTo").readOnly) {
				return;
			}
			openDialog(size.program, "", "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;center:Yes;help:No;resizable:No;status:No", handle);
		};

		window.AddNewParty = function () {
			openDialog("MisParCreate.asp?", "", "dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (value) {
				setValue("txtPayTo", value || "");
			});
		};

		window.showReceiptpopup = function (grnDate, receiptCode, against, route, invCode, receipt, itemType) {
			openDialog("../../Purchase/Transaction/RepActualReceiptDetailspopup.asp?ItemType=" + encodeURIComponent(itemType || "") + "&iGrnDt=" + encodeURIComponent(grnDate || "") + "&iRcptCode=" + encodeURIComponent(receiptCode || "") + "&sGRNAgainstStr=" + encodeURIComponent(against || "") + "&sReceiptRouteStr=" + encodeURIComponent(route || "") + "&sInvCode=" + encodeURIComponent(invCode || "") + "&sReceiptCode=" + encodeURIComponent(receipt || ""), "", "dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No", function () {});
		};

		window.ViewInvoiceDetailspopup = function (invNo, invCode) {
			openDialog("../../Purchase/Transaction/RepPurInvoiceDetailspopup.asp?iInvNo=" + encodeURIComponent(invNo || "") + "&sInvCode=" + encodeURIComponent(invCode || ""), "A", "dialogHeight:600px;dialogWidth:800px;center:Yes;help:No;resizable:No;status:No", function () {});
		};

		window.ResetList = function (item) {
			if (field("selUserId")) {
				field("selUserId").disabled = item && item.value !== "Y";
			}
		};

		window.CheckNoSerForMis = function () {
			var bookParts = currentBookParts();
			var passValue = [
				currentOrgId(),
				valueOf("hVouCode", ""),
				valueOf("hCallFrm", ""),
				valueOf("hVouCRDR", "D") || "D",
				bookParts[0] || "",
				currentDateValue("ctlDate")
			].join(":");
			var xhr = getText("NoSeriesCheck.asp?sValue=" + encodeURIComponent(passValue));
			if (trim(xhr.responseText) === "T") {
				return true;
			}
			if (trim(xhr.responseText) === "F") {
				alert("No Series is Not Defined ");
				return false;
			}
			alert("Error ");
			return false;
		};

		window.DispBook = function (select) {
			var bookSelect = field("selBook");
			var xhr;
			if (!select || select.selectedIndex === 0 || !bookSelect) {
				return;
			}
			xhr = getText("XMLGetOrgBook.asp?BkCode=" + encodeURIComponent(config.bookCode) + "&orgID=" + encodeURIComponent(select.value));
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				bookSelect.options.length = 1;
				childElements(xhr.responseXML.documentElement).forEach(function (node) {
					bookSelect.add(new Option(attr(node, 1), attr(node, 0)));
				});
			}
		};

		window.TDSAmount = function () {
			var group = valueOf("SelTDSGrp", "0");
			var entryNo = textOf("spEntryNo") || String(window.iEntryNo || 1);
			var amount = valueOf("txtAmount", "0");
			var xhr;
			var tdsRoot;
			var total = 0;
			setValue("hTdsNew", "Y");
			if (!config.enableTds || group === "0") {
				return;
			}
			xhr = getText("TDSCalcCash.asp?EntNo=" + encodeURIComponent(entryNo) + "&Amount=" + encodeURIComponent(amount) + "&GrpId=" + encodeURIComponent(group));
			tdsRoot = xmlRoot("TDSData");
			childElements(tdsRoot, "TDS").forEach(function (node) {
				tdsRoot.removeChild(node);
			});
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				childElements(xhr.responseXML.documentElement, "TDS").forEach(function (node) {
					total += toNumber(attr(node, "PayRecAmount"));
					tdsRoot.appendChild(importFor(tdsRoot, node));
				});
			}
			setValue("txtTdsAmount", formatNumber(total, 2));
			postXml("XMLSaveForTDS.asp?Name=TDS_Cash", xmlString(xmlDocument("TDSData")));
		};

		window.TDSCalc = function () {
			var group = valueOf("SelTDSGrp", "0");
			var entryNo = textOf("spEntryNo") || String(window.iEntryNo || 1);
			var amount = valueOf("txtAmount", "0");
			var newAmount = valueOf("hTdsAmt", "");
			var url;
			if (!config.enableTds) {
				return;
			}
			if (trim(group) === "0") {
				alert("select TDS Group");
				return;
			}
			url = "TDSGroupSelectionCash.asp?EntNo=" + encodeURIComponent(entryNo) + "&Amount=" + encodeURIComponent(amount) + "&NewAmt=" + encodeURIComponent(newAmount) + "&GrpId=" + encodeURIComponent(group) + "&CallFrom=" + encodeURIComponent(valueOf("hCallFrm", "")) + "&VouName=" + encodeURIComponent(valueOf("hVouName", "")) + "&NewVal=" + encodeURIComponent(valueOf("hTdsNew", "")) + "&Update=" + encodeURIComponent(valueOf("hUpdate", ""));
			openDialog(url, xmlObject("TDSData"), "dialogHeight:350px;dialogWidth:380px;center:Yes;status:no", function (value) {
				var root = xmlRoot("TDSData");
				var total = 0;
				childElements(root, "TDS").forEach(function (node) {
					root.removeChild(node);
				});
				childElements(value, "TDS").forEach(function (node) {
					total += toNumber(attr(node, "PayRecAmount"));
					root.appendChild(importFor(root, node));
				});
				setValue("txtTdsAmount", formatNumber(total, 2));
				postXml("XMLSaveForTDS.asp?Name=TDS_Cash", xmlString(xmlDocument("TDSData")));
			});
		};

		window.AmtFun = function () {
			if (trim(valueOf("SelTDSGrp", "0")) !== "0") {
				window.TDSAmount();
			}
		};

		window.Init = function () {
			var finFrom;
			var finTo;
			var today = new Date();
			ensureCompat();
			initRoots();
			finFrom = valueOf("hFinFrom", "");
			finTo = valueOf("hFinTo", "");
			setDateLimit("ctlDate", "min", finFrom);
			if (finTo && parseDisplayDate(finTo) < today) {
				setDateLimit("ctlDate", "max", finTo);
				setDateValue("ctlDate", finTo);
			} else {
				setDateLimit("ctlDate", "max", today);
				setDateValue("ctlDate", today);
			}
			if (config.bank) {
				setValue("txtInstNo", valueOf("hChequeNo", ""));
				setDateValue("ctlInsDate", valueOf("hChequeDate", ""));
			}
		};
		window.init = window.Init;

		window.setParty = function () {
			var parType = valueOf("hParType", "");
			var parSubType = valueOf("hParSubType", "");
			var parCode = valueOf("hParCode", "");
			var orgId = currentOrgId();
			var select = field("selAccHead");
			var partyType = "";
			var partyName = "";
			var xhr;
			if (!parType || !parSubType || !parCode || !select || !select.options) {
				return;
			}
			for (var i = 0; i < select.options.length; i += 1) {
				if (trim(parType + "?" + parSubType) === trim(select.options[i].value)) {
					select.selectedIndex = i;
					partyType = select.options[i].value + "?" + select.options[i].text;
					break;
				}
			}
			xhr = getText("../../Include/GetPartyName.asp?ParCode=" + encodeURIComponent(parCode));
			partyName = trim(xhr.responseText || "");
			loadPartyAccount(orgId, partyType, valueOf("hVouCRDR", checkedValue("selCRDR", "")), parType, parSubType, parCode, partyName);
		};
		window.setparty = window.setParty;
	};

	window.ITMSMiscVoucherEntry = {
		install: install
	};
})(window, document);
