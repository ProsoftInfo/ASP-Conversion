(function (window, document) {
	"use strict";

	var config = {
		bookCode: "01",
		moduleCode: "CA",
		journal: false,
		bank: false,
		payRecPage: "PayRecSelectionWithAllAdj.asp"
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
		if ("innerText" in element) {
			element.innerText = value == null ? "" : String(value);
		} else {
			element.textContent = value == null ? "" : String(value);
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
		if (window.ITMSModernCompat && window.ITMSModernCompat.upgradeXmlIslands) {
			window.ITMSModernCompat.upgradeXmlIslands(document);
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

	function createNode(xmlName, nodeName) {
		var object = xmlObject(xmlName);
		if (object && typeof object.createElement === "function") {
			return object.createElement(nodeName);
		}
		return document.implementation.createDocument("", "", null).createElement(nodeName);
	}

	function resetXmlRoot(xmlName, rootName) {
		var object = xmlObject(xmlName);
		if (object && typeof object.loadXML === "function") {
			object.loadXML("<" + rootName + "/>");
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
		if (node) {
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

	function createHttp() {
		return new XMLHttpRequest();
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
		if (!window.ITMSModernCompat || !window.ITMSModernCompat.openModalDialog) {
			alert("Modern browser compatibility script is still loading. Please try again.");
			return;
		}
		window.ITMSModernCompat.openModalDialog(url, args || "", features || "", callback || function () {});
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

	function initState() {
		var vouRoot = xmlRoot("VoucherData");
		var entryRoot = window.EntryRoot && window.EntryRoot.nodeType ? window.EntryRoot : xmlRoot("EntryData");
		window.VouRoot = vouRoot;
		window.EntryRoot = entryRoot;
		window.iEntryNo = Number(window.iEntryNo || valueOf("hEntryNo", 0) || selectNodes(vouRoot, "//Entry").length || 0);
		if (!window.bEditFlag && window.bEditFlag !== false) {
			window.bEditFlag = true;
		}
		return { vouRoot: vouRoot, entryRoot: entryRoot };
	}

	function currentBookParts() {
		return String(selectedValue(field("selBook")) || "").split("-");
	}

	function currentOrgId() {
		return valueOf("hOrgId", valueOf("hOrgID", valueOf("hAccUnitId", "")));
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
			id: currentOrgId(),
			name: valueOf("hOrgName", "")
		};
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
		return control.value || "";
	}

	function setCurrentDate(controlName, value) {
		var control = field(controlName || "ctlDate") || byId(controlName || "ctlDate");
		if (!control) {
			return;
		}
		if (typeof control.setDate === "function") {
			control.setDate(value);
		} else if (typeof control.SetDate === "function") {
			control.SetDate(value);
		} else {
			control.value = value || "";
		}
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

	function setCRDR(value) {
		var crdr = field("selCRDR");
		if (crdr && crdr.length != null) {
			for (var i = 0; i < crdr.length; i += 1) {
				crdr[i].checked = crdr[i].value === value;
			}
		}
	}

	function clearEntryRoot() {
		var root = createNode("EntryData", "Entry");
		setAttr(root, "No", "");
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
		return root;
	}

	function clearAccHeadData() {
		resetXmlRoot("AccHeadData", "account");
	}

	function appendEntryChild(node) {
		var state = initState();
		if (state.entryRoot && node) {
			state.entryRoot.appendChild(importFor(state.entryRoot, node));
		}
	}

	function setSelectedHeadDisplay(text) {
		setText("spAccHead", text || "");
	}

	function updateTdsEligibility(value) {
		setValue("hTdsElgi", value || "0");
		var amount = field("txtTdsAmount");
		var percent = field("txtTdsper");
		var disabled = String(value || "0") === "0";
		if (amount) {
			amount.disabled = disabled;
		}
		if (percent) {
			percent.disabled = disabled;
		}
	}

	function addAccountHead(no, costCenter, analytical, name, type, transFlag, pay, rec, adv) {
		var entryRoot = initState().entryRoot;
		var node = createNode("EntryData", "AccHead");
		setAttr(node, "No", trim(no));
		setAttr(node, "CostCenter", trim(costCenter || "0"));
		setAttr(node, "Analytical", trim(analytical || "0"));
		setAttr(node, "Name", name || "");
		setAttr(node, "Type", type || "G");
		setAttr(node, "TransFalg", trim(transFlag || "A"));
		setAttr(node, "TransFlag", trim(transFlag || "A"));
		if (pay != null) {
			setAttr(node, "Pay", pay);
		}
		if (rec != null) {
			setAttr(node, "Rec", rec);
		}
		if (adv != null) {
			setAttr(node, "Adv", adv);
		}
		if (entryRoot) {
			entryRoot.appendChild(node);
		}
		window.bVouFlag = true;
		window.sTransFlag = trim(transFlag || "A");
		return node;
	}

	function selectedBookAccHead() {
		var parts = currentBookParts();
		return parts[1] || valueOf("hBookAccHead", attr(xmlRoot("VoucherData"), "BookAcchead"));
	}

	function makeDisplayVisible() {
		var dis = byId("DisVoucher");
		if (dis && dis.style) {
			dis.style.visibility = "visible";
			dis.style.height = "auto";
		}
	}

	function insertCell(row, type, name, value, cls, align, valign, size, max, colspan, rowspan, options) {
		if (typeof window.InsertCell === "function") {
			return window.InsertCell(row, type, name, value, cls, align, valign, size || 0, max || 0, colspan || 0, rowspan || 0, options || "");
		}
		var cell = row.insertCell();
		cell.className = cls || "";
		if (align) {
			cell.align = align;
		}
		cell.innerHTML = value == null ? "" : String(value);
		return cell;
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

	function entryAccountText(entry) {
		var account = "";
		childElements(entry).forEach(function (child) {
			if (child.nodeName === "AccHead") {
				account = attr(child, "Type") === "P" ? attr(child, "Name") : (attr(child, "No") + " - " + attr(child, "Name"));
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
		var text = [];
		childElements(entry).forEach(function (header) {
			if (header.nodeName === "CostCenter" || header.nodeName === "Analytical" || header.nodeName === "PayRec") {
				childElements(header).forEach(function (node) {
					text.push([attr(node, "ShortName") || attr(node, "InvNo") || attr(node, "No"), attr(node, "Ratio"), attr(node, "Amount") || attr(node, "AmtToAdjust")].filter(Boolean).join(" - "));
				});
			}
		});
		return text.join("<br>");
	}

	function selectBookByValue(value) {
		var select = field("selBook");
		if (!select || !select.options) {
			return;
		}
		for (var i = 0; i < select.options.length; i += 1) {
			if (String(select.options[i].value) === String(value)) {
				select.selectedIndex = i;
				return;
			}
		}
	}

	function firstSelectedBook() {
		var select = field("selBook");
		if (select && select.options && select.options.length > 1 && select.selectedIndex <= 0) {
			select.selectedIndex = 1;
		}
	}

	function populateSelectFromXml(select, root) {
		childElements(root).forEach(function (node) {
			var option = new Option(attr(node, 1) || node.textContent || node.text || "", attr(node, 0));
			select.add(option);
		});
	}

	function applyPartyHeadSelection(orgId, partyType, vouType, selected) {
		var parts = String(selected || "").split(":");
		if (parts.length <= 1) {
			return false;
		}
		var partyName = parts[0] || "";
		var partyCode = parts[1] || "";
		var partySubType = parts[3] || "";
		var partyMainType = parts[4] || "";
		var xhr = getText("XMLGetPayRecCount.asp?orgID=" + encodeURIComponent(orgId) + "&ParSubType=" + encodeURIComponent(partySubType) + "&ParType=" + encodeURIComponent(partyMainType) + "&PartyCode=" + encodeURIComponent(partyCode));
		if (trim(xhr.responseText) && typeof window.GetPartyHeadXml === "function") {
			clearAccHeadData();
			window.GetPartyHeadXml(partyCode, partyName, xhr.responseText);
		}
		childElements(xmlRoot("AccHeadData")).forEach(function (header) {
			var fullPartyCode = partyType + "?" + attr(header, 0);
			setAttr(header, "No", fullPartyCode);
			addAccountHead(fullPartyCode, "0", "0", attr(header, "Name") || partyName, "P", "A", attr(header, "Pay"), attr(header, "Rec"), attr(header, "Adv"));
			setSelectedHeadDisplay(attr(header, "Name") || partyName);
			if (!trim(valueOf("txtPayTo", valueOf("txtPayto", "")))) {
				setValue("txtPayTo", attr(header, "Name") || partyName);
				setValue("txtPayto", attr(header, "Name") || partyName);
			}
		});
		return true;
	}

	function processPayRecIfNeeded(orgId, partyCode, vouType, pay, rec, adv) {
		var shouldShow = toNumber(pay) >= 1 || toNumber(rec) >= 1 || toNumber(adv) >= 1;
		if (!shouldShow) {
			if (typeof window.setPayableDisplay === "function") {
				window.setPayableDisplay(0);
			}
			return;
		}
		openDialog(config.payRecPage + "?orgId=" + encodeURIComponent(orgId) + "&ParCode=" + encodeURIComponent(String(partyCode).replace(/&/g, "and")) + "&Type=" + encodeURIComponent(vouType), "", "dialogHeight:500px;dialogWidth:550px;center:Yes;help:No;resizable:No;status:No", function (node) {
			var payRec;
			if (!node || attr(node, 0) !== "1") {
				if (typeof window.setPayableDisplay === "function") {
					window.setPayableDisplay(0);
				}
				return;
			}
			payRec = importFor(initState().entryRoot, node);
			appendEntryChild(payRec);
			if (typeof window.popPayRec === "function") {
				window.popPayRec(payRec);
			}
		});
	}

	function updateEntryAmountsFromInputs(entryRoot) {
		childElements(entryRoot).forEach(function (header) {
			childElements(header).forEach(function (node, index) {
				var code = attr(node, "No");
				var group = attr(node, "GroupCode");
				var payNo = attr(node, "PayableNo");
				var amountField;
				var ratioField;
				if (header.nodeName === "CostCenter") {
					ratioField = field("txtCCRatio" + code);
					amountField = field("txtCCAmount" + code);
				} else if (header.nodeName === "Analytical") {
					ratioField = field("txtANALRatio" + code + "Z" + group);
					amountField = field("txtANALAmount" + code + "Z" + group);
				} else if (header.nodeName === "PayRec") {
					amountField = field("txtDocAmount" + code + "Z" + payNo + "Z" + (index + 1));
				}
				if (ratioField) {
					setAttr(node, "Ratio", ratioField.value);
				}
				if (amountField) {
					if (header.nodeName === "PayRec") {
						setAttr(node, "AmtToAdjust", amountField.value);
					} else {
						setAttr(node, "Amount", amountField.value);
					}
				}
			});
		});
	}

	function fillEntryFromForm(entryRoot) {
		var narration;
		var crdr = checkedCRDR();
		var amount = valueOf("txtAmount", "0");
		var unit = currentAccountingUnit();
		setAttr(entryRoot, "CRDR", crdr);
		setAttr(entryRoot, "Payto", valueOf("txtPayTo", valueOf("txtPayto", valueOf("hPayTo", ""))));
		setAttr(entryRoot, "Amount", amount);
		setAttr(entryRoot, "AccUnit", unit.id);
		setAttr(entryRoot, "AccName", unit.name);
		setAttr(entryRoot, "TdsAmount", valueOf("txtTdsAmount", "0"));
		setAttr(entryRoot, "TDSElgi", valueOf("hTdsElgi", "0"));
		setAttr(entryRoot, "TdsPercentage", valueOf("txtTdsper", "0"));
		updateEntryAmountsFromInputs(entryRoot);
		childElements(entryRoot, "Narration").forEach(function (node) {
			entryRoot.removeChild(node);
		});
		narration = createNode("EntryData", "Narration");
		narration.textContent = valueOf("txtNarration", "");
		entryRoot.appendChild(narration);
	}

	function resetEntryForm() {
		setValue("txtAmount", "0.00");
		setValue("txtNarration", "");
		setValue("txtTdsAmount", "0.00");
		setValue("txtTdsper", "0.00");
		setSelectedHeadDisplay("");
		if (field("selAccHead")) {
			field("selAccHead").selectedIndex = 0;
		}
		if (typeof window.setADDDisplay === "function") {
			window.setADDDisplay(0);
		}
		if (typeof window.setPayableDisplay === "function") {
			window.setPayableDisplay(0);
		}
		clearEntryRoot();
	}

	function doSaveXml() {
		var state = initState();
		var saveUrl;
		var actionUrl;
		var xhr;
		if (!window.bSavFlag) {
			return;
		}
		if (typeof window.CheckApp === "function" && !window.CheckApp()) {
			return;
		}
		if (!config.journal && typeof window.CheckContraEnt === "function" && !window.CheckContraEnt()) {
			return;
		}
		if (field("selBook") && field("selBook").selectedIndex === 0) {
			alert("Select Book");
			return;
		}
		if (typeof window.CheckFinDate === "function" && !window.CheckFinDate()) {
			return;
		}
		window.UpdateXML();
		if (config.bank && trim(valueOf("hInsrAmt", "0")) !== "" && trim(valueOf("hInsrAmt", "0")) !== "0" && trim(window.CheckVouAmount()) !== trim(valueOf("hInsrAmt", "0"))) {
			alert("Voucher Amount and Instrument Amount should be same");
			return;
		}
		if (config.journal) {
			saveUrl = trim(valueOf("txtVouNo", "")) === "" ? "XMLSave.asp?Name=Voucher Entry&Mod=GJ" : "XMLSave.asp?Name=Voucher AMD&Mod=GJ";
			actionUrl = trim(valueOf("txtVouNo", "")) === "" ? "VouGenerate.asp" : "VouAmdGenerate.asp";
		} else {
			saveUrl = valueOf("hAmendTy", "N") === "N" ? "XMLSave.asp?Name=Voucher Entry&Mod=" + config.moduleCode : "XMLSave.asp?Name=Voucher AMD&Mod=" + config.moduleCode;
			actionUrl = valueOf("hAmendTy", "N") === "N" ? "VouGenerate.asp" : "VouAmdGenerate.asp";
		}
		xhr = postXml(saveUrl, xmlString(xmlDocument("VoucherData")));
		if (trim(xhr.responseText) !== "") {
			alert(xhr.responseText);
			return;
		}
		if (field("btnNext")) {
			field("btnNext").disabled = true;
		}
		if (field("btnAdd")) {
			field("btnAdd").disabled = true;
		}
		form().action = actionUrl;
		form().submit();
	}

	function install(userConfig) {
		config = Object.assign(config, userConfig || {});

		window.DisplayBook = function (objUnit) {
			var select = field("selBook");
			var bookCode = valueOf("hBookCode", valueOf("hBookcode", ""));
			var unitId;
			var xhr;
			if (!select) {
				return;
			}
			unitId = objUnit && objUnit.options ? selectedValue(objUnit) : currentOrgId();
			if (trim(unitId) === "" || trim(unitId) === "0") {
				unitId = currentOrgId();
			}
			select.options.length = 1;
			xhr = getText("XMLGetOrgBook.asp?BkCode=" + encodeURIComponent(config.bookCode) + "&orgID=" + encodeURIComponent(unitId));
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				var bookData = xmlObject("UnitBookData");
				if (bookData && typeof bookData.loadXML === "function") {
					bookData.loadXML(xmlString(xhr.responseXML));
				}
				populateSelectFromXml(select, xmlRoot("UnitBookData"));
			}
			firstSelectedBook();
			if (trim(bookCode) && field("hBookAccHead")) {
				selectBookByValue(trim(bookCode) + "-" + trim(field("hBookAccHead").value));
			}
			window.SetBookAccHead();
		};

		window.popAccHead = function () {
			var select = field("selAccHead");
			var headCount = Number(valueOf("hHeadCount", 0));
			var unit = currentAccountingUnit().id;
			var xhr;
			if (!select) {
				return;
			}
			while (headCount > 0 && select.options.length > 1) {
				select.remove(1);
				headCount -= 1;
			}
			setValue("hHeadCount", "0");
			xhr = getText("XMLGetOrgParType.asp?orgID=" + encodeURIComponent(unit));
			if (xhr.responseXML && xhr.responseXML.documentElement) {
				var outData = xmlObject("OutData");
				if (outData && typeof outData.loadXML === "function") {
					outData.loadXML(xmlString(xhr.responseXML));
				}
				populateSelectFromXml(select, xmlRoot("OutData"));
			}
		};

		window.selAccountHead = function (objAcc) {
			var select = objAcc || field("selAccHead");
			var option = selectedOption(select);
			var headCount = Number(valueOf("hHeadCount", 0));
			var orgId = currentAccountingUnit().id;
			var parts;
			if (!select || select.selectedIndex <= 0 || !option) {
				return;
			}
			parts = String(option.value || "").split("?");
			if (parts.length >= 5 && (select.selectedIndex <= headCount || option.value !== "G")) {
				updateTdsEligibility(parts[4]);
				clearEntryRoot();
				addAccountHead(parts[0], parts[1], parts[2], option.text, "G", parts[3]);
				setSelectedHeadDisplay(option.text);
				setValue("txtPayTo", valueOf("hPayTo", valueOf("hPayto", "")));
				window.showCCAnal(orgId, trim(parts[0]), trim(parts[1]), trim(parts[2]));
			} else if (option.value === "G" || select.selectedIndex === headCount + 1) {
				window.showGLHead(orgId);
			} else {
				window.showPartyHead(orgId, option.value + "?" + option.text, valueOf("hVouCRDR", checkedCRDR()));
			}
			if (field("txtNarration")) {
				field("txtNarration").focus();
			}
		};

		window.showGLHead = function (orgId) {
			var size = popupSize("5", "GLHeadSelection.asp", "500", "350");
			var bookNo = valueOf("hBookcode", "");
			var url = "../../Common/" + size.program + "?hSelectMode=R&hBal=Y&orgId=" + encodeURIComponent(orgId) + "&BookId=01&BookNo=" + encodeURIComponent(bookNo) + "&AccHead=" + encodeURIComponent(selectedBookAccHead());
			function handle(outValue) {
				var ret = "";
				var tds = "0";
				var accNode;
				if (!outValue || trim(attr(outValue, "Action")).toUpperCase() === "CLOSE") {
					return;
				}
				if (trim(attr(outValue, "Action")).toUpperCase() && trim(attr(outValue, "Action")).toUpperCase() !== "DONE") {
					openDialog("../../Common/" + size.program + "?" + attr(outValue, "PassQuery"), xmlObject("GLHeadData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", handle);
					return;
				}
				childElements(outValue).forEach(function (child) {
					tds = attr(child, "RetField6");
					ret = [0, 1, 2, 3, 4, 5, 6, 7].map(function (index) {
						return attr(child, "RetField" + index);
					}).join(":");
				});
				if (!ret) {
					return;
				}
				updateTdsEligibility(tds);
				clearAccHeadData();
				if (typeof window.GetGlHeadXml === "function") {
					window.GetGlHeadXml(ret);
				}
				accNode = childElements(xmlRoot("AccHeadData"), "AccHead")[0];
				if (accNode) {
					clearEntryRoot();
					addAccountHead(attr(accNode, "No"), attr(accNode, "CostCenter"), attr(accNode, "Analytical"), attr(accNode, "Name"), "G", attr(accNode, "TransFlag") || attr(accNode, "TransFalg"));
					setSelectedHeadDisplay(attr(accNode, "Name"));
					window.showCCAnal(orgId, attr(accNode, "No"), attr(accNode, "CostCenter"), attr(accNode, "Analytical"));
				}
			}
			openDialog(url, xmlObject("GLHeadData"), "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", handle);
		};

		window.showPartyHead = function (orgId, partyType, vouType) {
			var size = popupSize("12", "PartySelectionAcc.asp", "500", "500");
			var url = "../../Common/" + size.program + "?orgID=" + encodeURIComponent(orgId) + "&Party=" + encodeURIComponent(partyType);
			function handle(value) {
				var selected = String(value || "");
				var parts = selected.split(":");
				var accNode;
				if (parts.length <= 1 && selected !== "") {
					openDialog("../../Common/" + size.program + "?" + selected, "", "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", handle);
					return;
				}
				if (!applyPartyHeadSelection(orgId, partyType, vouType, selected)) {
					if (field("selAccHead")) {
						field("selAccHead").selectedIndex = 0;
					}
					return;
				}
				accNode = childElements(initState().entryRoot, "AccHead")[0];
				if (accNode) {
					processPayRecIfNeeded(orgId, attr(accNode, "No"), vouType, attr(accNode, "Pay"), attr(accNode, "Rec"), attr(accNode, "Adv"));
				}
			}
			openDialog(url, "", "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", handle);
		};

		window.checkFileds = function () {
			var amount = valueOf("txtAmount", "");
			if (trim(valueOf("txtNarration", "")) === "") {
				alert("Enter Narration");
				if (field("txtNarration")) {
					field("txtNarration").select();
				}
				return false;
			}
			if (typeof window.ValidateAmount === "function" && !window.ValidateAmount(amount)) {
				if (field("txtAmount")) {
					field("txtAmount").select();
				}
				return false;
			}
			return true;
		};

		window.AddNew = function () {
			window.AddEntry(valueOf("hAction", "New") === "Edit" ? "U" : "A");
		};

		window.AddEntry = function (flag) {
			var state = initState();
			var entryRoot = state.entryRoot || clearEntryRoot();
			var saveAfter = flag === "S";
			if (window.bVouFlag || trim(valueOf("txtAmount", "")) !== "" && trim(valueOf("txtAmount", "")) !== "0.00") {
				if (!window.checkFileds()) {
					return;
				}
				if (flag !== "U") {
					window.iEntryNo = Number(window.iEntryNo || 0) + 1;
				}
				setAttr(entryRoot, "No", window.iEntryNo);
				fillEntryFromForm(entryRoot);
				state.vouRoot.appendChild(importFor(state.vouRoot, entryRoot));
				window.bSavFlag = true;
				setValue("hEntryNo", String(window.iEntryNo + 1));
				window.DisplayVoucher("0");
				resetEntryForm();
				setText("spEntryNo", String(window.iEntryNo + 1));
			}
			if (saveAfter) {
				if (typeof window.CheckVouStat === "function" && !window.CheckVouStat()) {
					return;
				}
				doSaveXml();
			}
		};

		window.EditEntry = function (entryNo, editType) {
			var state = initState();
			var entries = childElements(state.vouRoot, "Entry");
			entries.some(function (entry) {
				if (String(attr(entry, "No")) !== String(entryNo)) {
					return false;
				}
				state.vouRoot.removeChild(entry);
				window.EntryRoot = entry;
				window.bVouFlag = true;
				window.bEditFlag = false;
				setText("spEntryNo", entryNo);
				setValue("txtAmount", attr(entry, "Amount"));
				setValue("txtPayTo", attr(entry, "Payto"));
				setValue("txtTdsAmount", attr(entry, "TdsAmount") || "0.00");
				setValue("txtTdsper", attr(entry, "TdsPercentage") || "0.00");
				setCRDR(attr(entry, "CRDR"));
				setValue("txtNarration", entryNarration(entry));
				childElements(entry).forEach(function (child) {
					if (child.nodeName === "AccHead") {
						setSelectedHeadDisplay(attr(child, "Name"));
						if (typeof window.SelectHead === "function") {
							window.SelectHead(attr(child, "No"), attr(child, "Type"), field("selAccHead"), Number(valueOf("hHeadCount", 0)));
						}
					}
					if (child.nodeName === "CostCenter" && typeof window.popCostCenter === "function") {
						window.setADDDisplay(1);
						window.popCostCenter(child);
					}
					if (child.nodeName === "Analytical" && typeof window.popAnalytical === "function") {
						window.setADDDisplay(1);
						window.popAnalytical(child);
					}
					if (child.nodeName === "PayRec" && typeof window.popPayRec === "function") {
						window.popPayRec(child);
					}
				});
				if (editType === "D") {
					window.DelEntry();
				}
				window.DisplayVoucher("0");
				return true;
			});
		};

		window.DelEntry = function () {
			resetEntryForm();
			window.bVouFlag = false;
			window.bEditFlag = true;
			window.bSavFlag = true;
			window.DisplayVoucher("0");
		};

		window.DisplayVoucher = function (displayType) {
			var state = initState();
			var table = clearVoucherTable();
			var total = 0;
			var creditTotal = 0;
			var debitTotal = 0;
			var row;
			var entries = childElements(state.vouRoot, "Entry");
			if (!table) {
				return;
			}
			makeDisplayVisible();
			entries.forEach(function (entry, index) {
				var no = index + 1;
				var crdr = attr(entry, "CRDR");
				var amount = toNumber(attr(entry, "Amount"));
				setAttr(entry, "No", no);
				row = table.insertRow(table.rows.length);
				insertCell(row, 1, "", no, "ExcelSerial", "Center", "top");
				if (String(displayType) === "0") {
					insertCell(row, 1, "", "<a href=\"#\" onclick=\"EditEntry('" + no + "','D'); return false;\">Delete</a>", "ExcelDisplayCell", "Center", "top");
					insertCell(row, 1, "", "<a href=\"#\" onclick=\"EditEntry('" + no + "','E'); return false;\"><b>Edit</b></a>", "ExcelDisplayCell", "Center", "top");
				} else {
					insertCell(row, 1, "", "", "ExcelDisplayCell", "Center", "top", 0, 0, 2);
				}
				insertCell(row, 1, "", entryAccountText(entry), "ExcelDisplayCell", "left", "top");
				if (!config.journal) {
					insertCell(row, 1, "", entryAdditionalText(entry), "ExcelDisplayCell", "left", "top");
					insertCell(row, 1, "", entryNarration(entry), "ExcelDisplayCell", "left", "top");
					insertCell(row, 1, "", formatNumber(amount, 2) + "&nbsp;" + crdr, "ExcelDisplayCell", "right", "top");
					insertCell(row, 1, "", formatNumber(attr(entry, "TdsAmount"), 2), "ExcelDisplayCell", "right", "top");
					insertCell(row, 1, "", formatNumber(attr(entry, "TdsPercentage"), 2), "ExcelDisplayCell", "right", "top");
				} else {
					insertCell(row, 1, "", entryNarration(entry), "ExcelDisplayCell", "left", "top");
					insertCell(row, 1, "", crdr === "C" ? formatNumber(amount, 2) : "", "ExcelDisplayCell", "right", "top");
					insertCell(row, 1, "", crdr === "D" ? formatNumber(amount, 2) : "", "ExcelDisplayCell", "right", "top");
					insertCell(row, 1, "", entryAdditionalText(entry), "ExcelDisplayCell", "left", "top");
				}
				if (crdr === "C") {
					creditTotal += amount;
					total -= amount;
				} else {
					debitTotal += amount;
					total += amount;
				}
			});
			row = table.insertRow(table.rows.length);
			if (config.journal) {
				insertCell(row, 1, "", "<b>Total</b>", "ExcelDisplayCell", "right", "top", 0, 0, 4);
				insertCell(row, 1, "", formatNumber(creditTotal, 2), "ExcelDisplayCell", "right", "top");
				insertCell(row, 1, "", formatNumber(debitTotal, 2), "ExcelDisplayCell", "right", "top");
			} else {
				insertCell(row, 1, "", "<b>Total</b>", "ExcelDisplayCell", "right", "top", 0, 0, 6);
				insertCell(row, 1, "", "<input type=\"text\" name=\"txtTotalAmt\" value=\"" + formatNumber(total, 2) + "\" size=\"13\" class=\"Formelemread\" style=\"text-align:right\">", "ExcelDisplayCell", "right", "top");
			}
			setText("spEntryNo", String(entries.length + 1));
			window.iEntryNo = entries.length;
			setValue("hEntryNo", String(entries.length + 1));
			setSelectedHeadDisplay("");
		};

		window.CheckAmount = function () {
			var root = initState().vouRoot;
			var credit = 0;
			var debit = 0;
			childElements(root, "Entry").forEach(function (entry) {
				if (attr(entry, "CRDR") === "C") {
					credit += toNumber(attr(entry, "Amount"));
				} else {
					debit += toNumber(attr(entry, "Amount"));
				}
			});
			if (formatNumber(credit, 2) !== formatNumber(debit, 2)) {
				alert("Cr Total should be equal to Dr Total");
				return false;
			}
			return true;
		};

		window.CheckVouAmount = function () {
			var root = initState().vouRoot;
			var credit = 0;
			var debit = 0;
			childElements(root, "Entry").forEach(function (entry) {
				if (attr(entry, "CRDR") === "C") {
					credit += toNumber(attr(entry, "Amount")) - toNumber(attr(entry, "TdsAmount"));
				} else {
					debit += toNumber(attr(entry, "Amount")) - toNumber(attr(entry, "TdsAmount"));
				}
			});
			return valueOf("hVouCRDR", checkedCRDR()) === "D" ? formatNumber(credit - debit, 2) : formatNumber(debit - credit, 2);
		};

		window.CheckVouStat = function () {
			return true;
		};

		window.CheckAdjVal = function () {
			return true;
		};

		window.CheckContraEnt = function () {
			var root = initState().vouRoot;
			var entries = childElements(root, "Entry");
			var bookAcc = selectedBookAccHead();
			var orgId = currentOrgId();
			var blocked = false;
			if (entries.length <= 1) {
				return true;
			}
			selectNodes(root, '//AccHead[@Type="G"]').forEach(function (node) {
				if (blocked) {
					return;
				}
				var xhr = getText("XMLContraEntAccChk.asp?BkAccHd=" + encodeURIComponent(bookAcc) + "&orgID=" + encodeURIComponent(orgId) + "&AccHead=" + encodeURIComponent(attr(node, "No")));
				if (trim(xhr.responseText) !== "0") {
					alert("Contra Entry is Created only One Entry is allowed ");
					blocked = true;
				}
			});
			return !blocked;
		};

		window.UpdateXML = function () {
			var root = initState().vouRoot;
			var select = field("selBook");
			var parts = currentBookParts();
			if (!root || !select || select.selectedIndex === 0) {
				return;
			}
			setAttr(root, "UnitNo", currentOrgId());
			setAttr(root, "UnitName", valueOf("hOrgName", ""));
			setAttr(root, "BookNo", config.journal ? selectedValue(select) : (parts[0] || selectedValue(select)));
			setAttr(root, "BookName", selectedText(select));
			setAttr(root, "CRDR", valueOf("hVouCRDR", checkedCRDR()));
			setAttr(root, "VouDate", currentDateValue("ctlDate"));
			setAttr(root, "BookAcchead", parts[1] || valueOf("hBookAccHead", selectedBookAccHead()));
		};

		window.SetBookAccHead = function () {
			var root = xmlRoot("UnitBookData");
			var parts = currentBookParts();
			if (parts[1]) {
				setValue("hBookAccHead", parts[1]);
			}
			childElements(root).some(function (node) {
				if (attr(node, 0) === parts[0]) {
					setValue("hBookAccHead", attr(node, 2));
					setValue("hBookOtherUnit", attr(node, 3));
					return true;
				}
				return false;
			});
			window.DisplayBalamt();
		};

		window.SelUnBook = window.SetUnBook = function () {
			window.DisplayBook();
			firstSelectedBook();
			if (valueOf("hCallFrm", "") === "A") {
				setValue("hAmendTy", "A");
				window.MakeDispVou("C");
			}
		};

		window.MakeDispVou = function (type) {
			var transNo = valueOf("hTransNo", "");
			var xhr;
			var root;
			if (trim(transNo) === "" || trim(transNo) === "0") {
				return;
			}
			xhr = getText("XMLGetVoucher.asp?TransNo=" + encodeURIComponent(transNo));
			if (!xhr.responseXML || !xhr.responseXML.documentElement) {
				return;
			}
			xmlObject("VoucherData").loadXML(xmlString(xhr.responseXML));
			root = xmlRoot("VoucherData");
			if (root) {
				selectBookByValue(attr(root, "BookNo") + (attr(root, "BookAcchead") ? "-" + attr(root, "BookAcchead") : ""));
			}
			window.DisplayVoucher(type === "C" ? "0" : "1");
		};

		window.SaveXML = doSaveXml;

		window.DelVouch = function () {
			if (trim(valueOf("txtVouNo", "")) === "") {
				alert("Select Voucher ");
				return;
			}
			form().action = "VouDeletion.asp";
			form().submit();
		};

		window.CancelAction = function (page) {
			form().action = page;
			form().submit();
		};

		window.DisplayPayRec = function () {
			var entry = childElements(initState().vouRoot, "Entry")[0];
			if (entry) {
				setValue("txtPayTo", attr(entry, "Payto"));
				if (field("txtPayTo")) {
					field("txtPayTo").readOnly = true;
				}
			}
		};

		window.CheckEntryType = function (select) {
			if (select.selectedIndex === 0) {
				if (field("btnNext")) {
					field("btnNext").disabled = true;
				}
				return;
			}
			if (field("btnNext")) {
				field("btnNext").disabled = false;
			}
			setValue("hVouCRDR", select.value);
			setCRDR(select.value === "D" ? "D" : "C");
		};

		window.DisplayBalamt = function () {
			var parts = currentBookParts();
			var xhr;
			var values;
			if (!parts[1]) {
				return;
			}
			xhr = getText("GetDayOpenByDate.asp?sValue=" + encodeURIComponent(currentOrgId() + ":" + parts[1] + ":" + currentDateValue("ctlDate")));
			values = String(xhr.responseText || "").split("*");
			if (values.length === 2) {
				setText("spBookBal", formatNumber(Math.abs(toNumber(values[0])), 2) + (toNumber(values[0]) >= 0 ? " Dr " : " Cr "));
				setText("spCurrBal", formatNumber(Math.abs(toNumber(values[1])), 2) + (toNumber(values[1]) >= 0 ? " Dr " : " Cr "));
			}
		};

		window.PopInsDet = function () {
			var bookParts = currentBookParts();
			var tempValues = [valueOf("hVouCRDR", checkedCRDR()), currentDateValue("ctlDate"), valueOf("hVouCode", config.moduleCode), currentOrgId(), valueOf("hTransNo", ""), valueOf("hVouName", "")].join(":");
			openDialog("BankInsDetails.asp?sTemp=" + encodeURIComponent(tempValues), xmlObject("VoucherData"), "dialogHeight:350px;dialogWidth:710px;center:Yes;help:No;resizable:No;status:No", function (node) {
				var total = 0;
				var details = [];
				selectNodes(node, "//BankInstrumentDet").forEach(function (ins) {
					var type = String(attr(ins, 2)).charAt(0).toUpperCase();
					var label = type === "C" ? "CH NO: " : type === "D" ? "DD NO: " : type === "B" ? "BANK CH: " : type === "T" ? "TT NO: " : "Cash: ";
					details.push(label + attr(ins, 1) + " - " + attr(ins, 3));
					total += toNumber(attr(ins, 6));
				});
				setValue("hInsDet", details.length ? details[0].split(":")[0] + ": " : "");
				setValue("hInsrAmt", formatNumber(total, 2));
				setText("spInsDet", details.join(","));
			});
			return bookParts;
		};

		window.PopCCAH = function () {
			var entryRoot = initState().entryRoot;
			var account = childElements(entryRoot, "AccHead")[0];
			if (!account) {
				alert("Select Account Head");
				return;
			}
			window.showCCAnal(currentAccountingUnit().id, attr(account, "No"), attr(account, "CostCenter"), attr(account, "Analytical"));
		};

		window.TDSAmount = function () {
			setValue("hTdsNew", "Y");
		};
		window.TDSChngAmt = function () {
			setValue("hTdsNew", "Y");
		};
		window.TDSCalc = function () {
			setValue("hTdsNew", "Y");
		};
		window.AmtFun = function () {
			if (typeof window.popAddAmount === "function") {
				window.popAddAmount();
			}
		};
		window.PrnVouch = function () {
			var transNo = valueOf("hTransNo", "");
			if (trim(transNo)) {
				window.open("VouPrint.asp?TransNo=" + encodeURIComponent(transNo), "_blank");
			}
		};

		window.AddNewParty = function () {
			openDialog("MisParCreate.asp", "", "dialogHeight:495px;dialogWidth:420px;center:Yes;help:No;resizable:No;status:No", function (value) {
				setValue("txtPayTo", value || "");
			});
		};

		window.SelMisParty = function () {
			var size = popupSize("4", "MisPartySelection.asp", "500", "350");
			function handle(value) {
				var parts = String(value || "").split(":");
				if (String(value) === "AN") {
					window.AddNewParty();
					return;
				}
				if (parts.length <= 1 && String(value || "") !== "") {
					openDialog("../../Common/" + size.program + "?" + value, "", "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", handle);
					return;
				}
				if (parts.length > 1) {
					setValue("txtPayTo", parts[0]);
				}
			}
			openDialog("../../Common/" + size.program + "?orgID=" + encodeURIComponent(currentAccountingUnit().id), "", "dialogHeight:" + size.height + "px;dialogWidth:" + size.width + "px;Status:No", handle);
		};

		window.popVoucSel = function () {};
		window.popMonBalance = function (value) {
			var parts = String(value || "").split("~");
			var date = currentDateValue("ctlDate");
			var till = String(date).slice(-4) + String(date).substr(3, 2);
			openDialog("PopMonBalance.asp?orgid=" + encodeURIComponent(parts[0] || "") + "&Acchead=" + encodeURIComponent(parts[1] || "") + "&TillDate=" + encodeURIComponent(till), "", "dialogHeight:390px;dialogWidth:450px;center:Yes;help:No;resizable:No;status:No");
		};
		window.popDayBalance = function (value) {
			var parts = String(value || "").split("~");
			openDialog("PopDayBalance.asp?orgid=" + encodeURIComponent(parts[0] || "") + "&Acchead=" + encodeURIComponent(parts[1] || "") + "&TillDate=" + encodeURIComponent(currentDateValue("ctlDate")), "", "dialogHeight:390px;dialogWidth:620px;center:Yes;help:No;resizable:No;status:No");
		};

		window.Init = function () {
			initState();
		};
		window.setdate = function () {
			var today = valueOf("hCurrDate", "");
			if (today) {
				setCurrentDate("ctlDate", today);
			}
		};
	};

	window.ITMSVoucherEntry = {
		install: install
	};
})(window, document);
